
# -------------------------------------------
# Sage 300 Setup Script - SQL Express + IIS + SSL + ODBC
# -------------------------------------------

# ---------------------------
# Setup Logging
# ---------------------------
$logFile = "$PSScriptRoot\setup_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
function Log-Message {
    param ([string]$message, [string]$level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp][$level] $message"
    Add-Content -Path $logFile -Value $entry
    Write-Output $entry
}

# ---------------------------
# Prerequisites: Check for sqlcmd
# ---------------------------
if (-not (Get-Command sqlcmd.exe -ErrorAction SilentlyContinue)) {
    Log-Message "sqlcmd not found. Please install SQL Server Command Line Utilities." "ERROR"
    exit 1
}

# ---------------------------
# Detect SQL Server Instances
# ---------------------------
Log-Message "Detecting local SQL Server instances..."
$sqlInstancesRaw = & sqlcmd -L 2>&1 | Select-String '\\' | ForEach-Object {
    ($_ -split '\\')[1].Trim()
} | Sort-Object -Unique

if (!$sqlInstancesRaw -or $sqlInstancesRaw.Count -eq 0) {
    Log-Message "No SQL Server instances found. Please make sure SQL Server is installed." "ERROR"
    exit 1
}

Log-Message "Available SQL Server Instances:"
for ($i = 0; $i -lt $sqlInstancesRaw.Count; $i++) {
    Log-Message "[$($i+1)] $($sqlInstancesRaw[$i])"
}

$selection = Read-Host "Enter the number of the SQL Server instance you want to use"
$selectedIndex = [int]$selection - 1
if ($selectedIndex -lt 0 -or $selectedIndex -ge $sqlInstancesRaw.Count) {
    Log-Message "Invalid SQL Server selection." "ERROR"
    exit 1
}
$sqlInstance = $sqlInstancesRaw[$selectedIndex]
$sqlServerName = "localhost\$sqlInstance"
Log-Message "Selected SQL Instance: $sqlServerName"

# ---------------------------
# Prompt for SQL Login Password
# ---------------------------
$securePassword = Read-Host -Prompt "Enter the password you want to use for the Sage 300 SQL login (e.g., 'sage300')" -AsSecureString
$passwordPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
)

# ---------------------------
# Ensure SqlServer Module is Installed
# ---------------------------
if (-not (Get-Module -ListAvailable -Name SqlServer)) {
    Log-Message "Installing SqlServer PowerShell module..."
    Install-PackageProvider -Name NuGet -Force -Scope CurrentUser
    Install-Module -Name SqlServer -Force -Scope CurrentUser -AllowClobber
}
Import-Module SqlServer
Log-Message "SqlServer PowerShell module loaded."

# ---------------------------
# Enable TCP/IP and Set Port
# ---------------------------
$tcpRegPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL*\MSSQLServer\SuperSocketNetLib\Tcp\IPAll"
$regPaths = Get-Item -Path $tcpRegPath -ErrorAction SilentlyContinue
if ($regPaths) {
    foreach ($path in $regPaths) {
        try {
            Set-ItemProperty -Path $path.PSPath -Name "TcpDynamicPorts" -Value ""
            Set-ItemProperty -Path $path.PSPath -Name "TcpPort" -Value "1433"
            Log-Message "TCP/IP set to port 1433 for $($path.PSPath)"
        } catch {
            Log-Message "Failed to update TCP/IP for $($path.PSPath)" "WARNING"
        }
    }
} else {
    Log-Message "Could not detect registry path for TCP/IP config." "WARNING"
}

Restart-Service -Name ("MSSQL$" + $sqlInstance)
Log-Message "SQL Service restarted."

# ---------------------------
# Allow Port 1433 in Firewall
# ---------------------------
Log-Message "Creating Firewall rule for port 1433..."
New-NetFirewallRule -DisplayName "SQL Server Port 1433" -Direction Inbound -Protocol TCP -LocalPort 1433 -Action Allow

# ---------------------------
# Enable Mixed Mode Authentication
# ---------------------------
Log-Message "Enabling Mixed Mode Authentication..."
$connectionString = "Server=$sqlServerName;Database=master;Integrated Security=True;"
Invoke-Sqlcmd -Query "EXEC xp_instance_regwrite N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer', N'LoginMode', REG_DWORD, 2" -ConnectionString $connectionString
Restart-Service -Name ("MSSQL$" + $sqlInstance)

# ---------------------------
# Create SQL Login
# ---------------------------
Log-Message "Creating SQL Login for Sage 300..."
$loginQuery = @"
IF NOT EXISTS (SELECT * FROM sys.sql_logins WHERE name = 'sage300')
BEGIN
    CREATE LOGIN [sage300] WITH PASSWORD = N'$passwordPlain', CHECK_POLICY = OFF;
    EXEC sp_addsrvrolemember 'sage300', 'sysadmin';
END
"@
Invoke-Sqlcmd -Query $loginQuery -ConnectionString $connectionString

# ---------------------------
# Create Databases and Grant Access
# ---------------------------
$createDatabases = @"
IF DB_ID('VAULT') IS NULL CREATE DATABASE [VAULT] COLLATE Latin1_General_BIN;
IF DB_ID('STORE') IS NULL CREATE DATABASE [STORE] COLLATE Latin1_General_BIN;
IF DB_ID('SYSCMP') IS NULL CREATE DATABASE [SYSCMP] COLLATE Latin1_General_BIN;
IF DB_ID('COMP01') IS NULL CREATE DATABASE [COMP01] COLLATE Latin1_General_BIN;
IF DB_ID('PORTAL') IS NULL CREATE DATABASE [PORTAL] COLLATE Latin1_General_BIN;
"@
Invoke-Sqlcmd -Query $createDatabases -ConnectionString $connectionString

$grantRights = @"
USE [VAULT]; CREATE USER [sage300] FOR LOGIN [sage300]; ALTER ROLE db_owner ADD MEMBER [sage300];
USE [STORE]; CREATE USER [sage300] FOR LOGIN [sage300]; ALTER ROLE db_owner ADD MEMBER [sage300];
USE [SYSCMP]; CREATE USER [sage300] FOR LOGIN [sage300]; ALTER ROLE db_owner ADD MEMBER [sage300];
USE [COMP01]; CREATE USER [sage300] FOR LOGIN [sage300]; ALTER ROLE db_owner ADD MEMBER [sage300];
USE [PORTAL]; CREATE USER [sage300] FOR LOGIN [sage300]; ALTER ROLE db_owner ADD MEMBER [sage300];
"@
Invoke-Sqlcmd -Query $grantRights -ConnectionString $connectionString

# ---------------------------
# Install IIS + SSL
# ---------------------------
Log-Message "Installing IIS features..."
$features = @(
  "IIS-WebServerRole", "IIS-WebServer", "IIS-CommonHttpFeatures", "IIS-StaticContent",
  "IIS-DefaultDocument", "IIS-DirectoryBrowsing", "IIS-HttpErrors", "IIS-ApplicationDevelopment",
  "IIS-ASPNET", "IIS-ASPNET45", "IIS-NetFxExtensibility45", "IIS-ISAPIExtensions",
  "IIS-ISAPIFilter", "IIS-CGI", "IIS-ManagementConsole", "IIS-Security"
)
foreach ($feature in $features) {
    dism.exe /Online /Enable-Feature /FeatureName:$feature /All /NoRestart | Out-Null
    Log-Message "Enabled IIS feature: $feature"
}

Log-Message "Creating self-signed SSL certificate..."
$cert = New-SelfSignedCertificate -DnsName "localhost" -CertStoreLocation "cert:\LocalMachine\My"
$thumbprint = $cert.Thumbprint
$guid = [guid]::NewGuid().ToString("B")

Import-Module WebAdministration
if (-not (Get-WebBinding -Name "Default Web Site" -Protocol "https" -ErrorAction SilentlyContinue)) {
    New-WebBinding -Name "Default Web Site" -Protocol https -Port 443
}
netsh http add sslcert ipport=0.0.0.0:443 certhash=$thumbprint appid=$guid | Out-Null
iisreset | Out-Null
Log-Message "SSL certificate bound to port 443."

# ---------------------------
# Create ODBC DSNs
# ---------------------------
Log-Message "Creating DSNs..."
$driverName = "ODBC Driver 17 for SQL Server"
if ($driverName -notin (Get-OdbcDriver | Select-Object -ExpandProperty Name)) {
    Log-Message "'$driverName' not installed. Please install it first." "ERROR"
    exit 1
}

$dsnList = @(
    @{ Name = "VAULT_DSN";  DB = "VAULT"  },
    @{ Name = "STORE_DSN";  DB = "STORE"  },
    @{ Name = "SYSCMP_DSN"; DB = "SYSCMP" },
    @{ Name = "COMP01_DSN"; DB = "COMP01" },
    @{ Name = "PORTAL_DSN"; DB = "PORTAL" }
)
foreach ($dsn in $dsnList) {
    try {
        Add-OdbcDsn -Name $dsn.Name -DsnType "System" -Platform "64-bit" `
            -Driver $driverName `
            -SetPropertyValue @(
                "Server=$sqlServerName",
                "Database=$($dsn.DB)",
                "UID=sage300",
                "PWD=$passwordPlain",
                "Trusted_Connection=No"
            )
        Log-Message "DSN $($dsn.Name) created."
    } catch {
        Log-Message "Failed to create DSN $($dsn.Name): $_" "WARNING"
    }
}

Log-Message "Sage 300 SQL + IIS + SSL setup complete."
