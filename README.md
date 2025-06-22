# Sage 300 SQL Express Setup Script

Automate the setup of Microsoft SQL Server Express, IIS, SSL, and ODBC Data Sources required for **Sage 300**. This PowerShell script streamlines the provisioning of a local environment for Sage 300 ERP systems by handling configuration tasks including authentication, database creation, and firewall setup.

## 🚀 Features

- Detects installed SQL Server Express instances
- Enables TCP/IP on port 1433 and allows it through the firewall
- Installs the SqlServer module if missing
- Enables mixed mode authentication
- Creates SQL login (`sage300`) with sysadmin rights
- Creates Sage 300-related databases with the required collation
- Grants proper user permissions to each database
- Installs IIS with required features and a self-signed certificate for SSL
- Binds SSL cert to port 443
- Configures ODBC DSNs using ODBC Driver 17 for SQL Server
- Modular and thoroughly logged

## 📦 Requirements

- Windows OS with administrative privileges
- SQL Server Express installed
- `sqlcmd` utility available in the system path
- ODBC Driver 17 for SQL Server
- PowerShell 5.1+ with script execution permissions
- Internet access (to install PowerShell modules and features if needed)

## 🔧 Usage

1. Open PowerShell as Administrator.
2. Execute the script:

```powershell
.\Sage300_SQL_Setup.ps1
