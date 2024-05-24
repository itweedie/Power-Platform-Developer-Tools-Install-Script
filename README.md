# Power Pages Developer Tools Install Script

This repository contains a PowerShell script that automates the installation of essential developer tools for Power Pages development. Using `winget`, it installs various software packages for a quick and efficient setup. It also downloads, extracts, and sets up the latest release of XrmToolBox.

## Features

- **Automated Installation**: Installs popular development tools including Visual Studio Code, Power BI Desktop, SQL Server Management Studio, Windows Terminal, Notepad++, PowerToys, Postman, and Visual Studio 2022 Professional.
- **PowerShell Modules**: Installs necessary PowerShell modules for PowerApps Administration.
- **Git Installation**: Ensures Git is installed for version control.
- **Azure Storage Explorer**: Installs Azure Storage Explorer for managing Azure storage resources.
- **ArchiMate Tool**: Downloads and installs Archi, a popular ArchiMate modeling tool.
- **XrmToolBox Setup**: Downloads the latest release of XrmToolBox, extracts it to the user's Desktop, and organizes it in a specified folder.

## Prerequisites

- Windows operating system
- PowerShell 5.1 or higher
- Internet connection

## How to Use

1. Open PowerShell with administrative privileges.
2. Run the following command to execute the script directly from GitHub:

   ```powershell
      # Ensure the Desktop directory exists
      $desktopPath = [System.Environment]::GetFolderPath('Desktop')
      if (-Not (Test-Path -Path $desktopPath)) {
          New-Item -ItemType Directory -Path $desktopPath | Out-Null
      }
      
      # Download and run the script from GitHub
      Invoke-WebRequest -Uri "https://raw.githubusercontent.com/itweedie/Power-Pages-Developer-Tools-Install-Script/main/install-power-platform-dev-tools.ps1" -OutFile "$desktopPath\install-power-platform-dev-tools.ps1"
      PowerShell -ExecutionPolicy Bypass -File "$desktopPath\install-power-platform-dev-tools.ps1"
   ```

3. Follow any on-screen prompts to complete the installations.

## Script Details

The script performs the following actions:

1. Installs development tools using `winget`.
2. Installs PowerShell modules for PowerApps.
3. Installs Git and Azure Storage Explorer.
4. Downloads and installs the ArchiMate tool.
5. Downloads the latest release of XrmToolBox, extracts it to the user's Desktop, and renames the folder to "XrmToolBox".

## Contributions

Contributions are welcome! Please feel free to submit issues or pull requests to enhance the functionality or add new features to the script.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
