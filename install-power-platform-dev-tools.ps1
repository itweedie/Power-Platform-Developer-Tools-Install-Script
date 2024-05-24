# Install useful packages from the Windows Store using winget

# Install Visual Studio Code
winget install --id=Microsoft.VisualStudioCode -e --accept-package-agreements

# Install Power BI Desktop
winget install --id=Microsoft.PowerBIDesktop -e --accept-package-agreements

# Install SQL Server Management Studio
winget install --id=Microsoft.SQLServerManagementStudio -e --accept-package-agreements

# Install Windows Terminal
winget install --id=Microsoft.WindowsTerminal -e --accept-package-agreements

# Install Notepad++
winget install --id=Notepad++.Notepad++ -e --accept-package-agreements

# Install PowerToys
winget install --id=Microsoft.PowerToys -e --accept-package-agreements

# Install Postman
winget install --id=Postman.Postman -e --accept-package-agreements

# Install Visual Studio 2022 Professional
winget install --id=Microsoft.VisualStudio.2022.Professional -e --accept-package-agreements

# Install PowerShell modules for PowerApps Administration
Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser
Install-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber -Scope CurrentUser

# Install Git on the Server
winget install --id=Git.Git -e --accept-package-agreements

# Install Azure Storage Explorer
winget install --id=Microsoft.AzureStorageExplorer -e --accept-package-agreements

# Download and install Archi (ArchiMate tool)
winget install --exact --id=Archi -e --source winget

# Download, extract, and set up XrmToolBox

# Define paths
$desktopPath = [System.Environment]::GetFolderPath('Desktop')
$zipFilePath = "$desktopPath\latest_release.zip"
$destinationFolder = "$desktopPath\XrmToolBox"

# Fetch the latest release information from GitHub
$latestRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/MscrmTools/XrmToolBox/releases/latest" -UseBasicParsing

# Extract the download URL for the first asset
$latestAssetUrl = $latestRelease.assets[0].browser_download_url

# Download the asset (zip file) to the desktop
Invoke-WebRequest -Uri $latestAssetUrl -OutFile $zipFilePath

# Create the destination folder if it doesn't exist
if (-Not (Test-Path -Path $destinationFolder)) {
    New-Item -ItemType Directory -Path $destinationFolder | Out-Null
}

# Unzip the downloaded file to the destination folder
Expand-Archive -Path $zipFilePath -DestinationPath $destinationFolder

# Delete the zip file after extraction
Remove-Item -Path $zipFilePath

# Output the path to the extracted folder
Write-Output "XrmToolBox has been downloaded and extracted to: $destinationFolder"

