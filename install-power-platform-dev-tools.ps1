# Array to store errors
$errors = @()

# Function to execute winget commands and capture errors
function Execute-WingetCommand {
    param (
        [string]$command
    )

    try {
        Invoke-Expression $command
    }
    catch {
        $errors += "Error executing: $command`n$_"
    }
}

# Function to install VSCode extensions and capture errors
function Install-VSCodeExtension {
    param (
        [string]$extension
    )

    try {
        code --install-extension $extension --force
    }
    catch {
        $errors += "Error installing VSCode extension: $extension`n$_"
    }
}

# Install useful packages from the Windows Store using winget

# Install Visual Studio Code
Execute-WingetCommand "winget install --id=Microsoft.VisualStudioCode -e --accept-package-agreements"

# Install Power BI Desktop
Execute-WingetCommand "winget install --id=Microsoft.PowerBIDesktop -e --accept-package-agreements"

# Install SQL Server Management Studio
Execute-WingetCommand "winget install --id=Microsoft.SQLServerManagementStudio -e --accept-package-agreements"

# Install Windows Terminal
Execute-WingetCommand "winget install --id=Microsoft.WindowsTerminal -e --accept-package-agreements"

# Install Notepad++
Execute-WingetCommand "winget install --id=Notepad++.Notepad++ -e --accept-package-agreements"

# Install PowerToys
Execute-WingetCommand "winget install --id=Microsoft.PowerToys -e --accept-package-agreements"

# Install Postman
Execute-WingetCommand "winget install --id=Postman.Postman -e --accept-package-agreements"

# Install Visual Studio 2022 Professional
Execute-WingetCommand "winget install --id=Microsoft.VisualStudio.2022.Professional -e --accept-package-agreements"

# Install Node.js (includes NPM)
Execute-WingetCommand "winget install --id=OpenJS.NodeJS -e --accept-package-agreements"

# Install PowerShell modules for PowerApps Administration
try {
    Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser
}
catch {
    $errors += "Error installing Microsoft.PowerApps.Administration.PowerShell`n$_"
}

try {
    Install-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber -Scope CurrentUser
}
catch {
    $errors += "Error installing Microsoft.PowerApps.PowerShell`n$_"
}

# Install Git on the Server
Execute-WingetCommand "winget install --id=Git.Git -e --accept-package-agreements"

# Install Azure Storage Explorer
Execute-WingetCommand "winget install --id=Microsoft.AzureStorageExplorer -e --accept-package-agreements"

# Download and install Archi (ArchiMate tool)
Execute-WingetCommand "winget install --exact --id=Archi -e --source winget"

# Download, extract, and set up XrmToolBox

# Define paths
$desktopPath = [System.Environment]::GetFolderPath('Desktop')
$zipFilePath = "$desktopPath\latest_release.zip"
$destinationFolder = "$desktopPath\XrmToolBox"

# Fetch the latest release information from GitHub
try {
    $latestRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/MscrmTools/XrmToolBox/releases/latest" -UseBasicParsing
    # Extract the download URL for the first asset
    $latestAssetUrl = $latestRelease.assets[0].browser_download_url
}
catch {
    $errors += "Error fetching latest release information from GitHub`n$_"
}

# Download the asset (zip file) to the desktop
try {
    Invoke-WebRequest -Uri $latestAssetUrl -OutFile $zipFilePath
}
catch {
    $errors += "Error downloading the asset (zip file) from GitHub`n$_"
}

# Create the destination folder if it doesn't exist
if (-Not (Test-Path -Path $destinationFolder)) {
    try {
        New-Item -ItemType Directory -Path $destinationFolder | Out-Null
    }
    catch {
        $errors += "Error creating the destination folder: $destinationFolder`n$_"
    }
}

# Unzip the downloaded file to the destination folder
try {
    Expand-Archive -Path $zipFilePath -DestinationPath $destinationFolder -Force
}
catch {
    $errors += "Error unzipping the downloaded file to the destination folder`n$_"
}

# Delete the zip file after extraction
try {
    Remove-Item -Path $zipFilePath
}
catch {
    $errors += "Error deleting the zip file: $zipFilePath`n$_"
}

# Output the path to the extracted folder
Write-Output "XrmToolBox has been downloaded and extracted to: $destinationFolder"

# Configure Git with user name and email from Windows
try {
    $userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $userEmail = "$userName@example.com"
    
    git config --global user.name "$userName"
    git config --global user.email "$userEmail"
}
catch {
    $errors += "Error configuring Git with user name and email`n$_"
}

# Install VSCode extensions
Install-VSCodeExtension "ms-azure-devops.azure-pipelines"
Install-VSCodeExtension "ms-azuretools.vscode-azureresourcegroups"
Install-VSCodeExtension "ms-dataverse.dataverse-devtools"
Install-VSCodeExtension "ms-vscode.azure-account"
Install-VSCodeExtension "microsoft-Isvexptools.powerplatform-vscode"
Install-VSCodeExtension "ms-vscode.powershell"
Install-VSCodeExtension "github.remotehub"

# Output any errors
if ($errors.Count -gt 0) {
    Write-Output "`nErrors encountered during the execution of the script:`n"
    $errors | ForEach-Object { Write-Output $_ }
} else {
    Write-Output "`nNo errors encountered during the execution of the script."
}
