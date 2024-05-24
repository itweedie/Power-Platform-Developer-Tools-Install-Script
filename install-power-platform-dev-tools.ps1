# Array to store errors
$errors = @()

# Function to execute winget commands and capture errors
function Execute-WingetCommand {
    param (
        [string]$command,
        [string]$description
    )

    Write-Host "Installing $description..."
    try {
        Invoke-Expression $command
        Write-Host "$description installed successfully."
    }
    catch {
        $errorMessage = "Error executing: $command`n$_"
        $errors += $errorMessage
        Write-Host $errorMessage -ForegroundColor Red
    }
}

# Function to install VSCode extensions and capture errors
function Install-VSCodeExtension {
    param (
        [string]$extension,
        [string]$description
    )

    Write-Host "Installing VSCode extension: $description..."
    try {
        # Set environment variable to bypass SSL certificate verification
        $env:NODE_TLS_REJECT_UNAUTHORIZED = "0"
        $output = code --install-extension $extension --force 2>&1
        # Remove the environment variable after installation
        Remove-Item Env:NODE_TLS_REJECT_UNAUTHORIZED

        if ($LASTEXITCODE -ne 0) {
            throw "Error installing VSCode extension: $extension`n$output"
        }

        Write-Host "VSCode extension $description installed successfully."
    }
    catch {
        $errorMessage = "Error installing VSCode extension: $extension`n$_"
        $errors += $errorMessage
        Write-Host $errorMessage -ForegroundColor Red
    }
}

# Configure npm to bypass strict SSL certificate checking
Write-Host "Configuring npm to bypass strict SSL certificate checking..."
try {
    npm config set strict-ssl false -g
    Write-Host "npm configured to bypass strict SSL certificate checking."
}
catch {
    $errorMessage = "Error configuring npm to bypass strict SSL certificate checking`n$_"
    $errors += $errorMessage
    Write-Host $errorMessage -ForegroundColor Red
}

# Install useful packages from the Windows Store using winget

Execute-WingetCommand "winget install --id=Microsoft.VisualStudioCode -e --accept-package-agreements" "Visual Studio Code"
Execute-WingetCommand "winget install --id=Microsoft.PowerBIDesktop -e --accept-package-agreements" "Power BI Desktop"
Execute-WingetCommand "winget install --id=Microsoft.SQLServerManagementStudio -e --accept-package-agreements" "SQL Server Management Studio"
Execute-WingetCommand "winget install --id=Microsoft.WindowsTerminal -e --accept-package-agreements" "Windows Terminal"
Execute-WingetCommand "winget install --id=Notepad++.Notepad++ -e --accept-package-agreements" "Notepad++"
Execute-WingetCommand "winget install --id=Microsoft.PowerToys -e --accept-package-agreements" "PowerToys"
Execute-WingetCommand "winget install --id=Postman.Postman -e --accept-package-agreements" "Postman"
Execute-WingetCommand "winget install --id=Microsoft.VisualStudio.2022.Professional -e --accept-package-agreements" "Visual Studio 2022 Professional"
Execute-WingetCommand "winget install --id=OpenJS.NodeJS -e --accept-package-agreements" "Node.js (includes NPM)"
Execute-WingetCommand "winget install --id=Git.Git -e --accept-package-agreements" "Git"
Execute-WingetCommand "winget install --id=Microsoft.AzureStorageExplorer -e --accept-package-agreements" "Azure Storage Explorer"

# Install PowerShell modules for PowerApps Administration
Write-Host "Installing PowerShell modules for PowerApps Administration..."
try {
    Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser -Force -AllowClobber
    Write-Host "Microsoft.PowerApps.Administration.PowerShell installed successfully."
}
catch {
    $errorMessage = "Error installing Microsoft.PowerApps.Administration.PowerShell`n$_"
    $errors += $errorMessage
    Write-Host $errorMessage -ForegroundColor Red
}

try {
    Install-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber -Scope CurrentUser -Force
    Write-Host "Microsoft.PowerApps.PowerShell installed successfully."
}
catch {
    $errorMessage = "Error installing Microsoft.PowerApps.PowerShell`n$_"
    $errors += $errorMessage
    Write-Host $errorMessage -ForegroundColor Red
}


# Download, extract, and set up XrmToolBox

# Define paths
$desktopPath = [System.Environment]::GetFolderPath('Desktop')
$zipFilePath = "$desktopPath\latest_release.zip"
$destinationFolder = "$desktopPath\XrmToolBox"

# Fetch the latest release information from GitHub
Write-Host "Fetching the latest release information for XrmToolBox from GitHub..."
try {
    $latestRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/MscrmTools/XrmToolBox/releases/latest" -UseBasicParsing
    $latestAssetUrl = $latestRelease.assets[0].browser_download_url
    Write-Host "Latest release information fetched successfully."
}
catch {
    $errorMessage = "Error fetching latest release information from GitHub`n$_"
    $errors += $errorMessage
    Write-Host $errorMessage -ForegroundColor Red
}

# Download the asset (zip file) to the desktop
Write-Host "Downloading XrmToolBox asset (zip file) to the desktop..."
try {
    Invoke-WebRequest -Uri $latestAssetUrl -OutFile $zipFilePath
    Write-Host "XrmToolBox asset downloaded successfully."
}
catch {
    $errorMessage = "Error downloading the asset (zip file) from GitHub`n$_"
    $errors += $errorMessage
    Write-Host $errorMessage -ForegroundColor Red
}

# Create the destination folder if it doesn't exist
Write-Host "Creating destination folder for XrmToolBox..."
if (-Not (Test-Path -Path $destinationFolder)) {
    try {
        New-Item -ItemType Directory -Path $destinationFolder | Out-Null
        Write-Host "Destination folder created successfully."
    }
    catch {
        $errorMessage = "Error creating the destination folder: $destinationFolder`n$_"
        $errors += $errorMessage
        Write-Host $errorMessage -ForegroundColor Red
    }
}

# Unzip the downloaded file to the destination folder
Write-Host "Unzipping the downloaded file to the destination folder..."
try {
    Expand-Archive -Path $zipFilePath -DestinationPath $destinationFolder -Force
    Write-Host "File unzipped successfully."
}
catch {
    $errorMessage = "Error unzipping the downloaded file to the destination folder`n$_"
    $errors += $errorMessage
    Write-Host $errorMessage -ForegroundColor Red
}

# Delete the zip file after extraction
Write-Host "Deleting the zip file after extraction..."
try {
    Remove-Item -Path $zipFilePath
    Write-Host "Zip file deleted successfully."
}
catch {
    $errorMessage = "Error deleting the zip file: $zipFilePath`n$_"
    $errors += $errorMessage
    Write-Host $errorMessage -ForegroundColor Red
}

# Output the path to the extracted folder
Write-Host "XrmToolBox has been downloaded and extracted to: $destinationFolder"

# Configure Git with user name and email from Windows
Write-Host "Configuring Git with user name and email from Windows..."
try {
    $userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $userEmail = "$userName@example.com"
    
    git config --global user.name "$userName"
    git config --global user.email "$userEmail"
    Write-Host "Git configured successfully with user name and email."
}
catch {
    $errorMessage = "Error configuring Git with user name and email`n$_"
    $errors += $errorMessage
    Write-Host $errorMessage -ForegroundColor Red
}

# Install VSCode extensions
Install-VSCodeExtension "ms-azure-devops.azure-pipelines" "Azure Repos"
Install-VSCodeExtension "ms-azuretools.vscode-azureresourcegroups" "Azure Resources"
Install-VSCodeExtension "ms-dataverse.dataverse-devtools" "Dataverse DevTools"
Install-VSCodeExtension "ms-vscode.azure-account" "Azure Account"
Install-VSCodeExtension "microsoft-Isvexptools.powerplatform-vscode" "Power Platform Tools"
Install-VSCodeExtension "ms-vscode.powershell" "PowerShell"
Install-VSCodeExtension "github.remotehub" "Remote Repositories"

# Output any errors
if ($errors.Count -gt 0) {
    Write-Host "`nErrors encountered during the execution of the script:`n" -ForegroundColor Red
    $errors | ForEach-Object { Write-Host $_ -ForegroundColor Red }
} else {
    # TODO: Not catching all the errors
    # Write-Host "`nNo errors encountered during the execution of the script."
}
