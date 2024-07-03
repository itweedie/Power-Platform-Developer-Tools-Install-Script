param (
    [switch]$InstallAll  # Parameter to determine if all applications should be installed without prompting
)

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

# Function to prompt user for installation
function Prompt-Install {
    param (
        [string]$description
    )

    if (-not $InstallAll) {
        $response = Read-Host ("Do you want to install " + $description + "? (y/n)")
        if ($response -ne 'y') {
            Write-Host "Skipping $description."
            return $false
        }
    }
    return $true
}

# Check if InstallAll parameter is passed, if not prompt user
if (-not $InstallAll.IsPresent) {
    $response = Read-Host "Do you want to install all applications without prompting? (y/n)"
    $InstallAll = $response -eq 'y'
}

# List of applications to install
$appsToInstall = @(
    @{ command = "winget install --id=Microsoft.VisualStudioCode -e --accept-package-agreements"; description = "Visual Studio Code" },
    @{ command = "winget install --id=Microsoft.PowerBIDesktop -e --accept-package-agreements"; description = "Power BI Desktop" },
    @{ command = "winget install --id=Microsoft.SQLServerManagementStudio -e --accept-package-agreements"; description = "SQL Server Management Studio" },
    @{ command = "winget install --id=Microsoft.WindowsTerminal -e --accept-package-agreements"; description = "Windows Terminal" },
    @{ command = "winget install --id=Notepad++.Notepad++ -e --accept-package-agreements"; description = "Notepad++" },
    @{ command = "winget install --id=Microsoft.PowerToys -e --accept-package-agreements"; description = "PowerToys" },
    @{ command = "winget install --id=Postman.Postman -e --accept-package-agreements"; description = "Postman" },
    @{ command = "winget install --id=Microsoft.VisualStudio.2022.Professional -e --accept-package-agreements"; description = "Visual Studio 2022 Professional" },
    @{ command = "winget install --id=OpenJS.NodeJS -e --accept-package-agreements"; description = "Node.js (includes NPM)" },
    @{ command = "winget install --id=Git.Git -e --accept-package-agreements"; description = "Git" },
    @{ command = "winget install --id=Microsoft.AzureStorageExplorer -e --accept-package-agreements"; description = "Azure Storage Explorer" }
)

# Install applications
foreach ($app in $appsToInstall) {
    if (Prompt-Install -description $app.description) {
        Execute-WingetCommand -command $app.command -description $app.description
    }
}

# Function to install PowerShell modules
function Install-PowerShellModule {
    param (
        [string]$moduleName,
        [string]$description
    )

    if (Prompt-Install -description $description) {
        Write-Host "Installing PowerShell module: $description..."
        try {
            Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber
            Write-Host "$description installed successfully."
        }
        catch {
            $errorMessage = "Error installing $description`n$_"
            $errors += $errorMessage
            Write-Host $errorMessage -ForegroundColor Red
        }
    }
}

# Install PowerShell modules for PowerApps Administration
Install-PowerShellModule -moduleName "Microsoft.PowerApps.Administration.PowerShell" -description "Microsoft.PowerApps.Administration.PowerShell"
Install-PowerShellModule -moduleName "Microsoft.PowerApps.PowerShell" -description "Microsoft.PowerApps.PowerShell"

# Function to download and set up XrmToolBox
function Setup-XrmToolBox {
    param (
        [string]$zipFilePath,
        [string]$destinationFolder
    )

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
}

# Paths for XrmToolBox
$desktopPath = [System.Environment]::GetFolderPath('Desktop')
$zipFilePath = "$desktopPath\latest_release.zip"
$destinationFolder = "$desktopPath\XrmToolBox"

# Prompt for XrmToolBox setup
if (Prompt-Install -description "XrmToolBox") {
    Setup-XrmToolBox -zipFilePath $zipFilePath -destinationFolder $destinationFolder
}


# Configure Git with user name and email from Windows
# Function to check if Git is installed
function Check-GitInstalled {
    $gitPath = (Get-Command git -ErrorAction SilentlyContinue).Path
    return $gitPath -ne $null
}

# Configure Git with user name and email from Windows
Write-Host "Configuring Git with user name and email from Windows..."

# Check if Git is installed
if (Check-GitInstalled) {
    # Ask user if they would like Git to be configured
    $configureGit = Read-Host "Git is installed. Would you like to configure Git with your Windows user name and email? (y/n)"
    if ($configureGit -eq 'y') {
        try {
            # Get current user name
            $userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            
            # Split user name to get the user part and domain part
            $userParts = $userName.Split('\')
            $userDomain = $userParts[0]
            $userShortName = $userParts[1]
            
            # Construct the email address assuming the domain follows a pattern, e.g., user@domain.com
            $userEmail = "$userShortName@$userDomain"
            
            # Confirm user name
            $confirmedUserName = Read-Host "Current username is '$userShortName'. If this is correct, press Enter. Otherwise, enter the correct username"
            if (-not [string]::IsNullOrEmpty($confirmedUserName)) {
                $userShortName = $confirmedUserName
            }

            # Confirm user email
            $confirmedUserEmail = Read-Host "Current email is '$userEmail'. If this is correct, press Enter. Otherwise, enter the correct email"
            if (-not [string]::IsNullOrEmpty($confirmedUserEmail)) {
                $userEmail = $confirmedUserEmail
            }
            
            # Configure Git
            git config --global user.name "$userShortName"
            git config --global user.email "$userEmail"
            Write-Host "Git configured successfully with user name '$userShortName' and email '$userEmail'."
        }
        catch {
            $errorMessage = "Error configuring Git with user name and email`n$_"
            $errors += $errorMessage
            Write-Host $errorMessage -ForegroundColor Red
        }
    } else {
        Write-Host "Git configuration skipped by user."
    }
} else {
    Write-Host "The current shell can't find Git. Please install Git and re-run this script to configure Git."
}



# List of VSCode extensions to install
$vscodeExtensionsToInstall = @(
    @{ extension = "ms-azure-devops.azure-pipelines"; description = "VS Code Extension for: Azure Repos" },
    @{ extension = "ms-azuretools.vscode-azureresourcegroups"; description = "VS Code Extension for: Azure Resources" },
    @{ extension = "ms-dataverse.dataverse-devtools"; description = "VS Code Extension for: Dataverse DevTools" },
    @{ extension = "ms-vscode.azure-account"; description = "VS Code Extension for: Azure Account" },
    @{ extension = "microsoft-Isvexptools.powerplatform-vscode"; description = "VS Code Extension for: Power Platform Tools" },
    @{ extension = "ms-vscode.powershell"; description = "VS Code Extension for: PowerShell" },
    @{ extension = "github.remotehub"; description = "VS Code Extension for: Remote Repositories" }
)

# Install VSCode extensions
foreach ($extension in $vscodeExtensionsToInstall) {
    if (Prompt-Install -description $extension.description) {
        Install-VSCodeExtension -extension $extension.extension -description $extension.description
    }
}

# Output any errors
if ($errors.Count -gt 0) {
    Write-Host "`nErrors encountered during the execution of the script:`n" -ForegroundColor Red
    $errors | ForEach-Object { Write-Host $_ -ForegroundColor Red }
} else {
    Write-Host "`nNo errors encountered during the execution of the script."
}
