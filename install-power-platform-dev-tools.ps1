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
Execute-WingetCommand "winget install
