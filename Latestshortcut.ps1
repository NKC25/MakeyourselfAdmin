#author- Naveen kumar chary Sriramoj
#script will download the makeyourselfadmin-virtual from github and creates the shortcut in 
# Define parameters
param (
    [string]$action = "install"  # Default action is install
)

# Define variables
$exeUrl = "https://github.com/NKC25/MakeyourselfAdmin/raw/main/MakeYourselfAdmin-Virtual.exe"
$destinationPath = "C:\PIM_Activation\MakeYourselfAdmin-Virtual.exe"
$shortcutName = "MakeYourselfAdmin-Virtual"
$startMenuPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\$shortcutName.lnk"
$oneDrivePath = "$env:OneDrive"
$desktopPath = "$oneDrivePath\Desktop\$shortcutName.lnk"
$logPath = "C:\PIM_Activation\Deploy-Log.txt"



# Create a log file
function Write-Log {
    param (
        [string]$message
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logPath -Value "$timestamp - $message"
}



    $destinationDir = Split-Path -Path $destinationPath -Parent
    if (!(Test-Path -Path $destinationDir)) {
        New-Item -ItemType Directory -Force -Path $destinationDir
        Write-Log "Created destination directory."
    }



# Function to create a shortcut
function Create-Shortcut {
    param (
        [string]$shortcutPath,
        [string]$targetPath,
        [string]$description
    )
    $WScriptShell = New-Object -ComObject WScript.Shell
    $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $targetPath
    $shortcut.WorkingDirectory = Split-Path -Path $targetPath -Parent
    $shortcut.Description = $description
    $shortcut.Save()
}

# Function to remove the existing application
function Remove-ExistingApp {
    Write-Log "Removing existing application if present."
    if (Test-Path -Path $destinationPath) {
        Remove-Item -Path $destinationPath -Force
        Write-Log "Removed existing .exe file."
    }

    if (Test-Path -Path $startMenuPath) {
        Remove-Item -Path $startMenuPath -Force
        Write-Log "Removed Start Menu shortcut."
    }

    if (Test-Path -Path $desktopPath) {
        Remove-Item -Path $desktopPath -Force
        Write-Log "Removed Desktop shortcut."
    }

 }

# Function to install the new application
function Install-NewApp {
    #Write-Log "Creating destination directory if it doesn't exist."
    
    Write-Log "Downloading the .exe file."
    Invoke-WebRequest -Uri $exeUrl -OutFile $destinationPath

    if (Test-Path -Path $destinationPath) {
        Write-Log "Downloaded .exe file successfully."
    } else {
        Write-Log "Failed to download .exe file."
    }

    Write-Log "Creating Start Menu shortcut."
    Create-Shortcut -shortcutPath $startMenuPath -targetPath $destinationPath -description $shortcutName

    Write-Log "Creating Desktop shortcut."
    Create-Shortcut -shortcutPath $desktopPath -targetPath $destinationPath -description $shortcutName
}

# Main script execution
try {
    Write-Log "Starting deployment script with action: $action."

    if ($action -eq "install") {
        # Install new application
        Install-NewApp
        Write-Log "Application installed successfully."
    } elseif ($action -eq "remove") {
        # Remove existing application
        Remove-ExistingApp
        Write-Log "Application removed successfully."
    } else {
        Write-Log "Invalid action specified. Use 'install' or 'remove'."
    }
} catch {
    Write-Log "An error occurred: $_"
}
