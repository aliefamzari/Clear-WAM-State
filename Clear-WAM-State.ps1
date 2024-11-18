# Author : Alif Amzari
# Purpose : For PRB0001166, KB0017794  - MS-Office apps sign in error due to duplicate Office accounts

# Define the log file path
$logFilePath = "C:\Windows\Logs\WAM_ClearLog.txt"

# Function to write logs to the specified log file in CMTrace-compatible format, supporting pipeline input
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [string]$message,
        
        [string]$logLevel = "INFO"  # Default log level
    )

    # Process pipeline input
    process {
        # Append a timestamp and format the message for CMTrace
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $component = "WAM_Clear"
        $logMessage = "[$timestamp] [$logLevel] [$component] $message"
        
        # Write the message to the log file
        Add-Content -Path $logFilePath -Value $logMessage
    }
}

# Log the start of the script
"Starting WAM state clearing process..." | Write-Log

# Function to get the SID of the logged-in user
function Get-LoggedInUserSID {
    $currentUser = (Get-WmiObject -Class Win32_Process -Filter "name = 'explorer.exe'").GetOwner().User
    "Current User: $currentUser" | Write-Output | Write-Log

    $userSID = (Get-WmiObject Win32_UserAccount -Filter "Name='$currentUser'").SID
    "User SID: $userSID" | Write-Output | Write-Log

    return $userSID
}

# Get the logged-in user's SID
$userSID = Get-LoggedInUserSID
if (-not $userSID) {
    "Could not retrieve the logged-in user SID. Exiting..." | Write-Log -logLevel "ERROR"
    exit
}

# Registry path in the HKEY_USERS hive
$registryPath = "Registry::HKEY_USERS\$userSID\Software\Microsoft\IdentityCRL\TokenBroker\DefaultAccount"
$backupRegistryPath = "Registry::HKEY_USERS\$userSID\Software\Microsoft\IdentityCRL\TokenBroker\DefaultAccount_backup"

# Stop the TokenBroker service and set its startup type to Disabled
"Disabling and stopping the TokenBroker service..." | Write-Log
Set-Service -Name "TokenBroker" -StartupType Disabled
Stop-Service -Name "TokenBroker" -Force -PassThru

# Define the file paths
$accountFolderPath = "$env:LOCALAPPDATA\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker\Accounts"
$settingsFilePath = "$env:LOCALAPPDATA\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\Settings\settings.dat"

# Check for the account folder and handle it accordingly
if (Test-Path -Path $accountFolderPath) {
    "Deleting account files..." | Write-Log
    Remove-Item -Path $accountFolderPath\* -Recurse -Force
} else {
    "Account folder not found. Re-registering Microsoft.AAD.BrokerPlugin app." | Write-Log -logLevel "WARNING"
    
    # Retrieve the AppxManifest path and re-register the package
    try {
        $manifestPath = (Get-AppxPackage -Name "Microsoft.AAD.BrokerPlugin").InstallLocation + "\AppxManifest.xml"
        if (Test-Path -Path $manifestPath) {
            "Re-registering Microsoft.AAD.BrokerPlugin..." | Write-Log
            Add-AppxPackage -Register $manifestPath -DisableDevelopmentMode -ForceApplicationShutdown
            "Re-registration of Microsoft.AAD.BrokerPlugin completed successfully." | Write-Log
        } else {
            "Manifest file not found at $manifestPath. Re-registration failed." | Write-Log -logLevel "ERROR"
        }
    } catch {
        "Error retrieving Microsoft.AAD.BrokerPlugin package information: $_" | Write-Log -logLevel "ERROR"
    }
}

# Backup and delete the settings.dat file
if (Test-Path -Path $settingsFilePath) {
    "Backing up and deleting settings.dat..." | Write-Log
    Copy-Item -Path $settingsFilePath -Destination "$settingsFilePath.bak" -Force
    Remove-Item -Path $settingsFilePath -Force
} else {
    "settings.dat not found. Skipping backup and deletion." | Write-Log -logLevel "WARNING"
}

# Rename the DefaultAccount registry key for the logged-in user, with error handling if it is already renamed
if (Test-Path -Path $backupRegistryPath) {
    "DefaultAccount registry key has already been renamed to DefaultAccount_backup. Skipping renaming step." | Write-Log -logLevel "WARNING"
} elseif (Test-Path -Path $registryPath) {
    "Renaming the DefaultAccount registry key for the logged-in user..." | Write-Log
    try {
        Rename-Item -Path $registryPath -NewName $backupRegistryPath.Split('\')[-1]
        "DefaultAccount registry key successfully renamed to DefaultAccount_backup." | Write-Log
    } catch {
        "Failed to rename DefaultAccount registry key: $_" | Write-Log -logLevel "ERROR"
    }
} else {
    "DefaultAccount registry key not found. Skipping renaming." | Write-Log -logLevel "WARNING"
}

# Set the TokenBroker service back to Manual startup and restart it
"Setting the TokenBroker service to Manual startup and restarting it..." | Write-Log
Set-Service -Name "TokenBroker" -StartupType Manual
Start-Service -Name "TokenBroker" -PassThru

"WAM state cleared successfully." | Write-Log
