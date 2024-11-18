# Author : Alif Amzari
# Purpose: KB0017794/PRB0001166 - MS Apps sign in issue due to multiple account. Script is to clear Web Account Manager state. 

# Define the log file path
$logFilePath = "C:\Windows\Logs\WAM_ClearState.log"

# Check if the log file already exists; if it does, clear its contents
if (Test-Path -Path $logFilePath) {
    Clear-Content -Path $logFilePath
} else {
    "Log file does not exist. It will be created when needed." | Write-Output
}

# Custom Write-Log function
function Write-Log {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory = $true)]
        [String]$Path,

        [parameter(Mandatory = $true)]
        [String]$Message,

        [parameter(Mandatory = $true)]
        [String]$Component,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Info", "Warning", "Error")]
        [String]$Type
    )

    switch ($Type) {
        "Info" { [int]$Type = 1 }
        "Warning" { [int]$Type = 2 }
        "Error" { [int]$Type = 3 }
    }

    # Create a log entry
    $Content = "<![LOG[$Message]LOG]!>" + `
        "<time=`"$(Get-Date -Format "HH:mm:ss.ffffff")`" " + `
        "date=`"$(Get-Date -Format "M-d-yyyy")`" " + `
        "component=`"$Component`" " + `
        "context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " + `
        "type=`"$Type`" " + `
        "thread=`"$([Threading.Thread]::CurrentThread.ManagedThreadId)`" " + `
        "file=`"`">"

    # Write the line to the log file
    Add-Content -Path $Path -Value $Content
}

# Log the start of the script
Write-Log -Path $logFilePath -Message "Starting WAM state clearing process..." -Component "Main" -Type "Info"

# Function to get the SID of the logged-in user
function Get-LoggedInUserSID {
    $currentUser = (Get-WmiObject -Class Win32_Process -Filter "name = 'explorer.exe'").GetOwner().User
    Write-Log -Path $logFilePath -Message "Current User: $currentUser" -Component "Get-LoggedInUserSID" -Type "Info"

    $userSID = (Get-WmiObject Win32_UserAccount -Filter "Name='$currentUser'").SID
    Write-Log -Path $logFilePath -Message "User SID: $userSID" -Component "Get-LoggedInUserSID" -Type "Info"

    return $userSID
}

# Get the logged-in user's SID
$userSID = Get-LoggedInUserSID
if (-not $userSID) {
    Write-Log -Path $logFilePath -Message "Could not retrieve the logged-in user SID. Exiting..." -Component "Main" -Type "Error"
    exit
}

# Get the profile directory of the logged-in user based on their SID
$userProfilePath = (Get-WmiObject Win32_UserProfile | Where-Object { $_.SID -eq $userSID }).LocalPath
Write-Log -Path $logFilePath -Message "User Profile Path: $userProfilePath" -Component "Main" -Type "Info"

# Construct the LocalAppData path for the logged-in user
$localAppDataPath = Join-Path $userProfilePath "AppData\Local"
Write-Log -Path $logFilePath -Message "LocalAppData Path: $localAppDataPath" -Component "Main" -Type "Info"

# Define the file paths using the logged-in user's LocalAppData path
$accountFolderPath = Join-Path $localAppDataPath "Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker\Accounts"
$settingsFilePath = Join-Path $localAppDataPath "Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\Settings\settings.dat"

# Stop the TokenBroker service and set its startup type to Disabled
Write-Log -Path $logFilePath -Message "Disabling and stopping the TokenBroker service..." -Component "Service" -Type "Info"
Set-Service -Name "TokenBroker" -StartupType Disabled
Stop-Service -Name "TokenBroker" -Force -PassThru

# Check for the account folder and handle it accordingly
if (Test-Path -Path $accountFolderPath) {
    Write-Log -Path $logFilePath -Message "Deleting account files $accountFolderPath\* " -Component "Main" -Type "Info"
    Remove-Item -Path $accountFolderPath\* -Recurse -Force
} else {
    Write-Log -Path $logFilePath -Message "Account folder not found. Re-registering Microsoft.AAD.BrokerPlugin app." -Component "AppX" -Type "Warning"
    try {
        $manifestPath = (Get-AppxPackage -Name "Microsoft.AAD.BrokerPlugin").InstallLocation + "\AppxManifest.xml"
        Add-AppxPackage -Register $manifestPath -DisableDevelopmentMode -ForceApplicationShutdown
        Write-Log -Path $logFilePath -Message "Microsoft.AAD.BrokerPlugin re-registered successfully." -Component "AppX" -Type "Info"
    } catch {
        Write-Log -Path $logFilePath -Message "Failed to re-register Microsoft.AAD.BrokerPlugin: $_" -Component "AppX" -Type "Error"
    }
}

# Backup and delete the settings.dat file
if (Test-Path -Path $settingsFilePath) {
    Write-Log -Path $logFilePath -Message "Backing up and deleting settings.dat..." -Component "Main" -Type "Info"
    Copy-Item -Path $settingsFilePath -Destination "$settingsFilePath.bak" -Force
    Remove-Item -Path $settingsFilePath -Force
} else {
    Write-Log -Path $logFilePath -Message "settings.dat not found. Skipping..." -Component "Main" -Type "Warning"
}

# Registry path in the HKEY_USERS hive
$registryPath = "Registry::HKEY_USERS\$userSID\Software\Microsoft\IdentityCRL\TokenBroker\DefaultAccount"
$backupRegistryPath = "Registry::HKEY_USERS\$userSID\Software\Microsoft\IdentityCRL\TokenBroker\DefaultAccount_backup"

# Rename the DefaultAccount registry key for the logged-in user, with error handling if it is already renamed
if (Test-Path -Path $backupRegistryPath) {
    Write-Log -Path $logFilePath -Message "DefaultAccount registry key already renamed to DefaultAccount_backup. Skipping..." -Component "Registry" -Type "Warning"
} elseif (Test-Path -Path $registryPath) {
    Write-Log -Path $logFilePath -Message "Renaming DefaultAccount registry key..." -Component "Registry" -Type "Info"
    try {
        Rename-Item -Path $registryPath -NewName $backupRegistryPath.Split('\')[-1]
        Write-Log -Path $logFilePath -Message "DefaultAccount registry key renamed successfully." -Component "Registry" -Type "Info"
    } catch {
        Write-Log -Path $logFilePath -Message "Failed to rename DefaultAccount registry key: $_" -Component "Registry" -Type "Error"
    }
} else {
    Write-Log -Path $logFilePath -Message "DefaultAccount registry key not found. Skipping..." -Component "Registry" -Type "Warning"
}

# Set the TokenBroker service back to Manual startup and restart it
Write-Log -Path $logFilePath -Message "Setting the TokenBroker service to Manual startup and restarting it..." -Component "Service" -Type "Info"
Set-Service -Name "TokenBroker" -StartupType Manual
Start-Service -Name "TokenBroker" -PassThru

Write-Log -Path $logFilePath -Message "WAM state clearing process completed successfully." -Component "Main" -Type "Info"
