# Author : Alif Amzari
# Purpose: KB0017794/PRB0001166 - MS Apps sign in issue due to multiple account. Script is to clear Web Account Manager state. 

# Define the log file path
$logFilePath = "C:\Windows\Logs\WAM_ClearState.log"

#region Initialize Log File
# Check if the log file already exists; if it does, clear its contents
if (Test-Path -Path $logFilePath) {
    Clear-Content -Path $logFilePath
} else {
    "Log file does not exist. It will be created when needed." | Write-Output
}
#endregion

#region Write-Log Function
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
#endregion

#region Get Logged-In User SID
function Get-LoggedInUserSID {
    $currentUser = (Get-WmiObject -Class Win32_Process -Filter "name = 'explorer.exe'").GetOwner().User
    Write-Log -Path $logFilePath -Message "Current User: $currentUser" -Component "Get-LoggedInUserSID" -Type "Info"

    $userSID = (Get-WmiObject Win32_UserAccount -Filter "Name='$currentUser'").SID
    Write-Log -Path $logFilePath -Message "User SID: $userSID" -Component "Get-LoggedInUserSID" -Type "Info"

    return $userSID
}
#endregion

#region Retrieve User Profile and Paths
# Get the logged-in user's SID
$userSID = Get-LoggedInUserSID
if (-not $userSID) {
    Write-Log -Path $logFilePath -Message "Could not retrieve the logged-in user SID. Exiting..." -Component "Get-LoggedInUserSID" -Type "Error"
    exit
}

# Get the profile directory of the logged-in user based on their SID
$userProfilePath = (Get-WmiObject Win32_UserProfile | Where-Object { $_.SID -eq $userSID }).LocalPath
Write-Log -Path $logFilePath -Message "User Profile Path: $userProfilePath" -Component "RetrieveUserProfilePath" -Type "Info"

# Construct the LocalAppData path for the logged-in user
$localAppDataPath = Join-Path $userProfilePath "AppData\Local"
Write-Log -Path $logFilePath -Message "LocalAppData Path: $localAppDataPath" -Component "RetrieveUserProfilePath" -Type "Info"
#endregion

#region Registry Modifications
$registryPath = "HKCU:\Software\Microsoft\IdentityCRL\TokenBroker\DefaultAccount"
$backupRegistryPath = "HKCU:\Software\Microsoft\IdentityCRL\TokenBroker\DefaultAccount_backup"

try {
    if (Test-Path -Path $registryPath) {
        Rename-Item -Path $registryPath -NewName $backupRegistryPath -Force
        Write-Log -Path $logFilePath -Message "Registry key renamed from DefaultAccount to DefaultAccount_backup." -Component "ModifyRegistry" -Type "Info"
    } else {
        Write-Log -Path $logFilePath -Message "Registry key DefaultAccount does not exist. Skipping..." -Component "ModifyRegistry" -Type "Warning"
    }
} catch {
    Write-Log -Path $logFilePath -Message "Failed to rename registry key: $_" -Component "ModifyRegistry" -Type "Error"
}
#endregion

#region BrokerPlugin Cleanup
# Define paths for the BrokerPlugin directory and Accounts folders
$brokerPluginPath = Join-Path $localAppDataPath "Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy"
$accountsFolderPath = Join-Path $brokerPluginPath "AC\TokenBroker\Accounts"
$accountsOldFolderPath = Join-Path $brokerPluginPath "AC\TokenBroker\Accounts.old"

# Handle Accounts.old folder
if (Test-Path -Path $accountsOldFolderPath) {
    Write-Log -Path $logFilePath -Message "Found Accounts.old folder. Deleting..." -Component "CleanupBrokerPlugin" -Type "Warning"

    try {
        Remove-Item -Path $accountsOldFolderPath -Recurse -Force
        Write-Log -Path $logFilePath -Message "Deleted Accounts.old folder." -Component "CleanupBrokerPlugin" -Type "Info"
    } catch {
        Write-Log -Path $logFilePath -Message "Failed to delete Accounts.old folder: $_" -Component "CleanupBrokerPlugin" -Type "Error"
    }
}

# Handle Accounts folder
if (Test-Path -Path $accountsFolderPath) {
    Write-Log -Path $logFilePath -Message "Found Accounts folder. Deleting its contents..." -Component "CleanupBrokerPlugin" -Type "Info"

    try {
        Remove-Item -Path $accountsFolderPath\* -Recurse -Force
        Write-Log -Path $logFilePath -Message "Deleted contents of Accounts folder." -Component "CleanupBrokerPlugin" -Type "Info"
    } catch {
        Write-Log -Path $logFilePath -Message "Failed to delete contents of Accounts folder: $_" -Component "CleanupBrokerPlugin" -Type "Error"
    }
} else {
    Write-Log -Path $logFilePath -Message "Accounts folder not found. Skipping..." -Component "CleanupBrokerPlugin" -Type "Warning"
}

# Backup and delete settings.dat file
$settingsFilePath = Join-Path $brokerPluginPath "Settings\settings.dat"
if (Test-Path -Path $settingsFilePath) {
    try {
        Copy-Item -Path $settingsFilePath -Destination "$settingsFilePath.bak" -Force
        Write-Log -Path $logFilePath -Message "settings.dat backed up successfully." -Component "CleanupBrokerPlugin" -Type "Info"
        Remove-Item -Path $settingsFilePath -Force
        Write-Log -Path $logFilePath -Message "settings.dat deleted successfully." -Component "CleanupBrokerPlugin" -Type "Info"
    } catch {
        Write-Log -Path $logFilePath -Message "Failed to handle settings.dat: $_" -Component "CleanupBrokerPlugin" -Type "Error"
    }
} else {
    Write-Log -Path $logFilePath -Message "settings.dat not found. Skipping..." -Component "CleanupBrokerPlugin" -Type "Warning"
}
#endregion

#region Re-Register BrokerPlugin
if (-not (Test-Path -Path $brokerPluginPath)) {
    Write-Log -Path $logFilePath -Message "BrokerPlugin directory not found. Attempting re-registration..." -Component "ReRegisterBrokerPlugin" -Type "Warning"

    try {
        $manifestPath = (Get-AppxPackage -Name "Microsoft.AAD.BrokerPlugin").InstallLocation + "\AppxManifest.xml"
        Add-AppxPackage -Register $manifestPath -DisableDevelopmentMode -ForceApplicationShutdown
        Write-Log -Path $logFilePath -Message "Microsoft.AAD.BrokerPlugin re-registered successfully." -Component "ReRegisterBrokerPlugin" -Type "Info"
    } catch {
        Write-Log -Path $logFilePath -Message "Failed to re-register Microsoft.AAD.BrokerPlugin: $_" -Component "ReRegisterBrokerPlugin" -Type "Error"
    }
}
#endregion

#region Restart TokenBroker Service
Write-Log -Path $logFilePath -Message "Restarting TokenBroker service..." -Component "TokenBrokerService" -Type "Info"
Set-Service -Name "TokenBroker" -StartupType Manual
Start-Service -Name "TokenBroker" -PassThru
Write-Log -Path $logFilePath -Message "TokenBroker service restarted successfully." -Component "TokenBrokerService" -Type "Info"
#endregion

#region Completion
Write-Log -Path $logFilePath -Message "WAM state clearing process completed." -Component "ScriptCompletion" -Type "Info"
#endregion
