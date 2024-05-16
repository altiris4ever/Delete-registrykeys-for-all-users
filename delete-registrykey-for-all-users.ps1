# Define the domain to be trimmed from usernames
$Domain = "domain\"

# Define the array of registry subkey paths that will be deleted
$regKeys = @(
    "SOFTWARE\Microsoft\Office\Outlook\Addins\SAS.OutlookAddIn",
    "Software\Microsoft\Office\Excel\Addins\SAS.ExcelAddIn",
    "Software\Microsoft\Office\PowerPoint\Addins\SAS.PowerPointAddIn",
    "Software\Microsoft\Office\Word\Addins\SAS.WordAddIn"
)

# Define registry hives
$reghive = "HKCU:\"
$reghive1 = "HKU"

# Log file location and name
$logFilePath = "C:\temp"
$logFileName = "Sas_reg_delete.log"
$fullLogPath = Join-Path $logFilePath $logFileName

# Initialize flag for all keys not found
$allKeysNotFound = $true

# Function to write output to both console and log file with timestamp for certain actions
function Write-Log {
    param (
        [string]$Message,
        [bool]$IncludeTimestamp = $false
    )
    if ($IncludeTimestamp) {
        $timestamp = Get-Date -Format "dd-MM-yy HH:mm:ss"
        $Message = "$($timestamp): $Message"
    }
    Write-Host $Message
    Add-Content -Path $fullLogPath -Value $Message
}

# Ensure the log directory exists
if (-not (Test-Path $logFilePath)) {
    New-Item -ItemType Directory -Path $logFilePath | Out-Null
}

# Function to delete registry keys and handle errors with timestamp
function Delete-RegistryKey {
    param (
        [string]$Path,
        [string]$Username,
        [string]$Key
    )
    try {
        Remove-Item -Path $Path -Force -Recurse
        $deleted = $true
        Write-Log "Registry key $Key has been deleted for user $Username." $true
    } catch {
        Write-Log "ERROR: Failed to delete registry key $Key for user $Username. Error: $_" $true
        $deleted = $false
    }
    return $deleted
}

# Log script start time
Write-Log "Script started." $true

# Retrieve all user profiles, excluding system profiles
$userProfiles = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.Special -eq $false }

# Process each registry key in the array
foreach ($regKey in $regKeys) {
    # Convert the regKey to the full path for HKCU
    $registryKeyPath = $reghive + $regKey

    # Log entry for each key
    Write-Log "******* Searching for and attempting to delete the following user registry key: ******"
    Write-Log $registryKeyPath
    $keyFoundAndDeleted = $false

    # Iterate through each user profile and delete the specified registry key
    foreach ($userProfile in $userProfiles) {
        $userSID = $userProfile.SID
        $SID = New-Object System.Security.Principal.SecurityIdentifier($userSID)
        $User = $SID.Translate([System.Security.Principal.NTAccount])
        $userHivePath = "Registry::HKEY_USERS\$userSID\$regKey"
        $user_trim = $User.Value.Trim("$Domain")

        # Check if the registry key exists in the user's hive
        if (Test-Path -Path $userHivePath) {
            Write-Log "Registry cleanup is being performed for logged on user: $user_trim"
            if (Delete-RegistryKey -Path $userHivePath -Username $user_trim -Key $regKey) {
                $keyFoundAndDeleted = $true
                $allKeysNotFound = $false
            }
        }
    }

    # Enumerate user directories, excluding certain system profiles
    $loggedOffUsers = Get-ChildItem -Path C:\Users | Where-Object { $_.Name -notmatch 'Public|Administrator|defaultuser0' }

    # Process each logged off user if a key was deleted for any user profile
    foreach ($userDir in $loggedOffUsers) {
        $TempName = $userDir.Name
        $TempHive = "HKU\$TempName"
        $ProfilePath = Join-Path $userDir.FullName NTUSER.DAT

        # Load the user's registry hive
        reg load $TempHive $ProfilePath 2>&1 | Out-Null

        $RegistryKey = "$TempHive\$regKey"

        # Check if the registry key exists before attempting to delete it
        if (Test-Path "Registry::$RegistryKey") {
            Write-Log "************** Attempting registry cleanup for logged off user: $TempName **************"
            if (Delete-RegistryKey -Path "Registry::$RegistryKey" -Username $TempName -Key $regKey) {
                $keyFoundAndDeleted = $true
                $allKeysNotFound = $false
            }
        }

        # Unload the user's registry hive
        reg unload $TempHive 2>&1 | Out-Null
    }

    # If the registry key was not found and not deleted, log that no keys were found
    if (-not $keyFoundAndDeleted) {
        Write-Log "Registry key was not found for any users."
    }
    Write-Log "**************************************************************************************"
}

# End of script
Write-Log "Script completed." $true

# Set exit code if all keys were not found
if ($allKeysNotFound) {
    exit 1
} else {
    exit 0
}
