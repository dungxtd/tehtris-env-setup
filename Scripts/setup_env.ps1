#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Security Testing Environment Setup Script
.DESCRIPTION
    Automates the setup of a security testing environment with proper logging and error handling.
    Temporarily disables security features, installs tools, and provides cleanup functionality.
.NOTES
    Author: Cascade AI Assistant
    Created: 2025-08-25
    Requires: Administrator privileges
#>

param(
    [switch]$SkipToolInstallation,
    [switch]$CleanupOnly,
    [string]$BaseDir = "C:\Temp\SecurityTesting",
    [string]$LogPath = "$BaseDir\Logs\setup_env.log"
)

# Global variables for state tracking
$Global:OriginalDefenderState = $null
$Global:OriginalFirewallStates = @{}
$Global:BaseDir = $BaseDir
$Global:ScriptsDir = Join-Path $BaseDir "Scripts"
$Global:ToolsDir = Join-Path $BaseDir "Tools"
$Global:LogsDir = Join-Path $BaseDir "Logs"
$Global:TempDir = "C:\Temp" # For temporary downloads
$Global:LogPath = $LogPath
$Global:ScriptStartTime = Get-Date

#region Logging Functions
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    # Color coding for console output
    $color = switch ($Level) {
        "INFO" { "White" }
        "WARN" { "Yellow" }
        "ERROR" { "Red" }
        "SUCCESS" { "Green" }
    }

    Write-Host $logMessage -ForegroundColor $color

    # Ensure log directory exists
    $logDir = Split-Path $Global:LogPath -Parent
    if (!(Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    # Write to log file
    $logMessage | Out-File -FilePath $Global:LogPath -Append -Encoding UTF8
}

function Write-Section {
    param([string]$Title)

    $separator = "=" * 60
    Write-Log $separator
    Write-Log "  $Title"
    Write-Log $separator
}
#endregion

#region Administrative Privilege Check
function Test-AdminPrivileges {
    Write-Log "Checking for Administrator privileges..." -Level "INFO"

    try {
        $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
        $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

        if ($isAdmin) {
            Write-Log "Administrator privileges confirmed." -Level "SUCCESS"
            return $true
        } else {
            Write-Log "Administrator privileges required. Please run as Administrator." -Level "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error checking administrative privileges: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}
#endregion

#region Security Configuration Functions
function Save-DefenderState {
    Write-Log "Saving current Windows Defender state..." -Level "INFO"

    try {
        $defenderPrefs = Get-MpPreference
        $Global:OriginalDefenderState = @{
            DisableRealtimeMonitoring = $defenderPrefs.DisableRealtimeMonitoring
            DisableBehaviorMonitoring = $defenderPrefs.DisableBehaviorMonitoring

        }
        Write-Log "Windows Defender state saved successfully." -Level "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Error saving Windows Defender state: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Disable-WindowsDefender {
    Write-Log "Temporarily disabling Windows Defender real-time protection..." -Level "INFO"

    try {
        Set-MpPreference -DisableRealtimeMonitoring $true
        Set-MpPreference -DisableBehaviorMonitoring $true


        Write-Log "Windows Defender real-time protection disabled." -Level "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Error disabling Windows Defender: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Save-FirewallStates {
    Write-Log "Saving current Windows Firewall states..." -Level "INFO"

    try {
        $profiles = @("Domain", "Private", "Public")

        foreach ($profile in $profiles) {
            $firewallProfile = Get-NetFirewallProfile -Name $profile
            $Global:OriginalFirewallStates[$profile] = $firewallProfile.Enabled
        }

        Write-Log "Windows Firewall states saved successfully." -Level "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Error saving Windows Firewall states: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Disable-WindowsFirewall {
    Write-Log "Temporarily disabling Windows Firewall for all profiles..." -Level "INFO"

    try {
        $profiles = @("Domain", "Private", "Public")

        foreach ($profile in $profiles) {
            Set-NetFirewallProfile -Name $profile -Enabled False
            Write-Log "Disabled firewall for $profile profile." -Level "INFO"
        }

        Write-Log "Windows Firewall disabled for all profiles." -Level "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Error disabling Windows Firewall: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}
#endregion


#endregion

#region Environment Variable Functions
function Add-DirectoryToPath {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Directory
    )

    Write-Log "Adding $Directory to system PATH..." -Level "INFO"

    try {
        $systemPath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")

        if ($systemPath -split ';' -notcontains $Directory) {
            $newPath = "$systemPath;$Directory"
            [System.Environment]::SetEnvironmentVariable("Path", $newPath, "Machine")
            $env:Path = $newPath # Update for current session
            Write-Log "$Directory added to system PATH successfully." -Level "SUCCESS"
        } else {
            Write-Log "$Directory is already in the system PATH. Skipping." -Level "INFO"
        }
        return $true
    }
    catch {
        Write-Log "Error adding directory to PATH: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

#region Tool Installation Functions
function Install-Mimikatz {
    Write-Log "Starting Mimikatz installation..." -Level "INFO"

    try {
        $mimikatzDir = Join-Path $Global:ToolsDir "mimikatz"
        $mimikatzExe = Join-Path $mimikatzDir "x64\mimikatz.exe"
        $mimikatzX64Dir = Join-Path $mimikatzDir "x64"

        # Always ensure it's in the PATH if installed
        if (Test-Path $mimikatzExe) {
            Add-DirectoryToPath -Directory $mimikatzX64Dir
            Write-Log "Mimikatz already installed. Skipping download." -Level "INFO"
            return $true
        }

        $tempZip = Join-Path $Global:TempDir "mimikatz.zip"

        # Create directory if it doesn't exist
        if (!(Test-Path $mimikatzDir)) {
            New-Item -ItemType Directory -Path $mimikatzDir -Force | Out-Null
        }

        # Check if installer already exists
        if (Test-Path $tempZip) {
            Write-Log "Mimikatz installer already downloaded. Using existing file." -Level "INFO"
        } else {
            Write-Log "Downloading Mimikatz from GitHub..." -Level "INFO"

            # Get latest release URL
            $apiUrl = "https://api.github.com/repos/gentilkiwi/mimikatz/releases/latest"
            $release = Invoke-RestMethod -Uri $apiUrl
            $downloadUrl = ($release.assets | Where-Object { $_.name -like "*mimikatz_trunk.zip" }).browser_download_url

            if (!$downloadUrl) {
                throw "Could not find Mimikatz download URL"
            }

            # Download the ZIP file
            Invoke-WebRequest -Uri $downloadUrl -OutFile $tempZip -UseBasicParsing
            Write-Log "Mimikatz downloaded successfully." -Level "SUCCESS"
        }

        # Extract the ZIP file
        Write-Log "Extracting Mimikatz..." -Level "INFO"
        Expand-Archive -Path $tempZip -DestinationPath $mimikatzDir -Force

        # Clean up temporary file (disabled)
        # Remove-Item $tempZip -Force

        # Add Mimikatz to PATH
        $mimikatzX64Dir = Join-Path $mimikatzDir "x64"
        Add-DirectoryToPath -Directory $mimikatzX64Dir

        Write-Log "Mimikatz installation completed successfully." -Level "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Error installing Mimikatz: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}



function Install-Nmap {
    Write-Log "Starting Nmap installation..." -Level "INFO"

    try {
        $installDir = "C:\Program Files (x86)\Nmap"
        $nmapExe = Join-Path $installDir "nmap.exe"

        if (Test-Path $nmapExe) {
            Write-Log "Nmap already installed. Skipping download." -Level "INFO"
            return $true
        }

        $tempInstaller = Join-Path $Global:TempDir "nmap-installer.exe"

        # Check if installer already exists
        if (Test-Path $tempInstaller) {
            Write-Log "Nmap installer already downloaded. Using existing file." -Level "INFO"
        } else {
            Write-Log "Downloading Nmap installer..." -Level "INFO"
            $nmapPage = Invoke-WebRequest -Uri "https://nmap.org/download.html" -UseBasicParsing
            $installerPattern = 'https://nmap\.org/dist/nmap-[\d\.]+-setup\.exe'
            $downloadUrl = [regex]::Match($nmapPage.Content, $installerPattern).Value

            if (!$downloadUrl) {
                throw "Could not find Nmap installer download URL"
            }

            Invoke-WebRequest -Uri $downloadUrl -OutFile $tempInstaller -UseBasicParsing
            Write-Log "Nmap installer downloaded successfully." -Level "SUCCESS"
        }

        # Launch interactive installer
        Write-Log "Launching Nmap installer. Please follow the on-screen instructions to install Nmap and Npcap." -Level "INFO"
        Write-Log "The script will continue after you close the installer." -Level "INFO"
        $installProcess = Start-Process -FilePath $tempInstaller -Wait -PassThru

        if (Test-Path $nmapExe) {
            Write-Log "Nmap installation confirmed." -Level "SUCCESS"
            return $true
        } else {
            throw "Nmap installation could not be confirmed. Please install manually or re-run the script."
        }
    }
    catch {
        Write-Log "Error installing Nmap: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}
#endregion

#region TEHTRIS EDR Integration Functions
function Invoke-TehtrisEdrInstaller {
    param(
        [Parameter(Mandatory=$true)]
        [string]$MsiPath
    )

    Write-Log "Invoking TEHTRIS EDR installer..." -Level "INFO"

    try {
        if (!(Test-Path $MsiPath)) {
            throw "MSI file not found at path: $MsiPath"
        }

        $pythonScript = Join-Path $Global:ScriptsDir "tehtris_edr_installer_minimal.py"
        if (!(Test-Path $pythonScript)) {
            throw "TEHTRIS EDR installer script not found at: $pythonScript"
        }

        Write-Log "Executing TEHTRIS EDR installer script..." -Level "INFO"
        $process = Start-Process -FilePath "python" -ArgumentList "`"$pythonScript`" `"$MsiPath`"" -Wait -PassThru -NoNewWindow

        if ($process.ExitCode -eq 0) {
            Write-Log "TEHTRIS EDR installer executed successfully." -Level "SUCCESS"
            return $true
        } else {
            throw "TEHTRIS EDR installer failed with exit code: $($process.ExitCode)"
        }
    }
    catch {
        Write-Log "Error invoking TEHTRIS EDR installer: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Invoke-TehtrisEdrUninstaller {
    param(
        [string]$Password,
        [string]$KeyFilePath
    )

    Write-Log "Invoking TEHTRIS EDR uninstaller..." -Level "INFO"

    try {
        $pythonScript = Join-Path $Global:ScriptsDir "tehtris_edr_uninstaller.py"
        if (!(Test-Path $pythonScript)) {
            throw "TEHTRIS EDR uninstaller script not found at: $pythonScript"
        }

        $arguments = "`"$pythonScript`""

        if ($Password) {
            $arguments += " --password `"$Password`""
        } elseif ($KeyFilePath) {
            if (!(Test-Path $KeyFilePath)) {
                throw "Key file not found at path: $KeyFilePath"
            }
            $arguments += " --key-file `"$KeyFilePath`""
        } else {
            throw "Either Password or KeyFilePath must be provided"
        }

        Write-Log "Executing TEHTRIS EDR uninstaller script..." -Level "INFO"
        $process = Start-Process -FilePath "python" -ArgumentList $arguments -Wait -PassThru -NoNewWindow

        if ($process.ExitCode -eq 0) {
            Write-Log "TEHTRIS EDR uninstaller executed successfully." -Level "SUCCESS"
            return $true
        } else {
            throw "TEHTRIS EDR uninstaller failed with exit code: $($process.ExitCode)"
        }
    }
    catch {
        Write-Log "Error invoking TEHTRIS EDR uninstaller: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}
#endregion

#region Cleanup and Restoration Functions
function Restore-WindowsDefender {
    Write-Log "Restoring Windows Defender to original state..." -Level "INFO"

    try {
        if ($Global:OriginalDefenderState) {
            Set-MpPreference -DisableRealtimeMonitoring $Global:OriginalDefenderState.DisableRealtimeMonitoring
            Set-MpPreference -DisableBehaviorMonitoring $Global:OriginalDefenderState.DisableBehaviorMonitoring


            Write-Log "Windows Defender restored to original state." -Level "SUCCESS"
            return $true
        } else {
            Write-Log "No original Windows Defender state found to restore." -Level "WARN"
            return $false
        }
    }
    catch {
        Write-Log "Error restoring Windows Defender: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Restore-WindowsFirewall {
    Write-Log "Restoring Windows Firewall to original states..." -Level "INFO"

    try {
        if ($Global:OriginalFirewallStates.Count -gt 0) {
            foreach ($profile in $Global:OriginalFirewallStates.Keys) {
                $originalState = $Global:OriginalFirewallStates[$profile]
                Set-NetFirewallProfile -Name $profile -Enabled $originalState
                Write-Log "Restored firewall for $profile profile to: $originalState" -Level "INFO"
            }

            Write-Log "Windows Firewall restored to original states." -Level "SUCCESS"
            return $true
        } else {
            Write-Log "No original Windows Firewall states found to restore." -Level "WARN"
            return $false
        }
    }
    catch {
        Write-Log "Error restoring Windows Firewall: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Invoke-Cleanup {
    Write-Section "CLEANUP AND RESTORATION"

    $cleanupSuccess = $true

    # Restore Windows Defender
    if (!(Restore-WindowsDefender)) {
        $cleanupSuccess = $false
    }

    # Restore Windows Firewall
    if (!(Restore-WindowsFirewall)) {
        $cleanupSuccess = $false
    }

    if ($cleanupSuccess) {
        Write-Log "All security settings restored successfully." -Level "SUCCESS"
    } else {
        Write-Log "Some security settings could not be restored. Please check manually." -Level "WARN"
    }

    return $cleanupSuccess
}
#endregion


#region Main Orchestration
function Initialize-Environment {
    Write-Section "ENVIRONMENT INITIALIZATION"

    # Check administrative privileges
    if (!(Test-AdminPrivileges)) {
        throw "Administrator privileges required. Exiting."
    }

    # Create project directories if they don't exist
    $dirsToCreate = @($Global:BaseDir, $Global:ScriptsDir, $Global:ToolsDir, $Global:LogsDir, $Global:TempDir)
    foreach ($dir in $dirsToCreate) {
        if (!(Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            Write-Log "Created directory: $dir" -Level "INFO"
        }
    }

    Write-Log "Environment initialization completed." -Level "SUCCESS"
}

function Invoke-SecurityPreConfiguration {
    Write-Section "SECURITY PRE-CONFIGURATION"

    $success = $true

    # Save and disable Windows Defender
    if (Save-DefenderState) {
        if (!(Disable-WindowsDefender)) {
            $success = $false
        }
    } else {
        $success = $false
    }

    # Save and disable Windows Firewall
    if (Save-FirewallStates) {
        if (!(Disable-WindowsFirewall)) {
            $success = $false
        }
    } else {
        $success = $false
    }

    if ($success) {
        Write-Log "Security pre-configuration completed successfully." -Level "SUCCESS"
    } else {
        Write-Log "Security pre-configuration encountered errors." -Level "ERROR"
    }

    return $success
}

function Invoke-ToolDeployment {
    Write-Section "TOOL DEPLOYMENT"

    $success = $true

    # Install Mimikatz
    if (!(Install-Mimikatz)) {
        $success = $false
    }

    # Install Nmap
    if (!(Install-Nmap)) {
        $success = $false
    }

    if ($success) {
        Write-Log "Tool deployment completed successfully." -Level "SUCCESS"
    } else {
        Write-Log "Tool deployment encountered errors." -Level "ERROR"
    }

    return $success
}

function Start-Orchestration {
    Write-Section "STARTING SECURITY ENVIRONMENT SETUP"

    try {
        Initialize-Environment

        if ($CleanupOnly) {
            Write-Log "CleanupOnly parameter specified. Skipping main setup." -Level "INFO"
            return # Skip to finally block for cleanup
        }

        # Main setup steps
        if (Invoke-SecurityPreConfiguration) {
            if (-not $SkipToolInstallation) {
                Invoke-ToolDeployment
            } else {
                Write-Log "SkipToolInstallation parameter specified. Skipping tool deployment." -Level "INFO"
            }
        }

        Write-Log "Main setup tasks completed. The environment is ready." -Level "SUCCESS"
    }
    catch {
        Write-Log "An unexpected error occurred during setup: $($_.Exception.Message)" -Level "ERROR"
    }
    finally {
        # This block will always execute, ensuring cleanup
        Invoke-Cleanup

        $scriptEndTime = Get-Date
        $scriptDuration = New-TimeSpan -Start $Global:ScriptStartTime -End $scriptEndTime
        Write-Section "SCRIPT EXECUTION SUMMARY"
        Write-Log "Script finished in: $($scriptDuration.TotalSeconds) seconds."
    }
}

# Entry point of the script
Start-Orchestration
#endregion
