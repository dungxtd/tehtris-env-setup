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
    [Parameter(Mandatory=$false)]
    [switch]$DisableSec,

    [Parameter(Mandatory=$false)]
    [switch]$EnableSec,

    [Parameter(Mandatory=$false)]
    [switch]$InstallTools,

    [Parameter(Mandatory=$false)]
    [string]$InstallEdrMsiPath,

    [Parameter(Mandatory=$false)]
    [switch]$UninstallEdr,

    [Parameter(Mandatory=$false)]
    [switch]$All
)

# Global variables for state tracking

# Determine paths based on the script's location
$Global:ScriptsDir = $PSScriptRoot
$Global:BaseDir = (Get-Item $Global:ScriptsDir).Parent.FullName
$Global:ToolsDir = Join-Path $Global:BaseDir "Tools"
$Global:LogsDir = Join-Path $Global:BaseDir "Logs"
$Global:TempDir = "C:\Temp" # For temporary downloads
$Global:LogPath = Join-Path $Global:LogsDir "setup_env.log"
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
function Set-WindowsDefender {
    param(
        [Parameter(Mandatory=$true)]
        [bool]$Disable
    )

    try {
        if ($Disable) {
            Write-Log "Disabling Windows Defender real-time protection..." -Level "INFO"
            Set-MpPreference -DisableRealtimeMonitoring $true
            Set-MpPreference -DisableBehaviorMonitoring $true
            Write-Log "Windows Defender real-time protection disabled." -Level "SUCCESS"
        } else {
            Write-Log "Enabling Windows Defender real-time protection..." -Level "INFO"
            Set-MpPreference -DisableRealtimeMonitoring $false
            Set-MpPreference -DisableBehaviorMonitoring $false
            Write-Log "Windows Defender real-time protection enabled." -Level "SUCCESS"
        }
        return $true
    }
    catch {
        Write-Log "Error configuring Windows Defender: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Set-WindowsFirewall {
    param(
        [Parameter(Mandatory=$true)]
        [bool]$Disable
    )

    try {
        $profiles = @("Domain", "Private", "Public")
        $targetState = if ($Disable) { 'False' } else { 'True' }
        $statusWord = if ($Disable) { "Disabling" } else { "Enabling" }

        Write-Log "$statusWord Windows Firewall for all profiles..." -Level "INFO"

        foreach ($profile in $profiles) {
            Set-NetFirewallProfile -Name $profile -Enabled $targetState
            Write-Log "$statusWord firewall for $profile profile." -Level "INFO"
        }

        Write-Log "Windows Firewall configuration complete." -Level "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Error configuring Windows Firewall: $($_.Exception.Message)" -Level "ERROR"
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
function Install-PythonSilently {
    Write-Log "Checking for Python installation..." -Level "INFO"

    try {
        # Check if python is in PATH
        $pythonPath = Get-Command python -ErrorAction SilentlyContinue
        if ($pythonPath) {
            Write-Log "Python is already installed at: $($pythonPath.Source)" -Level "SUCCESS"
            return $true
        }

        Write-Log "Python not found. Starting silent installation..." -Level "INFO"

        $tempInstaller = Join-Path $Global:TempDir "python-installer.exe"
        $downloadUrl = "https://www.python.org/ftp/python/3.11.4/python-3.11.4-amd64.exe" # Specify a recent, stable version

        Write-Log "Downloading Python installer from $downloadUrl..." -Level "INFO"
        Invoke-WebRequest -Uri $downloadUrl -OutFile $tempInstaller -UseBasicParsing
        Write-Log "Python installer downloaded successfully." -Level "SUCCESS"

        Write-Log "Starting silent installation..." -Level "INFO"
        $installProcess = Start-Process -FilePath $tempInstaller -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait -PassThru

        if ($installProcess.ExitCode -eq 0) {
            Write-Log "Python installed successfully." -Level "SUCCESS"
            # Update PATH for the current session
            $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
            return $true
        } else {
            throw "Python installation failed with exit code: $($installProcess.ExitCode)"
        }
    }
    catch {
        Write-Log "Error installing Python: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

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

function Invoke-EdrInstallation {
    Write-Section "TEHTRIS EDR INSTALLATION"
    if (-not ([string]::IsNullOrEmpty($InstallEdrMsiPath))) {
        Invoke-TehtrisEdrInstaller -MsiPath $InstallEdrMsiPath
    } else {
        Write-Log "The -InstallEdrMsiPath parameter is required to install the EDR." -Level "ERROR"
        return $false
    }
}

function Invoke-EdrUninstallation {
    Write-Section "TEHTRIS EDR UNINSTALLATION"

    try {
        Write-Host "To uninstall TEHTRIS EDR, you need to provide either:"
        Write-Host "1. Uninstall password"
        Write-Host "2. Path to uninstall key file"
        Write-Host

        $choice = Read-Host -Prompt "Enter '1' for password or '2' for key file"

        if ($choice -eq '1') {
            $password = Read-Host -Prompt "Enter uninstall password"
            Invoke-TehtrisEdrUninstaller -Password $password
        } elseif ($choice -eq '2') {
            $keyFilePath = Read-Host -Prompt "Enter path to uninstall key file"
            Invoke-TehtrisEdrUninstaller -KeyFilePath $keyFilePath
        } else {
            Write-Log "Invalid choice. Please enter '1' or '2'." -Level "ERROR"
            return $false
        }

        return $true
    }
    catch {
        Write-Log "Error during EDR uninstallation: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}




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

    # Ensure Python is installed
    if (!(Install-PythonSilently)) {
        throw "Python installation failed. Cannot proceed."
    }

    Write-Log "Environment initialization completed." -Level "SUCCESS"
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

# Main Orchestration
Write-Section "STARTING SECURITY ENVIRONMENT SETUP"

try {
    Initialize-Environment

    if ($All) {
        Write-Log "'-All' switch specified. Running full setup..." -Level "INFO"
        Write-Section "SECURITY CONFIGURATION"
        Set-WindowsDefender -Disable $true
        Set-WindowsFirewall -Disable $true
        Invoke-ToolDeployment
        if ($PSBoundParameters.ContainsKey('InstallEdrMsiPath')) {
            Invoke-EdrInstallation
        } else {
            Write-Log "-InstallEdrMsiPath not provided with -All switch. Skipping EDR installation." -Level "INFO"
        }
    } else {
        # Individual actions
        if ($DisableSec.IsPresent) {
            Write-Log "Processing -DisableSec parameter..." -Level "INFO"
            Write-Section "SECURITY CONFIGURATION"
            Set-WindowsDefender -Disable $true
            Set-WindowsFirewall -Disable $true
        }
        if ($EnableSec.IsPresent) {
            Write-Log "Processing -EnableSec parameter..." -Level "INFO"
            Write-Section "SECURITY CONFIGURATION"
            Set-WindowsDefender -Disable $false
            Set-WindowsFirewall -Disable $false
        }
        if ($InstallTools.IsPresent) {
            Write-Log "Processing -InstallTools parameter..." -Level "INFO"
            Invoke-ToolDeployment
        }
        if ($PSBoundParameters.ContainsKey('InstallEdrMsiPath')) {
            Write-Log "Processing -InstallEdrMsiPath parameter..." -Level "INFO"
            Invoke-EdrInstallation
        }
        if ($UninstallEdr.IsPresent) {
            Write-Log "Processing -UninstallEdr parameter..." -Level "INFO"
            Invoke-EdrUninstallation
        }
    }

    Write-Log "Script actions completed." -Level "SUCCESS"
}
catch {
    Write-Log "An unexpected error occurred: $($_.Exception.Message)" -Level "ERROR"
}
finally {
    $scriptEndTime = Get-Date
    $scriptDuration = New-TimeSpan -Start $Global:ScriptStartTime -End $scriptEndTime
    Write-Section "SCRIPT EXECUTION SUMMARY"
    Write-Log "Script finished in: $($scriptDuration.TotalSeconds) seconds."
}
#endregion
