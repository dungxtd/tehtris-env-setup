#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Security Testing Environment Setup Script
.DESCRIPTION
    Automates the setup of a security testing environment with proper logging and error handling.
    Temporarily disables security features, installs tools, and provides cleanup functionality.
.PARAMETER InstallEdrV1
    Install TEHTRIS EDR V1 using default installer name: TEHTRIS_EDR_1.8.1_rc2.exe
.PARAMETER InstallEdrV2
    Install TEHTRIS EDR V2 using default installer name: TEHTRIS_EDR_2.0.0_Windows_x86_64_MS-28.msi
.PARAMETER InstallEdrPath
    Install TEHTRIS EDR using a custom installer path
.EXAMPLE
    .\setup_env.ps1 -InstallEdrV1
    Installs TEHTRIS EDR V1 using the default V1 installer
.EXAMPLE
    .\setup_env.ps1 -InstallEdrV2 -UninstallEdrPassword "password123"
    Installs TEHTRIS EDR V2 and provides uninstall password for removing existing installations
.NOTES
    Author: Cascade AI Assistant
    Created: 2025-08-25
    Updated: 2025-08-27 (Added V1/V2 default parameters)
    Requires: Administrator privileges

    Default installer locations (in Tools directory):
    - V1: TEHTRIS_EDR_1.8.1_rc2.exe
    - V2: TEHTRIS_EDR_2.0.0_Windows_x86_64_MS-28.msi
#>

param(
    [Parameter(Mandatory=$false)]
    [switch]$DisableSec,

    [Parameter(Mandatory=$false)]
    [switch]$EnableSec,

    [Parameter(Mandatory=$false)]
    [switch]$InstallTools,

    [Parameter(Mandatory=$false)]
    [string]$InstallEdrPath,

    [Parameter(Mandatory=$false)]
    [switch]$InstallEdrV1,

    [Parameter(Mandatory=$false)]
    [switch]$InstallEdrV2,

    [Parameter(Mandatory=$false)]
    [string]$UninstallEdrPassword,

    [Parameter(Mandatory=$false)]
    [string]$UninstallEdrKeyFile,

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
        # Check all available python commands for a real installation
        $pythonCommands = Get-Command python -All -ErrorAction SilentlyContinue
        foreach ($pythonCmd in $pythonCommands) {
            if ($pythonCmd.Source -notlike '*\WindowsApps\*') {
                Write-Log "Valid Python installation found at: $($pythonCmd.Source)" -Level "SUCCESS"
                return $true
            }
        }

        Write-Log "Python not found. Starting silent installation..." -Level "INFO"

        $tempInstaller = Join-Path $Global:TempDir "python-installer.exe"
        $downloadUrl = "https://www.python.org/ftp/python/3.11.4/python-3.11.4-amd64.exe" # Specify a recent, stable version

        Write-Log "Downloading Python installer from $downloadUrl..." -Level "INFO"
        (New-Object System.Net.WebClient).DownloadFile($downloadUrl, $tempInstaller)
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
            (New-Object System.Net.WebClient).DownloadFile($downloadUrl, $tempZip)
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
            $nmapPageContent = (New-Object System.Net.WebClient).DownloadString("https://nmap.org/download.html")
            $installerPattern = 'https://nmap\.org/dist/nmap-[\d\.]+-setup\.exe'
            $downloadUrl = [regex]::Match($nmapPageContent, $installerPattern).Value

            if (!$downloadUrl) {
                throw "Could not find Nmap installer download URL"
            }

            (New-Object System.Net.WebClient).DownloadFile($downloadUrl, $tempInstaller)
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

function Install-PythonRequirements {
    Write-Log "Installing Python requirements from requirements.txt..." -Level "INFO"

    try {
        $requirementsFile = Join-Path $Global:BaseDir "requirements.txt"
        if (!(Test-Path $requirementsFile)) {
            throw "requirements.txt not found at $requirementsFile"
        }

        $process = Start-Process -FilePath "python" -ArgumentList "-m pip install -r `"$requirementsFile`"" -Wait -PassThru -NoNewWindow

        if ($process.ExitCode -eq 0) {
            Write-Log "Python requirements installed successfully." -Level "SUCCESS"
            return $true
        } else {
            throw "pip install failed with exit code: $($process.ExitCode)"
        }
    }
    catch {
        Write-Log "Error installing Python requirements: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}


#region TEHTRIS EDR Integration Functions
function Get-EdrVersionFromPath {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )

    $fileName = Split-Path $FilePath -Leaf
    Write-Log "Detecting EDR version from filename: $fileName" -Level "INFO"

    # Extract version using regex pattern (e.g., 1.8.1, 2.0.0)
    if ($fileName -match '(\d+)\.(\d+)\.(\d+)') {
        $version = $matches[0]
        Write-Log "Detected specific version: $version" -Level "INFO"
        return $version
    }

    # Fallback: detect major version
    if ($fileName -match '[_\-]1[\._\-]') {
        Write-Log "Detected version 1.x.x" -Level "INFO"
        return "1.x.x"
    } elseif ($fileName -match '[_\-]2[\._\-]') {
        Write-Log "Detected version 2.x.x" -Level "INFO"
        return "2.x.x"
    }

    # Default to 2.x.x if cannot detect
    Write-Log "Could not detect version, defaulting to 2.x.x" -Level "WARN"
    return "2.x.x"
}

function Get-DefaultEdrPath {
    param(
        [Parameter(Mandatory=$false)]
        [string]$Version = "auto"
    )

    Write-Log "Determining default EDR installer path for version: $Version" -Level "INFO"

    # Define default installer names
    $v1DefaultName = "TEHTRIS_EDR_1.8.1_rc2.exe"
    $v2DefaultName = "TEHTRIS_EDR_2.0.0_Windows_x86_64_MS-28.msi"

    $v1Path = Join-Path $Global:ToolsDir $v1DefaultName
    $v2Path = Join-Path $Global:ToolsDir $v2DefaultName

    if ($Version -eq "v1" -or $Version -eq "1") {
        Write-Log "[V1 DEFAULT] Using V1 installer: $v1DefaultName" -Level "INFO"
        return $v1Path
    }
    elseif ($Version -eq "v2" -or $Version -eq "2") {
        Write-Log "[V2 DEFAULT] Using V2 installer: $v2DefaultName" -Level "INFO"
        return $v2Path
    }
    else {
        # Auto-detect: Check for version 1.8.1 first
        if (Test-Path $v1Path) {
            Write-Log "[AUTO-DETECT] Found V1 installer: $v1Path" -Level "INFO"
            return $v1Path
        }

        # Check for version 2.0.0
        if (Test-Path $v2Path) {
            Write-Log "[AUTO-DETECT] Found V2 installer: $v2Path" -Level "INFO"
            return $v2Path
        }

        # Return default v2 path even if file doesn't exist (for error handling)
        Write-Log "[AUTO-DETECT] No EDR installer found, returning default V2 path: $v2DefaultName" -Level "WARN"
        return $v2Path
    }
}

function Get-DefaultEdrV1Path {
    Write-Log "[V1] Getting default V1 installer path..." -Level "INFO"
    return Get-DefaultEdrPath -Version "v1"
}

function Get-DefaultEdrV2Path {
    Write-Log "[V2] Getting default V2 installer path..." -Level "INFO"
    return Get-DefaultEdrPath -Version "v2"
}

function Invoke-TehtrisEdrInstaller {
    param(
        [Parameter(Mandatory=$true)]
        [string]$InstallerPath,
        [Parameter(Mandatory=$false)]
        [string]$UninstallPassword,
        [Parameter(Mandatory=$false)]
        [string]$UninstallKeyFile
    )

    Write-Log "Invoking TEHTRIS EDR installer..." -Level "INFO"

    # Detect and log EDR version
    $edrVersion = Get-EdrVersionFromPath -FilePath $InstallerPath
    Write-Log "EDR Version detected: $edrVersion" -Level "INFO"

    try {
        if (!(Test-Path $InstallerPath)) {
            throw "Installer file not found at path: $InstallerPath"
        }

        $pythonScript = Join-Path $Global:ScriptsDir "tehtris_edr_installer_minimal.py"
        if (!(Test-Path $pythonScript)) {
            throw "TEHTRIS EDR installer script not found at: $pythonScript"
        }

        $arguments = "`"$InstallerPath`""
        if ($UninstallPassword) {
            $arguments += " --uninstall-password `"$UninstallPassword`""
        }
        if ($UninstallKeyFile) {
            $arguments += " --uninstall-key-file `"$UninstallKeyFile`""
        }

        Write-Log "Executing TEHTRIS EDR installer script with arguments: $arguments" -Level "INFO"
        $process = Start-Process -FilePath "python" -ArgumentList "`"$pythonScript`" $arguments" -Wait -PassThru -NoNewWindow

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
    param(
        [Parameter(Mandatory=$true)]
        [string]$InstallerPath,
        [Parameter(Mandatory=$false)]
        [string]$UninstallPassword,
        [Parameter(Mandatory=$false)]
        [string]$UninstallKeyFile
    )

    Write-Section "TEHTRIS EDR INSTALLATION"

    $installerArgs = @{
        InstallerPath = $InstallerPath
    }
    if ($UninstallPassword) {
        $installerArgs['UninstallPassword'] = $UninstallPassword
    }
    if ($UninstallKeyFile) {
        $installerArgs['UninstallKeyFile'] = $UninstallKeyFile
    }
    Invoke-TehtrisEdrInstaller @installerArgs
}

function Invoke-EdrUninstallation {
    param(
        [Parameter(Mandatory=$false)]
        [string]$Password,
        [Parameter(Mandatory=$false)]
        [string]$KeyFile
    )

    Write-Section "TEHTRIS EDR UNINSTALLATION"

    if ($Password) {
        Invoke-TehtrisEdrUninstaller -Password $Password
    } elseif ($KeyFile) {
        Invoke-TehtrisEdrUninstaller -KeyFilePath $KeyFile
    } else {
        Write-Log "Either a password or a key file must be provided for uninstallation." -Level "ERROR"
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

        # Install Python requirements
    if (!(Install-PythonRequirements)) {
        throw "Python requirements installation failed. Cannot proceed."
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

    # If no parameters are provided, default to -All
    if ($PSBoundParameters.Count -eq 0) {
        Write-Log "No parameters provided. Defaulting to '-All' for a full setup." -Level "INFO"
        $All = $true
    }

    if ($All) {
        Write-Log "'-All' switch specified. Running full setup (EDR installation skipped)..." -Level "INFO"
        Write-Section "SECURITY CONFIGURATION"
        Set-WindowsDefender -Disable $true
        Set-WindowsFirewall -Disable $true
        Invoke-ToolDeployment
        Write-Log "EDR installation skipped with -All parameter. Use specific EDR parameters to install EDR." -Level "INFO"
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
        if ($PSBoundParameters.ContainsKey('InstallEdrPath')) {
            Write-Log "Processing -InstallEdrPath parameter..." -Level "INFO"
            Invoke-EdrInstallation -InstallerPath $InstallEdrPath -UninstallPassword $UninstallEdrPassword -UninstallKeyFile $UninstallEdrKeyFile
        }
        if ($InstallEdrV1.IsPresent) {
            Write-Log "Processing -InstallEdrV1 parameter..." -Level "INFO"
            $v1Path = Get-DefaultEdrV1Path
            Invoke-EdrInstallation -InstallerPath $v1Path -UninstallPassword $UninstallEdrPassword -UninstallKeyFile $UninstallEdrKeyFile
        }
        if ($InstallEdrV2.IsPresent) {
            Write-Log "Processing -InstallEdrV2 parameter..." -Level "INFO"
            $v2Path = Get-DefaultEdrV2Path
            Invoke-EdrInstallation -InstallerPath $v2Path -UninstallPassword $UninstallEdrPassword -UninstallKeyFile $UninstallEdrKeyFile
        }
        if ($PSBoundParameters.ContainsKey('UninstallEdrPassword') -or $PSBoundParameters.ContainsKey('UninstallEdrKeyFile')) {
            if (-not $PSBoundParameters.ContainsKey('InstallEdrPath') -and -not $InstallEdrV1.IsPresent -and -not $InstallEdrV2.IsPresent) {
                Invoke-EdrUninstallation -Password $UninstallEdrPassword -KeyFile $UninstallEdrKeyFile
            }
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
