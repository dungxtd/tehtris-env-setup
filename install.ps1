# Bootstrapper for Security Testing Environment Setup
# USAGE: powershell "iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/dungxtd/tehtris-env-setup/master/install.ps1'))" -InstallEdrV1

[CmdletBinding(SupportsShouldProcess=$true)]
param ()

# --- Configuration ---
$RepoUrl = "https://github.com/dungxtd/tehtris-env-setup/archive/refs/heads/master.zip"
$TempDir = "C:\Temp"
$RepoZipPath = Join-Path $TempDir "tehtris-repo.zip"
$ExtractedDirName = "tehtris-env-setup-master" # Default dir name from GitHub zip
$ExtractedPath = Join-Path $TempDir $ExtractedDirName
$SetupScriptPath = Join-Path $ExtractedPath "Scripts\setup_env.ps1"

# --- Bootstrap Functions ---
function Write-BootstrapLog {
    param([string]$Message, [string]$Color = "White")
    Write-Host "[BOOTSTRAPPER] $Message" -ForegroundColor $Color
}

# --- Main Execution ---
try {
    # 1. Check for Administrator Privileges
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Administrator privileges are required. Please re-run PowerShell as an Administrator."
    }

    Write-BootstrapLog "Starting environment setup..." -Color Cyan

    # 2. Create Temp Directory
    if (-not (Test-Path $TempDir)) {
        Write-BootstrapLog "Creating temporary directory at $TempDir"
        New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
    }

    # 3. Download the Repository
    Write-BootstrapLog "Downloading repository from $RepoUrl..." -Color Yellow
    Invoke-WebRequest -Uri $RepoUrl -OutFile $RepoZipPath -UseBasicParsing
    Write-BootstrapLog "Repository downloaded successfully." -Color Green

    # 4. Extract the Repository
    Write-BootstrapLog "Extracting repository to $TempDir..." -Color Yellow
    Expand-Archive -Path $RepoZipPath -DestinationPath $TempDir -Force
    Write-BootstrapLog "Repository extracted successfully." -Color Green

    # 5. Verify Setup Script Exists
    if (-not (Test-Path $SetupScriptPath)) {
        throw "Setup script not found at $SetupScriptPath. The repository structure may have changed."
    }

    # 6. Execute the Main Setup Script
    Write-BootstrapLog "Executing main setup script: $SetupScriptPath" -Color Cyan
    if ($PSBoundParameters.Count -gt 0) {
        Write-BootstrapLog "Arguments passed: $($PSBoundParameters.Keys -join ', ')" -Color Gray
    }
    
    # Execute the script with all parameters passed to the bootstrapper
    & $SetupScriptPath @PSBoundParameters

}
catch {
    Write-BootstrapLog "An error occurred during the bootstrap process:" -Color Red
    Write-Error $_.Exception.Message
}

