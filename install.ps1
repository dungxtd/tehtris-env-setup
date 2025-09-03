# Bootstrapper for Security Testing Environment Setup
# USAGE: powershell "iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/dungxtd/tehtris-env-setup/master/install.ps1'))" -InstallEdrV1

[CmdletBinding(SupportsShouldProcess=$true)]
param (
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

# --- Configuration ---
$ApiUrl = "https://api.github.com/repos/dungxtd/tehtris-env-setup/releases/latest"
$TempDir = "C:\Temp"
$RepoZipPath = Join-Path $TempDir "tehtris-env-setup.zip"
$ExtractedPath = Join-Path $TempDir "tehtris-env-setup"
$SetupScriptPath = Join-Path $ExtractedPath "Scripts\setup_env.ps1"
$VersionFilePath = Join-Path $ExtractedPath "version.txt"

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

    # 3. Check for Updates and Download if Necessary
    Write-BootstrapLog "Fetching latest release information from GitHub..." -Color Yellow
    $release = Invoke-RestMethod -Uri $ApiUrl -UseBasicParsing
    $latestVersion = $release.tag_name

    $localVersion = ""
    if (Test-Path $VersionFilePath) {
        $localVersion = Get-Content $VersionFilePath
    }

    if ($localVersion -eq $latestVersion) {
        Write-BootstrapLog "You already have the latest version ($latestVersion). Skipping download." -Color Green
    } else {
        Write-BootstrapLog "New version ($latestVersion) found. Local version is '$localVersion'." -Color Yellow
        $downloadUrl = $release.assets | Where-Object { $_.name -eq "tehtris-env-setup.zip" } | Select-Object -ExpandProperty browser_download_url

        if (-not $downloadUrl) {
            throw "Could not find 'tehtris-env-setup.zip' in the latest release."
        }

        Write-BootstrapLog "Downloading release asset from $downloadUrl..." -Color Yellow
        (New-Object System.Net.WebClient).DownloadFile($downloadUrl, $RepoZipPath)
        Write-BootstrapLog "Release downloaded successfully." -Color Green

        # 4. Extract the Repository
        Write-BootstrapLog "Extracting repository to $ExtractedPath..." -Color Yellow
        if (Test-Path $ExtractedPath) {
            Remove-Item -Recurse -Force $ExtractedPath
        }
        Expand-Archive -Path $RepoZipPath -DestinationPath $ExtractedPath -Force
        Set-Content -Path $VersionFilePath -Value $latestVersion
        Write-BootstrapLog "Repository extracted and version updated to $latestVersion." -Color Green
    }

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

    # 7. Post-flight check for Python PATH
    $pythonInPath = Get-Command python -ErrorAction SilentlyContinue
    if (-not ($pythonInPath -and $pythonInPath.Source -notlike '*\WindowsApps\*')) {
        $pythonInstallDir = Join-Path $env:ProgramFiles "Python311"
        if (Test-Path (Join-Path $pythonInstallDir "python.exe")) {
            Write-BootstrapLog "Python was installed. Please RESTART your terminal and run the command again to ensure all tools function correctly." -Color Magenta
        }
    }

}
catch {
    Write-BootstrapLog "An error occurred during the bootstrap process:" -Color Red
    Write-Error $_.Exception.Message
}

