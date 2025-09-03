# Security Testing Environment Setup

This repository provides a powerful one-line command to automate the setup of a security testing environment on any Windows machine. It handles downloading the project, disabling security features, and installing essential tools, all from a single command.

## ⚠️ Critical Prerequisite: Disable Tamper Protection

Before running the command, you **must** manually disable **Windows Defender Tamper Protection**.

1.  Open **Windows Security**.
2.  Go to **Virus & threat protection** > **Virus & threat protection settings**.
3.  Turn off **Tamper Protection**.

This step is mandatory because Tamper Protection is designed by Microsoft to prevent scripts and unauthorized applications from disabling security features. It cannot be disabled programmatically.

## Installation

Run the following command from any **administrative terminal** (PowerShell, Command Prompt, or bash).

```powershell
powershell.exe -ExecutionPolicy Bypass -Command "& {iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/dungxtd/tehtris-env-setup/master/install.ps1'))}"
```

This command intelligently handles updates. If you have a previous version, it will only download the new release if one is available.

## Usage with Parameters

To pass parameters, simply add them to the end of the command.

**Install EDR V1**
```powershell
powershell.exe -ExecutionPolicy Bypass -Command "& {iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/dungxtd/tehtris-env-setup/master/install.ps1'))} -InstallEdrV1"
```

**Install EDR V2** (replace `'your_password'` with the actual password)
```powershell
powershell.exe -ExecutionPolicy Bypass -Command "& {iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/dungxtd/tehtris-env-setup/master/install.ps1'))} -InstallEdrV2 -UninstallEdrPassword 'your_password'"
```

## Logging

Logs are stored in the `Logs` directory with detailed version-specific information.
