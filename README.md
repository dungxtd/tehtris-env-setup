# Security Testing Environment Setup

This repository automates the setup of a security testing environment on Windows, including disabling security features, installing tools, and installing TEHTRIS EDR.

## Prerequisites

*   Windows 10 or later
*   Administrator privileges
*   PowerShell 5.1 or later
*   Python 3.8 or later (install required packages with `pip install -r requirement.txt`)

## Usage

Run all scripts from an administrative PowerShell terminal.

**Full Setup**
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Scripts\setup_env.ps1 -All
```

**Install EDR V1 (Default)**
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Scripts\setup_env.ps1 -InstallEdrV1
```

**Install EDR V2 (Default)**
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Scripts\setup_env.ps1 -InstallEdrV2 -UninstallEdrPassword "your_password"
```

**Install EDR (Custom Path)**
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Scripts\setup_env.ps1 -InstallEdrPath "C:\path\to\tehtris.exe" -UninstallEdrPassword "your_password"
```

**Uninstall EDR Only**
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Scripts\setup_env.ps1 -UninstallEdrPassword "your_password"
```

## Logging

Logs are stored in the `Logs` directory with detailed version-specific information.
