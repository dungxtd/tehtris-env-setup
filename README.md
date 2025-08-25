# Security Testing Environment Setup

This repository contains a set of scripts designed to automate the setup of a security testing environment on a Windows machine. The scripts handle tasks such as temporarily disabling security features, installing common security tools, and automating the installation of TEHTRIS EDR.

## Main Script

-   `Scripts/setup_env.ps1`: The main PowerShell script that orchestrates the entire environment setup. It can temporarily disable Windows Defender and Firewall, install tools like Mimikatz and Nmap, and integrate with the TEHTRIS EDR installer. This script automatically handles all necessary Python scripts internally.

## Prerequisites

1.  **Operating System**: Windows 10 or later.
2.  **Administrator Privileges**: All scripts must be run from an elevated (Administrator) terminal.
3.  **PowerShell**: PowerShell 5.1 or later.
4.  **Python**: Python 3.8 or later, with `python` available in the system's PATH.
5.  **Python Packages**: The Python scripts require the following packages. You can install them using pip:
    ```sh
    pip install pyautogui psutil pywin32
    ```

## Usage

The `setup_env.ps1` script is designed to be modular. You can perform specific actions by using different switches. It must be run from an administrative PowerShell terminal.

**Parameters**:

*   `-DisableSec <$true | $false>`: Controls the security features (Windows Defender and Firewall).
    *   `$true`: Disables the security features.
    *   `$false` (default): Enables the security features.
*   `-InstallTools`: A switch to install the security tools (Mimikatz and Nmap).
*   `-InstallEdrMsiPath <path>`: Specifies the path to the TEHTRIS EDR MSI file to begin the installation.
*   `-UninstallEdr`: A switch to start the EDR uninstallation process.
*   `-All`: A switch to perform a full setup. This will disable security features, install all tools, and install the EDR if the `-InstallEdrMsiPath` is also provided.

**Examples**:

1.  **Full Setup (Disable Security, Install Tools, and Install EDR)**:
    ```powershell
    powershell.exe -ExecutionPolicy Bypass -File .\setup_env.ps1 -All -InstallEdrMsiPath "C:\path\to\tehtris.msi"
    ```

2.  **Disable Security Features Only**:
    ```powershell
    powershell.exe -ExecutionPolicy Bypass -File .\setup_env.ps1 -DisableSec $true
    ```

3.  **Enable Security Features Only**:
    ```powershell
    powershell.exe -ExecutionPolicy Bypass -File .\setup_env.ps1 -DisableSec $false
    ```

4.  **Install Tools Only**:
    ```powershell
    powershell.exe -ExecutionPolicy Bypass -File .\setup_env.ps1 -InstallTools
    ```

5.  **Install TEHTRIS EDR Only**:
    ```powershell
    powershell.exe -ExecutionPolicy Bypass -File .\setup_env.ps1 -InstallEdrMsiPath "C:\path\to\tehtris.msi"
    ```

6.  **Uninstall TEHTRIS EDR**:
    ```powershell
    powershell.exe -ExecutionPolicy Bypass -File .\setup_env.ps1 -UninstallEdr
    ```

## Logging

All operations are logged to files located in the `Logs` directory, which is created at the root of the project folder.

-   `setup_env.log`: Logs from the main PowerShell script.
-   `tehtris_installation.log`: Logs from the TEHTRIS EDR installer script.
