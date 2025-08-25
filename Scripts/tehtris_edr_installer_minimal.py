#!/usr/bin/env python3
"""
TEHTRIS EDR MSI Installer Automation Script
"""

import os
import sys
import time
import logging
import subprocess
import argparse
from pathlib import Path

try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.5
except ImportError:
    PYAUTOGUI_AVAILABLE = False

class TehtrisEDRInstaller:
    """Minimal TEHTRIS EDR installer automation."""

    def __init__(self, msi_path: str, uninstall_password: str = None, uninstall_key_file: str = None):
        self.msi_path = Path(msi_path)
        self.uninstall_password = uninstall_password
        self.uninstall_key_file = uninstall_key_file
        self.logger = self._setup_logging()

        # Configuration
        self.server_address = "xpgapp16.tehtris.net"
        self.tag = "XPG_QAT"
        self.license_key = "MH83-2CDX-9DXQ-LG89-92FF"

    def _setup_logging(self) -> logging.Logger:
        """Setup logging."""
        logger = logging.getLogger('TehtrisEDRInstaller')
        logger.setLevel(logging.INFO)

        # File handler
        file_handler = logging.FileHandler('tehtris_installation.log')
        file_handler.setLevel(logging.INFO)
        file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(file_formatter)

        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_formatter = logging.Formatter('%(levelname)s - %(message)s')
        console_handler.setFormatter(console_formatter)

        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

        return logger

    def validate_prerequisites(self) -> bool:
        """Validate prerequisites."""
        self.logger.info("Validating prerequisites...")

        if not self.msi_path.exists():
            self.logger.error(f"MSI file not found: {self.msi_path}")
            return False

        if not self._is_admin():
            self.logger.error("Administrator privileges required")
            return False

        self.logger.info("Prerequisites validated successfully")
        return True

    def _is_admin(self) -> bool:
        """Check admin privileges."""
        try:
            import ctypes
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def click_with_win32gui(self, button_text: str) -> bool:
        """Click button using win32gui."""
        try:
            import win32gui
            import win32con

            self.logger.info(f"Looking for button: {button_text}")

            # Find TEHTRIS windows
            tehtris_windows = []
            def find_windows(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        if "TEHTRIS EDR Setup" in window_text:
                            windows.append(hwnd)
                except:
                    pass
                return True

            win32gui.EnumWindows(find_windows, tehtris_windows)
            self.logger.info(f"Found {len(tehtris_windows)} TEHTRIS windows")

            # Search for button
            for tehtris_hwnd in tehtris_windows:
                try:
                    def find_button(hwnd, button_info):
                        try:
                            if win32gui.IsWindowVisible(hwnd):
                                window_text = win32gui.GetWindowText(hwnd)
                                class_name = win32gui.GetClassName(hwnd)

                                if window_text and class_name == 'Button':
                                    clean_text = window_text.replace('&', '').lower()
                                    clean_button = button_text.replace('&', '').lower()

                                    if clean_button in clean_text:
                                        button_info['hwnd'] = hwnd
                                        button_info['text'] = window_text
                                        return False
                        except:
                            pass
                        return True

                    button_info = {}
                    win32gui.EnumChildWindows(tehtris_hwnd, find_button, button_info)

                    if button_info.get('hwnd'):
                        win32gui.SendMessage(button_info['hwnd'], win32con.WM_LBUTTONDOWN, 0, 0)
                        win32gui.SendMessage(button_info['hwnd'], win32con.WM_LBUTTONUP, 0, 0)
                        self.logger.info(f"Clicked button: {button_info['text']}")
                        return True

                except Exception as e:
                    self.logger.debug(f"Error in window {tehtris_hwnd}: {e}")
                    continue

            self.logger.error(f"Button '{button_text}' not found")
            return False

        except Exception as e:
            self.logger.error(f"win32gui click failed: {e}")
            return False

    def fill_field_with_win32gui(self, field_label: str, value: str) -> bool:
        """Fill field using win32gui."""
        try:
            import win32gui
            import win32con

            self.logger.info(f"Looking for field: {field_label}")

            # Find TEHTRIS windows
            tehtris_windows = []
            def find_windows(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        if "TEHTRIS EDR Setup" in window_text:
                            windows.append(hwnd)
                except:
                    pass
                return True

            win32gui.EnumWindows(find_windows, tehtris_windows)

            # Field mapping
            field_mapping = {
                'server': 0,
                'tag': 1,
                'license': 2
            }

            field_index = field_mapping.get(field_label.lower())
            if field_index is None:
                self.logger.error(f"Unknown field: {field_label}")
                return False

            # Find edit controls
            for tehtris_hwnd in tehtris_windows:
                try:
                    edit_controls = []
                    def find_edits(hwnd, controls):
                        try:
                            if win32gui.IsWindowVisible(hwnd):
                                class_name = win32gui.GetClassName(hwnd)
                                if class_name in ['Edit', 'RichEdit20W']:
                                    controls.append(hwnd)
                        except:
                            pass
                        return True

                    win32gui.EnumChildWindows(tehtris_hwnd, find_edits, edit_controls)

                    if field_index < len(edit_controls):
                        edit_hwnd = edit_controls[field_index]
                        # Click on field to set focus
                        rect = win32gui.GetWindowRect(edit_hwnd)
                        center_x = (rect[0] + rect[2]) // 2
                        center_y = (rect[1] + rect[3]) // 2

                        if PYAUTOGUI_AVAILABLE:
                            pyautogui.click(center_x, center_y)
                            time.sleep(0.2)
                        win32gui.SendMessage(edit_hwnd, win32con.WM_SETTEXT, 0, "")
                        time.sleep(0.1)
                        win32gui.SendMessage(edit_hwnd, win32con.WM_SETTEXT, 0, value)
                        time.sleep(0.2)

                        # Send Tab to trigger validation
                        win32gui.SendMessage(edit_hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
                        win32gui.SendMessage(edit_hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)

                        self.logger.info(f"Filled {field_label} with '{value}'")
                        return True

                except Exception as e:
                    self.logger.debug(f"Error in window {tehtris_hwnd}: {e}")
                    continue

            self.logger.error(f"Field '{field_label}' not found")
            return False

        except Exception as e:
            self.logger.error(f"Fill field failed: {e}")
            return False

    def launch_installer(self) -> bool:
        """Launch MSI installer."""
        self.logger.info("Step 1: Launching installer...")

        try:
            # Minimize windows
            # if PYAUTOGUI_AVAILABLE:
            #     pyautogui.hotkey('win', 'd')
            #     time.sleep(1.5)

            # Launch installer
            subprocess.Popen(['msiexec', '/i', str(self.msi_path)], shell=True)
            time.sleep(5)

            self.logger.info("Installer launched successfully")
            return True

        except Exception as e:
            self.logger.error(f"Failed to launch installer: {e}")
            return False

    def handle_welcome_screen(self) -> bool:
        """Handle welcome screen by repeatedly trying to click 'Next'."""
        self.logger.info("Step 2: Handling welcome screen...")

        timeout = 30  # 30 seconds
        start_time = time.time()

        while time.time() - start_time < timeout:
            if self.click_with_win32gui("Next"):
                self.logger.info("Successfully clicked 'Next' on the welcome screen.")
                return True
            self.logger.info("Welcome screen not ready yet, retrying in 2 seconds...")
            time.sleep(2)

        self.logger.error("Failed to handle welcome screen within timeout.")
        return False

    def handle_license_agreement(self) -> bool:
        """Handle license agreement."""
        self.logger.info("Step 3: Handling license agreement...")
        time.sleep(0.5)

        if not self.click_with_win32gui("accept"):
            return False
        time.sleep(0.5)
        return self.click_with_win32gui("Next")

    def handle_activation_information(self) -> bool:
        """Handle activation information."""
        self.logger.info("Step 4: Handling activation information...")
        time.sleep(0.5)

        # Fill fields
        if not self.fill_field_with_win32gui("server", self.server_address):
            return False
        if not self.fill_field_with_win32gui("tag", self.tag):
            return False
        if not self.fill_field_with_win32gui("license", self.license_key):
            return False

        time.sleep(1)
        return self.click_with_win32gui("Next")

    def handle_installation(self) -> bool:
        """Handle installation."""
        self.logger.info("Step 5: Handling installation...")
        time.sleep(0.3)

        # Try clicking Install button or use Alt+I
        if not self.click_with_win32gui("Install"):
            if PYAUTOGUI_AVAILABLE:
                pyautogui.hotkey('alt', 'i')
                time.sleep(1)

        return True

    def wait_for_completion(self) -> bool:
        """Wait for completion."""
        self.logger.info("Step 6: Waiting for installation completion...")

        completion_timeout = 180  # 3 minutes
        start_time = time.time()

        while time.time() - start_time < completion_timeout:
            elapsed = int(time.time() - start_time)
            self.logger.info(f"Checking for completion... ({elapsed}s elapsed)")

            if self.click_with_win32gui("Finish"):
                self.logger.info("Clicked Finish button")
                return True
            elif self.click_with_win32gui("Close"):
                self.logger.info("Clicked Close button")
                return True

            time.sleep(2)

        # Final attempt with Alt+F
        if PYAUTOGUI_AVAILABLE:
            self.logger.info("Trying Alt+F as final attempt")
            pyautogui.hotkey('alt', 'f')
            time.sleep(1)
            return True

        self.logger.error("Installation did not complete within timeout")
        return False

    def verify_installation(self, post_install_check: bool = False) -> bool:
        """Verify installation by checking for TEHTRIS Agent processes."""
        if post_install_check:
            self.logger.info("Step 7: Verifying installation...")
        else:
            self.logger.info("Checking for existing TEHTRIS EDR installation...")

        try:
            import psutil

            self.logger.info("Checking for Agent processes with TEHTRIS description...")
            tehtris_agents = []

            for proc in psutil.process_iter(['pid', 'name', 'exe']):
                try:
                    proc_info = proc.info
                    proc_name = proc_info['name'].lower()

                    if 'agent' in proc_name:
                        exe_path = proc_info.get('exe', '')
                        if exe_path and 'tehtris' in exe_path.lower():
                            tehtris_agents.append({
                                'pid': proc_info['pid'],
                                'name': proc_info['name'],
                                'exe': exe_path
                            })
                            self.logger.info(f"[FOUND] TEHTRIS Agent: PID {proc_info['pid']} - {proc_info['name']} - {exe_path}")

                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue

            if tehtris_agents:
                if post_install_check:
                    self.logger.info(f"[SUCCESS] TEHTRIS EDR verified - Found {len(tehtris_agents)} Agent process(es)")
                else:
                    self.logger.warning("TEHTRIS EDR installation detected!")
                return True
            else:
                if post_install_check:
                    self.logger.warning("[NOT FOUND] No TEHTRIS Agent processes found")
                    self.logger.warning("Installation may have completed but processes haven't started yet")
                else:
                    self.logger.info("No existing TEHTRIS EDR installation found.")
                return False

        except Exception as e:
            self.logger.warning(f"Verification failed: {e}")
            return post_install_check if post_install_check else False

    def uninstall_existing_edr(self, credential_type: str, credential_value: str) -> bool:
        """Uninstall existing TEHTRIS EDR."""
        self.logger.info("Starting TEHTRIS EDR uninstallation...")

        try:
            # Find the uninstaller script
            uninstaller_script = os.path.join(os.path.dirname(__file__), "tehtris_edr_uninstaller.py")
            if not os.path.exists(uninstaller_script):
                self.logger.error(f"Uninstaller script not found: {uninstaller_script}")
                return False

            # Prepare command arguments
            cmd = [sys.executable, uninstaller_script]
            if credential_type == 'password':
                cmd.extend(['--password', credential_value])
            elif credential_type == 'keyfile':
                cmd.extend(['--keyfile', credential_value])

            self.logger.info("Executing uninstaller...")
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)

            if result.returncode == 0:
                self.logger.info("TEHTRIS EDR uninstalled successfully.")
                # Wait a bit for cleanup
                time.sleep(5)
                return True
            else:
                self.logger.error(f"Uninstaller failed with exit code {result.returncode}")
                self.logger.error(f"Error output: {result.stderr}")
                return False

        except subprocess.TimeoutExpired:
            self.logger.error("Uninstaller timed out after 5 minutes.")
            return False
        except Exception as e:
            self.logger.error(f"Error during uninstallation: {e}")
            return False

    def run_installation(self) -> bool:
        """Run complete installation."""
        self.logger.info("Starting TEHTRIS EDR installation automation")

        try:
            # Check for existing installation first
            if self.verify_installation(post_install_check=False):
                self.logger.warning("Existing TEHTRIS EDR installation detected. Attempting to uninstall...")
                if not self.uninstall_password and not self.uninstall_key_file:
                    self.logger.error("Existing installation found, but no uninstall credentials provided.")
                    return False

                credential_type = 'password' if self.uninstall_password else 'keyfile'
                credential_value = self.uninstall_password if self.uninstall_password else self.uninstall_key_file

                if not self.uninstall_existing_edr(credential_type, credential_value):
                    self.logger.error("Failed to uninstall existing TEHTRIS EDR. Cannot proceed with installation.")
                    return False

                self.logger.info("Existing installation removed. Proceeding with new installation...")

            if not self.validate_prerequisites():
                return False

            if not self.launch_installer():
                return False

            if not self.handle_welcome_screen():
                return False

            if not self.handle_license_agreement():
                return False

            if not self.handle_activation_information():
                return False

            if not self.handle_installation():
                return False

            if not self.wait_for_completion():
                return False

            if not self.verify_installation(post_install_check=True):
                return False

            self.logger.info("TEHTRIS EDR installation completed successfully!")
            return True

        except Exception as e:
            self.logger.error(f"Installation failed: {e}")
            return False

def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(description="TEHTRIS EDR Installer")
    parser.add_argument("msi_path", help="Path to the TEHTRIS EDR MSI installer")
    parser.add_argument("--uninstall-password", help="Password for uninstalling a previous version")
    parser.add_argument("--uninstall-key-file", help="Key file for uninstalling a previous version")

    args = parser.parse_args()

    installer = TehtrisEDRInstaller(
        msi_path=args.msi_path,
        uninstall_password=args.uninstall_password,
        uninstall_key_file=args.uninstall_key_file
    )

    success = installer.run_installation()
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()
