#!/usr/bin/env python3
"""
TEHTRIS EDR Uninstaller Automation Script
"""

import os
import sys
import time
import logging
import subprocess
import re
from pathlib import Path

try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.5
except ImportError:
    PYAUTOGUI_AVAILABLE = False

class TehtrisEDRUninstaller:
    """TEHTRIS EDR uninstaller automation."""

    def __init__(self, password: str = None, key_file: str = None):
        self.password = password
        self.key_file = Path(key_file) if key_file else None
        self.logger = self._setup_logging()
        
        # Detect EDR version from installed software
        self.edr_version = self._detect_installed_edr_version()
        self.logger.info(f"[DETECTED] EDR version: {self.edr_version}")

    def _setup_logging(self) -> logging.Logger:
        """Setup logging."""
        logger = logging.getLogger('TehtrisEDRUninstaller')
        logger.setLevel(logging.INFO)

        # File handler
        file_handler = logging.FileHandler('tehtris_uninstallation.log', mode='w')
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

    def _detect_installed_edr_version(self) -> str:
        """Detect installed EDR version from registry."""
        try:
            import winreg
            
            uninstall_key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            
            # Try 64-bit and 32-bit registry views
            for access_right in [winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY]:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, uninstall_key_path, 0, winreg.KEY_READ | access_right) as key:
                        for i in range(winreg.QueryInfoKey(key)[0]):
                            subkey_name = winreg.EnumKey(key, i)
                            with winreg.OpenKey(key, subkey_name) as subkey:
                                try:
                                    display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                                    if "tehtris edr" in display_name.lower():
                                        # Try to get version from DisplayVersion
                                        try:
                                            version = winreg.QueryValueEx(subkey, "DisplayVersion")[0]
                                            # Extract version using regex pattern
                                            version_pattern = r'(\d+)\.(\d+)\.(\d+)'
                                            match = re.search(version_pattern, version)
                                            if match:
                                                return f"{match.group(1)}.{match.group(2)}.{match.group(3)}"
                                            elif version.startswith('1.'):
                                                return "1.x.x"
                                            elif version.startswith('2.'):
                                                return "2.x.x"
                                        except OSError:
                                            pass
                                        
                                        # Fallback: detect from display name
                                        if "1." in display_name or "v1" in display_name.lower():
                                            return "1.x.x"
                                        elif "2." in display_name or "v2" in display_name.lower():
                                            return "2.x.x"
                                        
                                        # Default assumption
                                        return "2.x.x"
                                except OSError:
                                    continue
                except FileNotFoundError:
                    continue
            
            # Default if no TEHTRIS EDR found
            return "2.x.x"
            
        except Exception as e:
            self.logger.warning(f"Could not detect EDR version: {e}")
            return "2.x.x"

    def _add_silent_flags_v1(self, uninstall_string: str) -> str:
        """Add silent flags to V1 uninstall command to avoid confirmation dialogs."""
        self.logger.info(f"[V1] Original uninstall command: {uninstall_string}")
        
        # Check if it's an MSI uninstall command (msiexec)
        if "msiexec" in uninstall_string.lower():
            # For MSI, add silent flags
            silent_flags = " /quiet /norestart"
            if "/quiet" not in uninstall_string.lower():
                silent_command = uninstall_string + silent_flags
                self.logger.info(f"[V1] Added MSI silent flags: {silent_flags}")
            else:
                silent_command = uninstall_string
                self.logger.info(f"[V1] MSI silent flags already present")
        else:
            # For EXE uninstaller, try common silent flags
            silent_flags = [" /S", " /SILENT", " /VERYSILENT", " /q", " /quiet"]
            silent_command = uninstall_string
            
            # Add flags that aren't already present
            for flag in silent_flags:
                if flag.lower().strip() not in uninstall_string.lower():
                    silent_command += flag
            
            # Add additional flags to suppress dialogs
            if "/SUPPRESSMSGBOXES" not in silent_command:
                silent_command += " /SUPPRESSMSGBOXES"
            if "/NORESTART" not in silent_command:
                silent_command += " /NORESTART"
            
            self.logger.info(f"[V1] Added EXE silent flags")
        
        self.logger.info(f"[V1] Final silent command: {silent_command}")
        return silent_command

    def validate_prerequisites(self) -> bool:
        """Validate prerequisites."""
        self.logger.info(f"[{self.edr_version}] Validating prerequisites...")

        # For V1, fully automated - no credentials needed
        if self.edr_version.startswith('1.'):
            self.logger.info("[V1] Fully automated uninstall - credentials not required")
            if self.password or self.key_file:
                self.logger.info("[V1] Credentials provided but not needed for automated V1 uninstall")
            if self.key_file and not self.key_file.exists():
                self.logger.error(f"[V1] Key file not found: {self.key_file}")
                return False
        else:
            # For V2, password or key file is required
            if not self.password and not self.key_file:
                self.logger.error("[V2] Either password or key file must be provided for V2")
                return False

            if self.key_file and not self.key_file.exists():
                self.logger.error(f"[V2] Key file not found: {self.key_file}")
                return False

        self.logger.info(f"[{self.edr_version}] Prerequisites validated successfully")
        return True



    def click_with_win32gui(self, button_text: str) -> bool:
        """Click button using win32gui."""
        try:
            import win32gui
            import win32con

            self.logger.info(f"[{self.edr_version}] Looking for button: {button_text}")

            # Find TEHTRIS windows - check for both setup and simple dialogs
            tehtris_windows = []
            def find_windows(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        # V1 might have simple "TEHTRIS EDR" dialog, V2 has "TEHTRIS EDR Setup"
                        if ("TEHTRIS EDR Setup" in window_text or 
                            ("TEHTRIS EDR" in window_text and self.edr_version.startswith('1.'))):
                            windows.append(hwnd)
                except:
                    pass
                return True

            win32gui.EnumWindows(find_windows, tehtris_windows)
            self.logger.info(f"[{self.edr_version}] Found {len(tehtris_windows)} TEHTRIS windows")

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
                        self.logger.info(f"[{self.edr_version}] Clicked button: {button_info['text']}")
                        return True

                except Exception as e:
                    self.logger.debug(f"Error in window {tehtris_hwnd}: {e}")
                    continue

            self.logger.error(f"[{self.edr_version}] Button '{button_text}' not found")
            return False

        except Exception as e:
            self.logger.error(f"[{self.edr_version}] win32gui click failed: {e}")
            return False

    def click_radio_button(self, radio_text: str) -> bool:
        """Click radio button using win32gui."""
        try:
            import win32gui
            import win32con

            self.logger.info(f"Looking for radio button: {radio_text}")

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

            # Search for radio button
            for tehtris_hwnd in tehtris_windows:
                try:
                    def find_radio(hwnd, radio_info):
                        try:
                            if win32gui.IsWindowVisible(hwnd):
                                window_text = win32gui.GetWindowText(hwnd)
                                class_name = win32gui.GetClassName(hwnd)

                                if class_name == 'Button':
                                    # Check if it's a radio button (style BS_RADIOBUTTON)
                                    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                                    if style & 0x04:  # BS_RADIOBUTTON
                                        if radio_text.lower() in window_text.lower():
                                            radio_info['hwnd'] = hwnd
                                            radio_info['text'] = window_text
                                            return False
                        except:
                            pass
                        return True

                    radio_info = {}
                    win32gui.EnumChildWindows(tehtris_hwnd, find_radio, radio_info)

                    if radio_info.get('hwnd'):
                        win32gui.SendMessage(radio_info['hwnd'], win32con.WM_LBUTTONDOWN, 0, 0)
                        win32gui.SendMessage(radio_info['hwnd'], win32con.WM_LBUTTONUP, 0, 0)
                        self.logger.info(f"Clicked radio button: {radio_info['text']}")
                        return True

                except Exception as e:
                    self.logger.debug(f"Error in window {tehtris_hwnd}: {e}")
                    continue

            self.logger.error(f"Radio button '{radio_text}' not found")
            return False

        except Exception as e:
            self.logger.error(f"Radio button click failed: {e}")
            return False

    def fill_password_field(self, password: str) -> bool:
        """Fill password field using win32gui."""
        try:
            import win32gui
            import win32con

            self.logger.info("Looking for password field")

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

            # Find password field (usually an Edit control with ES_PASSWORD style)
            for tehtris_hwnd in tehtris_windows:
                try:
                    def find_password_field(hwnd, field_info):
                        try:
                            if win32gui.IsWindowVisible(hwnd):
                                class_name = win32gui.GetClassName(hwnd)
                                if class_name in ['Edit', 'RichEdit20W']:
                                    # Check if it's a password field
                                    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                                    if style & 0x20:  # ES_PASSWORD
                                        field_info['hwnd'] = hwnd
                                        return False
                                    # If no password style, check if it's the only edit field
                                    elif not field_info.get('hwnd'):
                                        field_info['hwnd'] = hwnd
                        except:
                            pass
                        return True

                    field_info = {}
                    win32gui.EnumChildWindows(tehtris_hwnd, find_password_field, field_info)

                    if field_info.get('hwnd'):
                        edit_hwnd = field_info['hwnd']
                        # Click on field to set focus
                        rect = win32gui.GetWindowRect(edit_hwnd)
                        center_x = (rect[0] + rect[2]) // 2
                        center_y = (rect[1] + rect[3]) // 2

                        if PYAUTOGUI_AVAILABLE:
                            pyautogui.click(center_x, center_y)
                            time.sleep(0.2)

                        # Clear and set password
                        win32gui.SendMessage(edit_hwnd, win32con.WM_SETTEXT, 0, "")
                        time.sleep(0.1)
                        win32gui.SendMessage(edit_hwnd, win32con.WM_SETTEXT, 0, password)
                        time.sleep(0.2)

                        # Send Tab to trigger validation and set focus
                        self.logger.info("Tabbing to set focus away from the password field...")
                        win32gui.SendMessage(edit_hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
                        win32gui.SendMessage(edit_hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
                        time.sleep(0.2)

                        self.logger.info("Password field filled")
                        return True

                except Exception as e:
                    self.logger.debug(f"Error in window {tehtris_hwnd}: {e}")
                    continue

            self.logger.error("Password field not found")
            return False

        except Exception as e:
            self.logger.error(f"Fill password field failed: {e}")
            return False



    def fill_key_file_path(self, file_path: Path) -> bool:
        """Fill the key file path field."""
        self.logger.info(f"Looking for key file path field for: {file_path}")
        try:
            import win32gui
            import win32con

            tehtris_windows = []
            def find_windows(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd) and "TEHTRIS EDR Setup" in win32gui.GetWindowText(hwnd):
                    windows.append(hwnd)
                return True
            win32gui.EnumWindows(find_windows, tehtris_windows)

            for tehtris_hwnd in tehtris_windows:
                edit_controls = []
                def find_edits(hwnd, controls):
                    if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) in ['Edit', 'RichEdit20W']:
                        # Exclude password fields
                        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                        if not (style & 0x20): # ES_PASSWORD
                            controls.append(hwnd)
                    return True
                win32gui.EnumChildWindows(tehtris_hwnd, find_edits, edit_controls)

                if edit_controls:
                    edit_hwnd = edit_controls[0] # Assume it's the first non-password edit field
                    abs_path = str(file_path.resolve()).strip()
                    win32gui.SendMessage(edit_hwnd, win32con.WM_SETTEXT, 0, abs_path)
                    time.sleep(0.5)
                    self.logger.info(f"Filled key file path with '{abs_path}'")

                    # Click the field to ensure focus and validation
                    if PYAUTOGUI_AVAILABLE:
                        rect = win32gui.GetWindowRect(edit_hwnd)
                        center_x = (rect[0] + rect[2]) // 2
                        center_y = (rect[1] + rect[3]) // 2
                        pyautogui.click(center_x, center_y)
                        self.logger.info("Clicked key file path field to set focus.")
                        time.sleep(0.2)

                    return True

            self.logger.error("Key file path field not found")
            return False

        except Exception as e:
            self.logger.error(f"Fill key file path failed: {e}")
            return False

    def find_and_launch_uninstaller(self) -> bool:
        """Find and launch the TEHTRIS EDR uninstaller using standard Windows method."""
        self.logger.info(f"[{self.edr_version}] Step 1: Finding and launching uninstaller...")
        try:
            import winreg

            uninstall_key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            uninstall_info = None

            # Try 64-bit and 32-bit registry views
            for access_right in [winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY]:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, uninstall_key_path, 0, winreg.KEY_READ | access_right) as key:
                        for i in range(winreg.QueryInfoKey(key)[0]):
                            subkey_name = winreg.EnumKey(key, i)
                            with winreg.OpenKey(key, subkey_name) as subkey:
                                try:
                                    display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                                    if "tehtris edr" in display_name.lower():
                                        try:
                                            uninstall_string = winreg.QueryValueEx(subkey, "UninstallString")[0]
                                            
                                            # Version-based uninstall logic (ignore MSI/EXE detection)
                                            if self.edr_version.startswith('1.'):
                                                # V1: Add silent flags to avoid confirmation dialogs
                                                silent_command = self._add_silent_flags_v1(uninstall_string)
                                                uninstall_info = {
                                                    'version': 'v1',
                                                    'command': silent_command,
                                                    'original_command': uninstall_string,
                                                    'display_name': display_name
                                                }
                                            else:
                                                # V2: Always use MSI uninstall method
                                                import re
                                                product_code_match = re.search(r'{[A-F0-9\-]+}', uninstall_string)
                                                if product_code_match:
                                                    product_code = product_code_match.group(0)
                                                    uninstall_info = {
                                                        'version': 'v2',
                                                        'product_code': product_code,
                                                        'display_name': display_name,
                                                        'command': f'msiexec /x {product_code}'
                                                    }
                                                else:
                                                    # Fallback: use direct uninstall string for V2
                                                    uninstall_info = {
                                                        'version': 'v2',
                                                        'command': uninstall_string,
                                                        'display_name': display_name
                                                    }
                                            break
                                        except OSError:
                                            continue
                                except OSError:
                                    continue
                    if uninstall_info:
                                        break
                except FileNotFoundError:
                    continue

            if not uninstall_info:
                self.logger.error(f"[{self.edr_version}] TEHTRIS EDR uninstaller not found in registry.")
                return False

            self.logger.info(f"[{self.edr_version}] Found uninstaller: {uninstall_info['display_name']}")
            self.logger.info(f"[{self.edr_version}] Uninstaller version: {uninstall_info['version']}")

            # Launch the appropriate uninstall command based on version
            command = uninstall_info['command']
            if uninstall_info['version'] == 'v1':
                self.logger.info(f"[V1] Launching silent uninstall: {command}")
                # For V1 silent uninstall, capture output and wait for completion
                try:
                    result = subprocess.run(command, shell=True, capture_output=True, text=True, timeout=300)
                    self.uninstall_result = {
                        'returncode': result.returncode,
                        'stdout': result.stdout,
                        'stderr': result.stderr,
                        'command': command
                    }
                    self.logger.info(f"[V1] Uninstall command completed with exit code: {result.returncode}")
                    if result.stdout:
                        self.logger.info(f"[V1] Output: {result.stdout}")
                    if result.stderr:
                        self.logger.warning(f"[V1] Error output: {result.stderr}")
                    return True
                except subprocess.TimeoutExpired:
                    self.logger.error(f"[V1] Uninstall command timed out after 5 minutes")
                    return False
                except Exception as e:
                    self.logger.error(f"[V1] Failed to execute uninstall command: {e}")
                    return False
            else:
                self.logger.info(f"[V2] Launching MSI uninstall: {command}")
                subprocess.Popen(command, shell=True)

            time.sleep(5)

            # Wait for the TEHTRIS EDR Setup window to appear
            self.logger.info(f"[{self.edr_version}] Waiting for TEHTRIS EDR Setup window to appear...")
            setup_window_timeout = 30  # 30 seconds
            start_time = time.time()

            while time.time() - start_time < setup_window_timeout:
                if self._check_tehtris_window_exists():
                    self.logger.info(f"[{self.edr_version}] TEHTRIS EDR Setup window detected - Uninstaller launched successfully")
                    return True
                time.sleep(1)

            self.logger.error(f"[{self.edr_version}] TEHTRIS EDR Setup window did not appear within timeout")
            return False

        except Exception as e:
            self.logger.error(f"[{self.edr_version}] Failed to launch uninstaller: {e}")
            return False

    def _check_tehtris_window_exists(self) -> bool:
        """Check if TEHTRIS EDR Setup window exists."""
        try:
            import win32gui

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
            return len(tehtris_windows) > 0
        except Exception as e:
            self.logger.debug(f"Error checking for TEHTRIS window: {e}")
            return False



    def check_for_error_dialog(self) -> (int, str):
        """Check for the 'Error during uninstallation' dialog and extract the message."""
        try:
            import win32gui

            dialog_hwnd = win32gui.FindWindow(None, "Error during uninstallation")
            if dialog_hwnd:
                self.logger.warning("Error dialog detected.")

                # Enumerate all child windows to find the static text controls
                child_windows = []
                def find_children(hwnd, lparam):
                    child_windows.append(hwnd)
                    return True
                win32gui.EnumChildWindows(dialog_hwnd, find_children, None)

                error_messages = []
                for hwnd in child_windows:
                    if win32gui.GetClassName(hwnd) == "Static":
                        text = win32gui.GetWindowText(hwnd)
                        if text:
                            error_messages.append(text)

                full_error_message = " ".join(error_messages).strip()
                if full_error_message:
                    return dialog_hwnd, full_error_message
                else:
                    return dialog_hwnd, "Could not retrieve error message from dialog."

            return 0, None

        except Exception as e:
            self.logger.debug(f"Error checking for error dialog: {e}")
            return 0, None


    def handle_uninstallation_error(self, dialog_hwnd: int, error_message: str):
        """Handle the uninstallation error state."""
        self.logger.error(f"Uninstallation failed with error: {error_message}")

        # Click OK on the error dialog
        try:
            import win32gui
            import win32con
            self.logger.info("Clicking 'OK' on the error dialog...")
            ok_button_hwnd = win32gui.FindWindowEx(dialog_hwnd, 0, "Button", "OK")
            if ok_button_hwnd:
                win32gui.SendMessage(ok_button_hwnd, win32con.BM_CLICK, 0, 0)
            else:
                self.click_with_win32gui("OK") # Fallback
            time.sleep(1)
        except Exception as e:
            self.logger.warning(f"Could not click 'OK' on the error dialog: {e}")

        # Cancel the main uninstallation
        self.logger.info("Attempting to cancel the uninstallation...")
        if not self.click_with_win32gui("Cancel"):
            self.logger.warning("Could not click 'Cancel' button.")

        time.sleep(2)

        # Look for a "Finish" button to close the wizard
        self.logger.info("Looking for 'Finish' button to close the wizard...")
        if not self.click_with_win32gui("Finish"):
             self.logger.warning("Could not find 'Finish' button after cancelling.")


    def center_window(self) -> bool:
        """Center the TEHTRIS EDR Setup window on the screen."""
        self.logger.info("Centering the uninstaller window...")
        try:
            import win32gui
            import win32api

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

            if tehtris_windows:
                hwnd = tehtris_windows[0]

                # Get screen dimensions
                screen_width = win32api.GetSystemMetrics(0)
                screen_height = win32api.GetSystemMetrics(1)

                # Get window dimensions
                rect = win32gui.GetWindowRect(hwnd)
                window_width = rect[2] - rect[0]
                window_height = rect[3] - rect[1]

                # Calculate center position
                center_x = (screen_width - window_width) // 2
                center_y = (screen_height - window_height) // 2

                win32gui.MoveWindow(hwnd, center_x, center_y, window_width, window_height, True)
                self.logger.info("Successfully centered the window.")
                return True
            else:
                self.logger.warning("Could not find TEHTRIS EDR Setup window to center it.")
                return True # Not a fatal error

        except Exception as e:
            self.logger.error(f"Failed to center window: {e}")
            return False

    def handle_welcome_screen(self) -> bool:
        """Handle welcome screen or initial dialog."""
        if self.edr_version.startswith('1.'):
            self.logger.info(f"[V1] Step 2: Handling initial TEHTRIS EDR dialog...")
            # V1 may have a simple dialog with title "TEHTRIS EDR" -> click OK
            time.sleep(1)
            if self.click_with_win32gui("OK"):
                self.logger.info(f"[V1] Successfully clicked OK on initial dialog")
                time.sleep(1)
                # After OK, proceed with Next if there's a welcome screen
                return self.click_with_win32gui("Next")
            else:
                # Maybe no initial dialog, try Next directly
                return self.click_with_win32gui("Next")
        else:
            self.logger.info(f"[V2] Step 2: Handling welcome screen...")
            time.sleep(1)
            return self.click_with_win32gui("Next")

    def handle_verification_screen(self) -> bool:
        """Handle verification screen."""
        self.logger.info(f"[{self.edr_version}] Step 3: Handling verification...")
        time.sleep(1)

        if self.password:
            self.logger.info(f"[{self.edr_version}] Using password verification.")
            if not self.click_radio_button("Enter password"):
                 self.logger.warning(f"[{self.edr_version}] Could not select 'Enter password' radio. Assuming it's default.")
            time.sleep(0.5)
            if not self.fill_password_field(self.password):
                return False

        elif self.key_file:
            self.logger.info(f"[{self.edr_version}] Using key file verification.")
            if not self.click_radio_button("Use key file"):
                return False
            time.sleep(0.5)
            if not self.fill_key_file_path(self.key_file):
                return False

        time.sleep(0.5)
        return self.click_with_win32gui("Next")

    def handle_remove_screen(self) -> bool:
        """Handle remove screen with fallback."""
        self.logger.info(f"[{self.edr_version}] Step 4: Confirming removal...")
        time.sleep(1)

        # Try clicking the Remove button directly
        if self.click_with_win32gui("Remove"):
            return True

        # Fallback to using the keyboard shortcut (Alt+R)
        self.logger.warning("win32gui click failed for 'Remove' button. Trying keyboard shortcut Alt+R.")
        if PYAUTOGUI_AVAILABLE:
            try:
                pyautogui.hotkey('alt', 'r')
                time.sleep(1)
                self.logger.info("Successfully sent Alt+R shortcut.")
                return True
            except Exception as e:
                self.logger.error(f"pyautogui hotkey failed: {e}")
                return False
        else:
            self.logger.error("pyautogui is not available, cannot use fallback hotkey.")
            return False

    def wait_for_completion(self) -> bool:
        """Wait for uninstallation to complete, checking for errors."""
        self.logger.info(f"[{self.edr_version}] Step 5: Waiting for uninstallation completion...")

        completion_timeout = 300  # 5 minutes
        start_time = time.time()

        while time.time() - start_time < completion_timeout:
            elapsed = int(time.time() - start_time)
            self.logger.info(f"Checking for completion... ({elapsed}s elapsed)")

            # Check for the specific error dialog
            dialog_hwnd, error_message = self.check_for_error_dialog()
            if dialog_hwnd:
                self.handle_uninstallation_error(dialog_hwnd, error_message)
                return False # Indicate failure

            # Check for successful completion
            if self.click_with_win32gui("Finish"):
                self.logger.info("Clicked Finish button")
                return True

            time.sleep(2)

        self.logger.error("Uninstallation did not complete within timeout.")
        return False

    def run_uninstallation(self) -> bool:
        """Run complete uninstallation process."""
        self.logger.info(f"[{self.edr_version}] Starting TEHTRIS EDR uninstallation automation")

        try:
            if not self.validate_prerequisites():
                return False
            time.sleep(0.5)

            if not self.find_and_launch_uninstaller():
                return False

            # For V1, the command already completed - no UI handling needed
            if self.edr_version.startswith('1.'):
                return self._verify_v1_uninstall_result()

            # For V2 only, use GUI automation
            time.sleep(0.5)

            # Center the uninstaller window
            if not self.center_window():
                self.logger.warning("Could not center the window.")
            time.sleep(0.5)

            if not self.handle_welcome_screen():
                return False
            time.sleep(0.5)

            if not self.handle_verification_screen():
                return False
            time.sleep(0.5)

            if not self.handle_remove_screen():
                return False
            time.sleep(0.5)

            if not self.wait_for_completion():
                return False

            self.logger.info(f"[{self.edr_version}] TEHTRIS EDR uninstallation completed successfully!")
            return True

        except Exception as e:
            self.logger.error(f"[{self.edr_version}] Uninstallation failed: {e}")
            return False

    def _verify_v1_uninstall_result(self) -> bool:
        """Verify V1 uninstall result based on command output and simple checks."""
        if not hasattr(self, 'uninstall_result'):
            self.logger.error("[V1] No uninstall result available")
            return False
        
        result = self.uninstall_result
        self.logger.info(f"[V1] Verifying uninstall result...")
        self.logger.info(f"[V1] Command: {result['command']}")
        self.logger.info(f"[V1] Exit code: {result['returncode']}")
        
        # Check exit code first
        if result['returncode'] == 0:
            self.logger.info("[V1] Command exit code indicates success (0)")
            success_by_exitcode = True
        else:
            self.logger.warning(f"[V1] Command exit code indicates potential issue ({result['returncode']})")
            success_by_exitcode = False
        
        # Additional verification: check if dasc.exe process is still running
        time.sleep(3)  # Wait a moment for cleanup
        processes_stopped = self._check_processes_stopped()
        
        self.logger.info(f"[V1] Final verification:")
        self.logger.info(f"[V1] - Exit code success: {success_by_exitcode}")
        self.logger.info(f"[V1] - Processes stopped: {processes_stopped}")
        
        # Consider successful if either:
        # 1. Exit code is 0 (command succeeded), OR
        # 2. Processes stopped (uninstall actually worked even if exit code was non-zero)
        if success_by_exitcode or processes_stopped:
            self.logger.info("[V1] Uninstall verification successful!")
            return True
        else:
            self.logger.error("[V1] Uninstall verification failed - both exit code and process check failed")
            return False
    
    def _check_processes_stopped(self) -> bool:
        """Simple check if TEHTRIS processes have stopped."""
        try:
            import psutil
            for proc in psutil.process_iter(['name']):
                if proc.info['name'].lower() == 'dasc.exe':
                    self.logger.info("[V1] TEHTRIS process still running")
                    return False
            self.logger.info("[V1] No TEHTRIS processes found - stopped successfully")
            return True
        except Exception as e:
            self.logger.debug(f"[V1] Error checking processes: {e}")
            return False

def main():
    """Main entry point."""
    import argparse
    parser = argparse.ArgumentParser(
        description="TEHTRIS EDR Uninstaller",
        epilog="""
Note: For V1 installations, password/keyfile is optional.
For V2 installations, either password or keyfile is required.

Examples:
  # V1 uninstall (no credentials needed)
  python tehtris_edr_uninstaller.py
  
  # V2 uninstall with password
  python tehtris_edr_uninstaller.py -p "password123"
  
  # V2 uninstall with keyfile
  python tehtris_edr_uninstaller.py -k "path/to/keyfile.key"
        """,
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    group = parser.add_mutually_exclusive_group(required=False)  # Made optional
    group.add_argument("-p", "--password", help="Uninstallation password (required for V2, optional for V1)")
    group.add_argument("-k", "--keyfile", help="Path to the uninstallation key file (required for V2, optional for V1)")

    args = parser.parse_args()

    # Create uninstaller instance
    uninstaller = TehtrisEDRUninstaller(password=args.password, key_file=args.keyfile)
    
    # Show detected version and requirements
    print(f"Detected EDR version: {uninstaller.edr_version}")
    if uninstaller.edr_version.startswith('1.'):
        print("[V1] Fully automated uninstall - no UI interaction needed")
        if args.password or args.keyfile:
            print("[V1] Note: Credentials provided but not needed for V1")
    else:
        print("[V2] Password or keyfile is required for V2")
        if not args.password and not args.keyfile:
            print("ERROR: V2 requires either password (-p) or keyfile (-k)")
            sys.exit(1)
    
    success = uninstaller.run_uninstallation()
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()
