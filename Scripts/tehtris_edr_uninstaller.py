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

    def scan_available_buttons(self) -> list:
        """Scan and return all available buttons in TEHTRIS windows."""
        available_buttons = []
        try:
            import win32gui
            
            # Find TEHTRIS windows
            tehtris_windows = []
            def find_windows(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        if ("TEHTRIS EDR Setup" in window_text or 
                            ("TEHTRIS EDR" in window_text and self.edr_version.startswith('1.'))):
                            windows.append(hwnd)
                except:
                    pass
                return True
            
            win32gui.EnumWindows(find_windows, tehtris_windows)
            
            for tehtris_hwnd in tehtris_windows:
                try:
                    def find_all_buttons(hwnd, button_list):
                        try:
                            if win32gui.IsWindowVisible(hwnd):
                                window_text = win32gui.GetWindowText(hwnd)
                                class_name = win32gui.GetClassName(hwnd)
                                
                                if window_text and class_name == 'Button':
                                    clean_text = window_text.replace('&', '').lower().strip()
                                    if clean_text and len(clean_text) > 0:
                                        button_list.append(clean_text)
                        except:
                            pass
                        return True
                    
                    win32gui.EnumChildWindows(tehtris_hwnd, find_all_buttons, available_buttons)
                except:
                    pass
            
            # Remove duplicates and sort
            available_buttons = sorted(list(set(available_buttons)))
            return available_buttons
            
        except Exception as e:
            self.logger.debug(f"Button scanning failed: {e}")
            return []

    def _scan_radio_buttons(self) -> list:
        """Scan and return all available radio buttons."""
        radio_buttons = []
        try:
            import win32gui
            import win32con
            
            # Find TEHTRIS windows
            tehtris_windows = []
            def find_windows(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        if ("TEHTRIS EDR Setup" in window_text or 
                            ("TEHTRIS EDR" in window_text and self.edr_version.startswith('1.'))):
                            windows.append(hwnd)
                except:
                    pass
                return True
            
            win32gui.EnumWindows(find_windows, tehtris_windows)
            
            for tehtris_hwnd in tehtris_windows:
                try:
                    def find_radios(hwnd, radio_list):
                        try:
                            if win32gui.IsWindowVisible(hwnd):
                                window_text = win32gui.GetWindowText(hwnd)
                                class_name = win32gui.GetClassName(hwnd)
                                
                                if class_name == 'Button':
                                    # Check if it's a radio button
                                    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                                    if style & 0x04:  # BS_RADIOBUTTON
                                        if window_text and window_text.strip():
                                            radio_list.append(window_text.strip().lower())
                        except:
                            pass
                        return True
                    
                    win32gui.EnumChildWindows(tehtris_hwnd, find_radios, radio_buttons)
                except:
                    pass
            
            return list(set(radio_buttons))
            
        except Exception as e:
            self.logger.debug(f"Radio button scanning failed: {e}")
            return []

    def _count_text_areas(self) -> int:
        """Count text areas/edit controls in current window."""
        try:
            import win32gui
            
            # Find TEHTRIS windows
            tehtris_windows = []
            def find_windows(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        if ("TEHTRIS EDR Setup" in window_text or 
                            ("TEHTRIS EDR" in window_text and self.edr_version.startswith('1.'))):
                            windows.append(hwnd)
                except:
                    pass
                return True
            
            win32gui.EnumWindows(find_windows, tehtris_windows)
            
            text_area_count = 0
            for tehtris_hwnd in tehtris_windows:
                try:
                    def count_text_areas(hwnd, counter):
                        try:
                            if win32gui.IsWindowVisible(hwnd):
                                class_name = win32gui.GetClassName(hwnd)
                                if class_name in ['Edit', 'RichEdit20W', 'RichEdit20A']:
                                    counter[0] += 1
                        except:
                            pass
                        return True
                    
                    counter = [0]
                    win32gui.EnumChildWindows(tehtris_hwnd, count_text_areas, counter)
                    text_area_count += counter[0]
                except:
                    pass
            
            return text_area_count
            
        except Exception as e:
            self.logger.debug(f"Text area counting failed: {e}")
            return 0

    def detect_current_step(self) -> str:
        """Detect current uninstaller step based on available buttons and UI elements."""
        buttons = self.scan_available_buttons()
        radio_buttons = self._scan_radio_buttons()
        text_areas = self._count_text_areas()
        
        self.logger.info(f"[{self.edr_version}] Available buttons: {buttons}")
        self.logger.info(f"[{self.edr_version}] Available radio buttons: {radio_buttons}")
        self.logger.info(f"[{self.edr_version}] Text areas count: {text_areas}")
        
        # V2 Uninstaller step detection patterns:
        # Step 1: ['back', 'next', 'cancel']
        # Step 2: ['back', 'next', 'cancel'] + 2 radio buttons + 1 textarea
        
        # Check for completion/final steps first
        if any(btn in buttons for btn in ['finish', 'close', 'done', 'exit']):
            return 'complete'
        elif any(btn in buttons for btn in ['remove', 'uninstall', 'delete']):
            return 'remove'
        elif any(btn in buttons for btn in ['yes', 'ok', 'confirm']):
            return 'confirmation'
        
        # V2 specific step detection
        if self.edr_version.startswith('2.') or not self.edr_version.startswith('1.'):
            # Check for Step 2: back/next/cancel + 2 radio buttons + 1 textarea
            if (any(btn in buttons for btn in ['back', '< back']) and 
                any(btn in buttons for btn in ['next', 'next >', 'continue']) and
                any(btn in buttons for btn in ['cancel']) and
                len(radio_buttons) >= 2 and text_areas >= 1):
                self.logger.info(f"[{self.edr_version}] Detected Step 2 (verification) - has radio buttons and textarea")
                return 'verification'  # Step 2
            
            # Check for Step 1: just back/next/cancel (no radio buttons or text areas)
            elif (any(btn in buttons for btn in ['back', '< back']) and 
                  any(btn in buttons for btn in ['next', 'next >', 'continue']) and
                  any(btn in buttons for btn in ['cancel']) and
                  len(radio_buttons) == 0 and text_areas == 0):
                self.logger.info(f"[{self.edr_version}] Detected Step 1 (welcome) - basic navigation only")
                return 'welcome'  # Step 1
        
        # V1 fallback patterns
        elif any(btn in buttons for btn in ['next', 'next >', 'continue']):
            return 'welcome'
        
        self.logger.warning(f"[{self.edr_version}] Unknown step detected")
        return 'unknown'

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
        """Find and launch the TEHTRIS EDR uninstaller."""
        self.logger.info(f"[{self.edr_version}] Step 1: Finding and launching uninstaller...")
        try:
            import winreg

            uninstall_key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            uninstall_command = None
            display_name = None

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
                                        uninstall_command = winreg.QueryValueEx(subkey, "UninstallString")[0]
                                        self.logger.info(f"[{self.edr_version}] Found uninstaller: {display_name}")
                                        break
                                except OSError:
                                    continue
                    if uninstall_command:
                        break
                except FileNotFoundError:
                    continue

            if not uninstall_command:
                self.logger.error(f"[{self.edr_version}] TEHTRIS EDR uninstaller not found in registry.")
                return False

            # For V1, add silent flags and run directly
            if self.edr_version.startswith('1.'):
                silent_command = self._add_silent_flags_v1(uninstall_command)
                self.logger.info(f"[V1] Launching silent uninstall: {silent_command}")
                
                try:
                    result = subprocess.run(silent_command, shell=True, capture_output=True, text=True, timeout=300)
                    self.uninstall_result = {
                        'returncode': result.returncode,
                        'stdout': result.stdout,
                        'stderr': result.stderr,
                        'command': silent_command
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
            
            # For V2, launch normally and proceed with GUI automation
            else:
                self.logger.info(f"[V2] Launching uninstaller: {uninstall_command}")
                subprocess.Popen(uninstall_command, shell=True)
                time.sleep(5)

                # Wait for the TEHTRIS EDR Setup window to appear
                self.logger.info(f"[V2] Waiting for TEHTRIS EDR Setup window to appear...")
                setup_window_timeout = 30  # 30 seconds
                start_time = time.time()

                while time.time() - start_time < setup_window_timeout:
                    if self._check_tehtris_window_exists():
                        self.logger.info(f"[V2] TEHTRIS EDR Setup window detected - Uninstaller launched successfully")
                        return True
                    time.sleep(1)

                self.logger.error(f"[V2] TEHTRIS EDR Setup window did not appear within timeout")
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
        """Check for various error dialogs and extract the message."""
        try:
            import win32gui

            # List of potential error dialog titles
            error_dialog_titles = [
                "Error during uninstallation",
                "Error",
                "Uninstall Error", 
                "Installation Error",
                "TEHTRIS EDR Error",
                "Windows Installer"
            ]
            
            # Also check for dialogs containing error keywords
            error_keywords = ["error", "failed", "cannot", "unable"]
            
            self.logger.debug(f"[{self.edr_version}] Checking for error dialogs...")
            
            # First, try specific error dialog titles
            for title in error_dialog_titles:
                dialog_hwnd = win32gui.FindWindow(None, title)
                if dialog_hwnd:
                    self.logger.warning(f"[{self.edr_version}] Error dialog detected with title: '{title}'")
                    error_message = self._extract_error_message(dialog_hwnd, title)
                    return dialog_hwnd, error_message
            
            # Then, enumerate all visible dialogs and check for error keywords
            error_dialogs = []
            def find_error_dialogs(hwnd, dialogs):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        class_name = win32gui.GetClassName(hwnd)
                        
                        # Look for dialog boxes with error-related text
                        if window_text and class_name in ['#32770', 'Dialog']:
                            for keyword in error_keywords:
                                if keyword.lower() in window_text.lower():
                                    dialogs.append((hwnd, window_text))
                                    break
                except:
                    pass
                return True
                
            win32gui.EnumWindows(find_error_dialogs, error_dialogs)
            
            if error_dialogs:
                dialog_hwnd, window_title = error_dialogs[0]  # Take the first error dialog found
                self.logger.warning(f"[{self.edr_version}] Error dialog detected by keyword: '{window_title}'")
                error_message = self._extract_error_message(dialog_hwnd, window_title)
                return dialog_hwnd, error_message

            return 0, None

        except Exception as e:
            self.logger.debug(f"[{self.edr_version}] Error checking for error dialog: {e}")
            return 0, None

    def _extract_error_message(self, dialog_hwnd: int, dialog_title: str) -> str:
        """Extract error message from dialog."""
        try:
            import win32gui
            
            # Enumerate all child windows to find the static text controls
            child_windows = []
            def find_children(hwnd, lparam):
                child_windows.append(hwnd)
                return True
            win32gui.EnumChildWindows(dialog_hwnd, find_children, None)

            error_messages = []
            for hwnd in child_windows:
                try:
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name in ["Static", "Edit"]:  # Include Edit controls too
                        text = win32gui.GetWindowText(hwnd)
                        if text and len(text.strip()) > 2:  # Ignore very short text
                            error_messages.append(text.strip())
                except:
                    continue

            if error_messages:
                full_error_message = " | ".join(error_messages)
                return f"Dialog: '{dialog_title}' - Message: {full_error_message}"
            else:
                return f"Error dialog detected: '{dialog_title}' (no detailed message found)"
                
        except Exception as e:
            self.logger.debug(f"[{self.edr_version}] Error extracting message: {e}")
            return f"Error dialog detected: '{dialog_title}' (message extraction failed)"


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
        """Handle welcome screen (Step 1) with smart step detection."""
        self.logger.info(f"[{self.edr_version}] Step 1: Handling welcome screen...")
        
        timeout = 30  # 30 seconds
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            # Detect current step
            current_step = self.detect_current_step()
            
            if current_step == 'verification':
                self.logger.info(f"[{self.edr_version}] Already moved to Step 2 (verification)")
                return True
            elif current_step == 'remove':
                self.logger.info(f"[{self.edr_version}] Already moved to remove step")
                return True
            elif current_step == 'complete':
                self.logger.info(f"[{self.edr_version}] Already moved to completion")
                return True
            
            # Try version-specific handling
            if self.edr_version.startswith('1.'):
                # V1 may have initial dialog with OK button first
                if self.click_with_win32gui("OK"):
                    self.logger.info(f"[V1] Successfully clicked OK on initial dialog")
                    time.sleep(0.5)
                    continue  # Re-check step after OK
                elif self.click_with_win32gui("Next"):
                    self.logger.info(f"[V1] Successfully clicked Next on Step 1")
                    time.sleep(0.5)
                    # Verify step advancement
                    new_step = self.detect_current_step()
                    if new_step != 'welcome' and new_step != 'unknown':
                        self.logger.info(f"[V1] Step 1 completed, moved to {new_step}")
                        return True
            else:
                # V2 handling - should be Step 1: back/next/cancel only
                if current_step == 'welcome':
                    if self.click_with_win32gui("Next"):
                        self.logger.info(f"[V2] Successfully clicked Next on Step 1 (welcome)")
                        time.sleep(0.5)
                        # Verify step advancement
                        new_step = self.detect_current_step()
                        if new_step == 'verification':
                            self.logger.info(f"[V2] Step 1 completed, moved to Step 2 (verification)")
                            return True
                        elif new_step != 'welcome' and new_step != 'unknown':
                            self.logger.info(f"[V2] Step 1 completed, moved to {new_step}")
                            return True
            
            self.logger.info(f"[{self.edr_version}] Step 1 not ready yet, retrying in 1 second...")
            time.sleep(1)
        
        self.logger.error(f"[{self.edr_version}] Failed to handle Step 1 within timeout")
        return False

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

        # First, let's see what buttons are available
        self._debug_available_buttons()

        # Try multiple variations of the Remove button text
        remove_variations = ["Remove", "&Remove", "Uninstall", "&Uninstall", "Delete", "&Delete"]
        
        for remove_text in remove_variations:
            self.logger.info(f"[{self.edr_version}] Trying to click: '{remove_text}'")
            if self.click_with_win32gui(remove_text):
                self.logger.info(f"[{self.edr_version}] Successfully clicked '{remove_text}' button")
                time.sleep(2)  # Wait for removal to start
                return True

        # Fallback to using the keyboard shortcut (Alt+R)
        self.logger.warning(f"[{self.edr_version}] All Remove button variations failed. Trying keyboard shortcut Alt+R.")
        if PYAUTOGUI_AVAILABLE:
            try:
                pyautogui.hotkey('alt', 'r')
                time.sleep(2)
                self.logger.info(f"[{self.edr_version}] Successfully sent Alt+R shortcut.")
                return True
            except Exception as e:
                self.logger.error(f"[{self.edr_version}] pyautogui hotkey failed: {e}")
                return False
        else:
            self.logger.error(f"[{self.edr_version}] pyautogui is not available, cannot use fallback hotkey.")
            return False

    def _debug_available_buttons(self):
        """Debug helper to list all available buttons."""
        try:
            import win32gui
            self.logger.info(f"[{self.edr_version}] DEBUG: Scanning for available buttons...")
            
            # Find TEHTRIS windows
            tehtris_windows = []
            def find_windows(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        if ("TEHTRIS EDR Setup" in window_text or 
                            ("TEHTRIS EDR" in window_text and self.edr_version.startswith('1.'))):
                            windows.append(hwnd)
                except:
                    pass
                return True

            win32gui.EnumWindows(find_windows, tehtris_windows)
            
            for tehtris_hwnd in tehtris_windows:
                try:
                    def find_buttons(hwnd, button_list):
                        try:
                            if win32gui.IsWindowVisible(hwnd):
                                window_text = win32gui.GetWindowText(hwnd)
                                class_name = win32gui.GetClassName(hwnd)
                                
                                if window_text and class_name == 'Button':
                                    button_list.append(window_text)
                        except:
                            pass
                        return True

                    button_list = []
                    win32gui.EnumChildWindows(tehtris_hwnd, find_buttons, button_list)
                    
                    if button_list:
                        self.logger.info(f"[{self.edr_version}] DEBUG: Available buttons: {button_list}")
                    else:
                        self.logger.info(f"[{self.edr_version}] DEBUG: No buttons found in window")
                        
                except Exception as e:
                    self.logger.debug(f"[{self.edr_version}] DEBUG: Error scanning buttons: {e}")
                    
        except Exception as e:
            self.logger.debug(f"[{self.edr_version}] DEBUG: Button scan failed: {e}")

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
                self.logger.info(f"[{self.edr_version}] Clicked Finish button - uninstall completed successfully")
                time.sleep(2)  # Give time for the uninstaller to fully close
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

            completion_result = self.wait_for_completion()
            self.logger.info(f"[{self.edr_version}] wait_for_completion returned: {completion_result}")
            
            if not completion_result:
                self.logger.error(f"[{self.edr_version}] wait_for_completion returned False")
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
