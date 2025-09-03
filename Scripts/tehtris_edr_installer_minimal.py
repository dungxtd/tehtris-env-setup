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
import re
from pathlib import Path

try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.1
except ImportError:
    PYAUTOGUI_AVAILABLE = False

class TehtrisEDRInstaller:
    """Minimal TEHTRIS EDR installer automation."""

    def __init__(self, installer_path: str, uninstall_password: str = None, uninstall_key_file: str = None):
        self.installer_path = Path(installer_path)
        self.uninstall_password = uninstall_password
        self.uninstall_key_file = uninstall_key_file
        self.logger = self._setup_logging()

        # Detect EDR version from filename
        self.edr_version = self._detect_edr_version()
        self.logger.info(f"Detected EDR version: {self.edr_version}")

        # Configuration
        self.server_address = "xfiapp12.tehtris.net"
        self.tag = "XFI_QAT"
        self.license_key = "HTDT-45FK-BQH8-HHMQ-585H"

        # Version-specific configuration
        self.requires_license_key = self._requires_license_key()

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

    def _detect_edr_version(self) -> str:
        """Detect EDR version from filename."""
        filename = self.installer_path.name.lower()

        # Extract version using regex pattern
        # Matches patterns like: 1.8.1, 2.0.0, etc.
        version_pattern = r'(\d+)\.(\d+)\.(\d+)'
        match = re.search(version_pattern, filename)

        if match:
            major, minor, patch = match.groups()
            return f"{major}.{minor}.{patch}"

        # Fallback: try to detect major version from filename
        if re.search(r'[_\-]1[\._\-]', filename):
            return "1.x.x"
        elif re.search(r'[_\-]2[\._\-]', filename):
            return "2.x.x"

        # Default to 2.x.x if cannot detect
        return "2.x.x"

    def _requires_license_key(self) -> bool:
        """Determine if this version requires license key during installation."""
        # Version 1.x.x does not require license key
        # Version 2.x.x requires license key
        if self.edr_version.startswith('1.'):
            return False
        return True

    def validate_prerequisites(self) -> bool:
        """Validate prerequisites."""
        self.logger.info("Validating prerequisites...")

        if not self.installer_path.exists():
            self.logger.error(f"Installer file not found: {self.installer_path}")
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

    def click_with_win32gui(self, button_text: str, button_variations: list = None) -> bool:
        """Click button using win32gui with multiple variations support."""
        try:
            import win32gui
            import win32con

            # Default variations if none provided
            if button_variations is None:
                button_variations = [button_text]
            
            self.logger.info(f"Looking for button: {button_text} (variations: {len(button_variations)})")

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
                                    
                                    # Check all variations
                                    for variation in button_variations:
                                        clean_variation = variation.replace('&', '').lower()
                                        if clean_variation in clean_text or clean_text in clean_variation:
                                            button_info['hwnd'] = hwnd
                                            button_info['text'] = window_text
                                            button_info['matched_variation'] = variation
                                            return False
                        except:
                            pass
                        return True

                    button_info = {}
                    win32gui.EnumChildWindows(tehtris_hwnd, find_button, button_info)

                    if button_info.get('hwnd'):
                        win32gui.SendMessage(button_info['hwnd'], win32con.WM_LBUTTONDOWN, 0, 0)
                        win32gui.SendMessage(button_info['hwnd'], win32con.WM_LBUTTONUP, 0, 0)
                        matched_var = button_info.get('matched_variation', button_text)
                        self.logger.info(f"Clicked button: {button_info['text']} (matched: {matched_var})")
                        return True

                except Exception as e:
                    self.logger.debug(f"Error in window {tehtris_hwnd}: {e}")
                    continue

            self.logger.error(f"Button '{button_text}' not found (tried {len(button_variations)} variations)")
            return False

        except Exception as e:
            self.logger.error(f"win32gui click failed: {e}")
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

            # First, let's debug what radio buttons are available
            self._debug_radio_buttons()

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
                                        clean_text = window_text.lower().strip()
                                        clean_radio = radio_text.lower().strip()
                                        # More flexible matching for radio buttons
                                        if (clean_radio in clean_text or 
                                            clean_text in clean_radio or
                                            ("accept" in clean_radio and "accept" in clean_text and "terms" in clean_text) or
                                            ("i accept" in clean_radio and "i accept" in clean_text)):
                                            radio_info['hwnd'] = hwnd
                                            radio_info['text'] = window_text
                                            return False
                        except:
                            pass
                        return True

                    radio_info = {}
                    win32gui.EnumChildWindows(tehtris_hwnd, find_radio, radio_info)

                    if radio_info.get('hwnd'):
                        # Use BM_CLICK message for more reliable radio button clicking
                        win32gui.SendMessage(radio_info['hwnd'], win32con.BM_CLICK, 0, 0)
                        self.logger.info(f"Clicked radio button: {radio_info['text']}")
                        time.sleep(0.3)  # Wait for radio button state to update
                        return True

                except Exception as e:
                    self.logger.debug(f"Error in window {tehtris_hwnd}: {e}")
                    continue

            self.logger.error(f"Radio button '{radio_text}' not found")
            return False

        except Exception as e:
            self.logger.error(f"Radio button click failed: {e}")
            return False
    
    def _debug_radio_buttons(self):
        """Debug helper to list all available radio buttons."""
        try:
            import win32gui
            import win32con
            
            self.logger.info("DEBUG: Scanning for available radio buttons...")
            
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
            
            for tehtris_hwnd in tehtris_windows:
                try:
                    def find_radios(hwnd, radio_list):
                        try:
                            if win32gui.IsWindowVisible(hwnd):
                                window_text = win32gui.GetWindowText(hwnd)
                                class_name = win32gui.GetClassName(hwnd)
                                
                                if class_name == 'Button':
                                    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                                    if style & 0x04:  # BS_RADIOBUTTON
                                        if window_text.strip():
                                            radio_list.append(window_text)
                        except:
                            pass
                        return True

                    radio_list = []
                    win32gui.EnumChildWindows(tehtris_hwnd, find_radios, radio_list)
                    
                    if radio_list:
                        self.logger.info(f"DEBUG: Available radio buttons: {radio_list}")
                    else:
                        self.logger.info("DEBUG: No radio buttons found in window")
                        
                except Exception as e:
                    self.logger.debug(f"DEBUG: Error scanning radio buttons: {e}")
                    
        except Exception as e:
            self.logger.debug(f"DEBUG: Radio button scan failed: {e}")

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
                            time.sleep(0.1)
                        win32gui.SendMessage(edit_hwnd, win32con.WM_SETTEXT, 0, "")
                        time.sleep(0.05)
                        win32gui.SendMessage(edit_hwnd, win32con.WM_SETTEXT, 0, value)
                        time.sleep(0.1)

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
                        if "TEHTRIS EDR Setup" in window_text:
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
                    continue
                    
        except Exception as e:
            self.logger.debug(f"Button scan failed: {e}")
        
        # Remove duplicates and return
        return list(set(available_buttons))
    
    def detect_current_step(self) -> str:
        """Detect current installer step based on available buttons."""
        buttons = self.scan_available_buttons()
        self.logger.info(f"Available buttons: {buttons}")
        
        # Step detection logic based on actual button patterns:
        # Step 1 (Welcome): ['next >', 'cancel', '< back']
        # Step 2 (License): ['i do not accept the terms in the license agreement', 'cancel', 'i accept the terms in the license agreement', '< back', 'next >']  
        # Step 3 (Activation): ['next >', 'cancel', '< back'] + 2 input fields
        
        if any(btn in buttons for btn in ['install', 'installing']):
            return 'installation'
        elif any(btn in buttons for btn in ['finish', 'close', 'done']):
            return 'complete'
        # Step 2 (License): Has "i accept" and "i do not accept" radio buttons
        elif any(btn in buttons for btn in ['i accept the terms in the license agreement', 'i do not accept the terms in the license agreement']):
            return 'license'
        elif any(btn in buttons for btn in ['next', 'next >', 'continue']) and any(btn in buttons for btn in ['< back', 'back']) and any(btn in buttons for btn in ['cancel']):
            # Both Step 1 and Step 3 have ['next >', 'cancel', '< back']
            # Differentiate by checking for input fields (Step 3 has 2 input fields)
            if self._has_edit_fields():
                return 'activation'  # Step 3
            else:
                return 'welcome'     # Step 1
        else:
            return 'unknown'
    
    def _has_edit_fields(self) -> bool:
        """Check if current window has edit fields (indicates activation step)."""
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
                    field_count = len(edit_controls)
                    
                    # EDR v1: 2 input fields (Server, Tag)
                    # EDR v2: 3 input fields (Server, Tag, License Key)
                    if field_count >= 2:
                        self.logger.info(f"Found {field_count} input fields - indicates activation step")
                        return True
                except:
                    continue
            
            return False
        except:
            return False
    
    def check_for_next_button_fast(self) -> bool:
        """Fast check specifically for Next button - optimized for lag detection."""
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
            
            for tehtris_hwnd in tehtris_windows:
                try:
                    def find_next_button(hwnd, found_next):
                        try:
                            if win32gui.IsWindowVisible(hwnd):
                                window_text = win32gui.GetWindowText(hwnd)
                                class_name = win32gui.GetClassName(hwnd)
                                
                                if window_text and class_name == 'Button':
                                    clean_text = window_text.replace('&', '').lower().strip()
                                    if clean_text in ['next', 'next >', 'continue']:
                                        found_next['exists'] = True
                                        return False  # Stop enumeration
                        except:
                            pass
                        return True
                    
                    found_next = {'exists': False}
                    win32gui.EnumChildWindows(tehtris_hwnd, find_next_button, found_next)
                    
                    if found_next['exists']:
                        return True
                        
                except:
                    continue
            
            return False
        except Exception as e:
            self.logger.debug(f"Fast next button check failed: {e}")
            return False
    
    def wait_for_step_transition(self, from_step: str, expected_steps: list, timeout: int = 10) -> str:
        """Wait for step transition and return the new step."""
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            current_step = self.detect_current_step()
            
            if current_step in expected_steps:
                self.logger.info(f"Step transition successful: {from_step} -> {current_step}")
                return current_step
            elif current_step != from_step and current_step != 'unknown':
                self.logger.info(f"Unexpected step transition: {from_step} -> {current_step}")
                return current_step
            
            time.sleep(0.5)
        
        self.logger.warning(f"Step transition timeout: still on {from_step}")
        return self.detect_current_step()

    def launch_installer(self) -> bool:
        """Launch installer (MSI or EXE)."""
        self.logger.info("Step 1: Launching installer...")

        try:
            # Minimize windows
            # if PYAUTOGUI_AVAILABLE:
            #     pyautogui.hotkey('win', 'd')
            #     time.sleep(1.5)

            # Determine installer type and launch accordingly
            file_extension = self.installer_path.suffix.lower()
            
            if file_extension == '.msi':
                # Launch MSI installer with msiexec
                subprocess.Popen(['msiexec', '/i', str(self.installer_path)], shell=True)
                self.logger.info("MSI installer launched with msiexec")
            elif file_extension == '.exe':
                # Launch EXE installer directly
                subprocess.Popen([str(self.installer_path)], shell=True)
                self.logger.info("EXE installer launched directly")
            else:
                self.logger.error(f"Unsupported installer type: {file_extension}")
                return False
                
            time.sleep(3)
            self.logger.info("Installer launched successfully")
            return True

        except Exception as e:
            self.logger.error(f"Failed to launch installer: {e}")
            return False

    def handle_welcome_screen(self) -> bool:
        """Handle welcome screen (Step 1) - buttons: ['next >', 'cancel', '< back']"""
        self.logger.info("Step 1: Handling welcome screen...")

        # Step 1 only has Next button - NO license agreement buttons
        next_variations = ["Next", "Next >", "&Next", "Continue", "next >", "next"]
        
        timeout = 30  # 30 seconds
        start_time = time.time()

        while time.time() - start_time < timeout:
            # Verify we're on welcome screen (Step 1)
            current_step = self.detect_current_step()
            if current_step == 'license':
                self.logger.info("Already moved to license step (Step 2)")
                return True
            elif current_step == 'activation':
                self.logger.info("Already moved to activation step (Step 3)")
                return True
            elif current_step != 'welcome' and current_step != 'unknown':
                self.logger.info(f"Already moved to {current_step} step")
                return True
            
            if self.click_with_win32gui("Next", next_variations):
                self.logger.info("Successfully clicked 'Next >' on Step 1 (welcome screen)")
                
                # Wait and verify step transition
                time.sleep(1)
                new_step = self.detect_current_step()
                
                if new_step == 'license':
                    self.logger.info("Step 1 completed, moved to Step 2 (license)")
                    return True
                elif new_step != 'welcome' and new_step != 'unknown':
                    self.logger.info(f"Step 1 completed, moved to {new_step}")
                    return True
                else:
                    self.logger.info("Step 1: Click didn't advance step, retrying...")
                    continue
            
            self.logger.info("Step 1: Welcome screen not ready yet, retrying in 1 second...")
            time.sleep(1)

        self.logger.error("Step 1: Failed to handle welcome screen within timeout")
        return False

    def handle_license_agreement(self) -> bool:
        """Handle license agreement (Step 2) with radio button selection."""
        self.logger.info("Step 2: Handling license agreement...")
        
        # Wait longer to ensure we're really on license step
        time.sleep(0.8)  # Increased wait time
        
        # Verify we're actually on license step before proceeding
        current_step = self.detect_current_step()
        if current_step != 'license':
            if current_step == 'welcome':
                self.logger.warning("Still on welcome screen, installer may have lagged")
                return False  # Will be handled by main loop to restart from welcome
            elif current_step == 'activation':
                self.logger.info("License step was skipped, already on activation")
                return True
            elif current_step == 'unknown':
                self.logger.warning("Step unclear, waiting for UI to stabilize...")
                time.sleep(1)
                current_step = self.detect_current_step()
                if current_step != 'license':
                    self.logger.info(f"Not on license step (detected: {current_step}), proceeding...")
                    return current_step == 'activation'

        # First, try to find and click "I accept" radio button
        accept_radio_variations = [
            "I accept the terms in the License Agreement",
            "I accept the terms",
            "I accept", 
            "I Accept", 
            "accept", 
            "Accept"
        ]
        
        radio_clicked = False
        for radio_text in accept_radio_variations:
            if self.click_radio_button(radio_text):
                self.logger.info(f"Successfully selected '{radio_text}' radio button")
                radio_clicked = True
                break
        
        if not radio_clicked:
            self.logger.warning("Could not find 'I accept' radio button, trying regular accept button...")
            # Fallback to regular button clicking
            accept_variations = [
                "accept", "Accept", "I accept", "I Accept", "&Accept", 
                "agree", "Agree", "I agree", "I Agree", "&Agree",
                "yes", "Yes", "OK"
            ]
            
            if not self.click_with_win32gui("accept", accept_variations):
                self.logger.warning("No accept button or radio button found")
                # Continue to Next button attempt
        
        # Wait a moment for radio button selection to register
        time.sleep(0.5)
        
        # Now click Next button to proceed
        next_variations = ["Next", "Next >", "&Next", "Continue", "next >", "next"]
        if self.click_with_win32gui("Next", next_variations):
            self.logger.info("Successfully clicked Next after accepting license")
            
            # Wait and verify the step transition
            time.sleep(1)  # Give time for UI to update
            new_step = self.detect_current_step()
            
            if new_step == 'activation' or new_step == 'installation' or new_step != 'license':
                self.logger.info(f"License step completed, moved to: {new_step}")
                return True
            else:
                self.logger.warning(f"Step didn't change as expected, current: {new_step}")
                return False
        else:
            # Could not find Next button, check if we have step 1 reload scenario
            self.logger.warning("Could not find Next button after accepting license, checking for step reload...")
            if self.check_for_next_button_fast():
                current_step = self.detect_current_step()
                if current_step == 'welcome':
                    self.logger.warning("Installer lagged and reloaded to step 1 (welcome screen)")
                    return False  # Will be handled by main loop to restart from welcome
                elif current_step == 'activation':
                    self.logger.info("Skipped license step, moving to activation")
                    return True  # License was skipped, continue to activation
            
            self.logger.error("Could not proceed after license acceptance")
            return False

    def handle_activation_information(self) -> bool:
        """Handle activation information."""
        self.logger.info("Step 4: Handling activation information...")
        time.sleep(0.3)

        # Fill server address field
        if not self.fill_field_with_win32gui("server", self.server_address):
            return False

        # Fill tag field
        if not self.fill_field_with_win32gui("tag", self.tag):
            return False

        # Fill license key field only for version 2.x.x
        if self.requires_license_key:
            self.logger.info("Version 2.x.x detected - filling license key field")
            if not self.fill_field_with_win32gui("license", self.license_key):
                return False
        else:
            self.logger.info("Version 1.x.x detected - skipping license key field")

        time.sleep(0.5)
        next_variations = ["Next", "Next >", "&Next", "Continue", "next >", "next"]
        return self.click_with_win32gui("Next", next_variations)

    def handle_installation(self) -> bool:
        """Handle installation."""
        self.logger.info("Step 5: Handling installation...")
        time.sleep(0.3)

        # Try clicking Install button or use Alt+I
        if not self.click_with_win32gui("Install"):
            if PYAUTOGUI_AVAILABLE:
                pyautogui.hotkey('alt', 'i')
                time.sleep(0.5)

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

            time.sleep(1)

        # Final attempt with Alt+F
        if PYAUTOGUI_AVAILABLE:
            self.logger.info("Trying Alt+F as final attempt")
            pyautogui.hotkey('alt', 'f')
            time.sleep(0.5)
            return True

        self.logger.error("Installation did not complete within timeout")
        return False

    def verify_installation(self, post_install_check: bool = False) -> bool:
        """Verify installation by checking for TEHTRIS processes based on version."""
        if post_install_check:
            self.logger.info("Step 7: Verifying installation...")
        else:
            self.logger.info("Checking for existing TEHTRIS EDR installation...")

        try:
            import psutil

            if self.edr_version.startswith('1.'):
                self.logger.info("[EDR V1] Checking for dasc.exe with TEHTRIS/EDR/Agent description...")
                return self._verify_v1_installation(post_install_check)
            else:
                self.logger.info("[EDR V2] Checking for Agent processes with TEHTRIS description...")
                return self._verify_v2_installation(post_install_check)

        except Exception as e:
            self.logger.warning(f"Verification failed: {e}")
            return post_install_check if post_install_check else False

    def _verify_v1_installation(self, post_install_check: bool) -> bool:
        """Verify v1.x.x installation by checking for dasc.exe running as Windows Service."""
        try:
            import psutil

            tehtris_processes = []

            for proc in psutil.process_iter(['pid', 'name', 'exe', 'username']):
                try:
                    proc_info = proc.info
                    proc_name = proc_info['name'].lower()

                    # Look for dasc.exe process (V1 runs as Windows Service)
                    if proc_name == 'dasc.exe':
                        exe_path = proc_info.get('exe', '')
                        username = proc_info.get('username', '')
                        
                        # V1 dasc.exe typically runs under SYSTEM or as a service
                        is_service = ('system' in username.lower() if username else True) or not username
                        
                        tehtris_processes.append({
                            'pid': proc_info['pid'],
                            'name': proc_info['name'],
                            'exe': exe_path,
                            'username': username or 'N/A',
                            'is_service': is_service
                        })
                        
                        service_info = "(Service)" if is_service else "(User Process)"
                        self.logger.info(f"[FOUND V1] TEHTRIS dasc.exe {service_info}: PID {proc_info['pid']} - {exe_path}")
                        self.logger.info(f"[FOUND V1] Running as: {username or 'SYSTEM/Service'}")
                        
                        # Try to get additional service info
                        try:
                            # Check if it's actually running as a Windows Service
                            import subprocess
                            result = subprocess.run(['tasklist', '/svc', '/fi', f'PID eq {proc_info["pid"]}'], 
                                                  capture_output=True, text=True, shell=True)
                            if 'Services' in result.stdout:
                                self.logger.info(f"[FOUND V1] Confirmed as Windows Service")
                        except Exception:
                            pass

                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue
                except Exception as e:
                    # Log but continue - some processes may not have all info accessible
                    self.logger.debug(f"[V1 DEBUG] Could not get full info for process: {e}")
                    continue

            if tehtris_processes:
                if post_install_check:
                    self.logger.info(f"[SUCCESS V1] TEHTRIS EDR V1 verified - Found {len(tehtris_processes)} dasc.exe process(es)")
                else:
                    self.logger.warning("[DETECTED V1] TEHTRIS EDR V1 installation detected!")
                return True
            else:
                if post_install_check:
                    self.logger.warning("[NOT FOUND V1] No TEHTRIS V1 processes (dasc.exe) found")
                    self.logger.warning("[NOT FOUND V1] Installation may have completed but processes haven't started yet")
                else:
                    self.logger.info("[NOT FOUND V1] No existing TEHTRIS EDR V1 installation found.")
                return False

        except Exception as e:
            self.logger.warning(f"[ERROR V1] V1 verification failed: {e}")
            return False

    def _verify_v2_installation(self, post_install_check: bool) -> bool:
        """Verify v2.x.x installation by checking for Agent processes with TEHTRIS description."""
        try:
            import psutil

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
                            self.logger.info(f"[FOUND V2] TEHTRIS Agent: PID {proc_info['pid']} - {proc_info['name']} - {exe_path}")

                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue

            if tehtris_agents:
                if post_install_check:
                    self.logger.info(f"[SUCCESS V2] TEHTRIS EDR V2 verified - Found {len(tehtris_agents)} Agent process(es)")
                else:
                    self.logger.warning("[DETECTED V2] TEHTRIS EDR V2 installation detected!")
                return True
            else:
                if post_install_check:
                    self.logger.warning("[NOT FOUND V2] No TEHTRIS V2 Agent processes found")
                    self.logger.warning("[NOT FOUND V2] Installation may have completed but processes haven't started yet")
                else:
                    self.logger.info("[NOT FOUND V2] No existing TEHTRIS EDR V2 installation found.")
                return False

        except Exception as e:
            self.logger.warning(f"[ERROR V2] V2 verification failed: {e}")
            return False

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
                time.sleep(3)
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
        """Run complete installation with optimized step handling."""
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

                self.logger.info("Existing installation removed. Pausing before new installation...")
                time.sleep(0.3)

            if not self.validate_prerequisites():
                return False
            time.sleep(0.3)

            if not self.launch_installer():
                return False
            time.sleep(1)

            # Use optimized step-by-step processing
            success = self._run_optimized_installation_steps()
            
            if not success:
                self.logger.error("Installation steps failed")
                return False

            if not self.wait_for_completion():
                return False
            time.sleep(0.3)

            if not self.verify_installation(post_install_check=True):
                return False

            self.logger.info("TEHTRIS EDR installation completed successfully!")
            return True

        except Exception as e:
            self.logger.error(f"Installation failed: {e}")
            return False
    
    def _run_optimized_installation_steps(self) -> bool:
        """Run installation steps with optimized handling for installer lag and step reloads."""
        max_cycles = 30  # Allow more cycles for installation process
        current_cycle = 0
        
        expected_sequence = ['welcome', 'license', 'activation', 'installation']
        current_step_index = 0
        
        while current_cycle < max_cycles and current_step_index < len(expected_sequence):
            current_cycle += 1
            expected_step = expected_sequence[current_step_index]
            current_step = self.detect_current_step()
            
            self.logger.info(f"Cycle {current_cycle}: Expected '{expected_step}', Found '{current_step}'")
            
            # Handle complete step (installation finished)
            if current_step == 'complete':
                self.logger.info("Installation completed detected")
                return True
            
            # If we're at the expected step, handle it
            if current_step == expected_step:
                if expected_step == 'welcome':
                    if not self.handle_welcome_screen():
                        return False
                elif expected_step == 'license':
                    if not self.handle_license_agreement():
                        # License handler returns False if step 1 reload detected
                        if self.detect_current_step() == 'welcome':
                            self.logger.warning("License step detected step 1 reload, restarting from welcome")
                            current_step_index = 0  # Reset to welcome step
                            continue
                        return False
                elif expected_step == 'activation':
                    if not self.handle_activation_information():
                        return False
                elif expected_step == 'installation':
                    if not self.handle_installation():
                        return False
                
                current_step_index += 1
                time.sleep(0.8)  # Longer wait to prevent missclicks
                
            # If installer went back to welcome (lag scenario)
            elif current_step == 'welcome' and expected_step != 'welcome':
                self.logger.warning("Installer reloaded to welcome screen, restarting from step 1")
                current_step_index = 0  # Reset to welcome step
                
            # If we're ahead of expected step, advance the index
            elif current_step in expected_sequence and expected_sequence.index(current_step) > current_step_index:
                self.logger.info(f"Skipping to detected step: {current_step}")
                current_step_index = expected_sequence.index(current_step)
                
            # Unknown or problematic step
            elif current_step == 'unknown':
                self.logger.warning("Unknown step detected, waiting...")
                time.sleep(0.5)
                
            else:
                self.logger.warning(f"Unexpected step '{current_step}' when expecting '{expected_step}'")
                time.sleep(0.5)
        
        if current_cycle >= max_cycles:
            self.logger.error(f"Maximum cycles reached without completing all steps")
            return False
            
        return True

def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(description="TEHTRIS EDR Installer")
    parser.add_argument("installer_path", help="Path to the TEHTRIS EDR installer (MSI or EXE)")
    parser.add_argument("--uninstall-password", help="Password for uninstalling a previous version")
    parser.add_argument("--uninstall-key-file", help="Key file for uninstalling a previous version")

    args = parser.parse_args()

    installer = TehtrisEDRInstaller(
        installer_path=args.installer_path,
        uninstall_password=args.uninstall_password,
        uninstall_key_file=args.uninstall_key_file
    )

    success = installer.run_installation()
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()
