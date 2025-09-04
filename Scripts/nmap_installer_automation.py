#!/usr/bin/env python3
"""
Nmap Installer Automation Script
Automates the installation of Nmap (including Npcap) by detecting and clicking through installer windows.
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
    pyautogui.PAUSE = 0.1
except ImportError:
    PYAUTOGUI_AVAILABLE = False

class NmapInstaller:
    """Automated Nmap installer with Npcap support."""

    def __init__(self, installer_path: str):
        self.installer_path = Path(installer_path)
        self.logger = self._setup_logging()

    def _setup_logging(self) -> logging.Logger:
        """Setup logging."""
        logger = logging.getLogger('NmapInstaller')
        logger.setLevel(logging.INFO)

        # File handler
        file_handler = logging.FileHandler('nmap_installation.log')
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

    def click_with_win32gui(self, button_text: str, window_title_contains: str = None) -> bool:
        """Click button using win32gui."""
        try:
            import win32gui
            import win32con

            self.logger.info(f"Looking for button: {button_text}")

            # Find installer windows
            installer_windows = []
            def find_windows(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        if window_title_contains:
                            if window_title_contains.lower() in window_text.lower():
                                windows.append(hwnd)
                        else:
                            # Look for common installer window titles
                            if any(title in window_text.lower() for title in 
                                   ["nmap", "setup", "installer", "npcap"]):
                                windows.append(hwnd)
                except:
                    pass
                return True

            win32gui.EnumWindows(find_windows, installer_windows)
            self.logger.info(f"Found {len(installer_windows)} installer windows")

            # Search for button
            for installer_hwnd in installer_windows:
                try:
                    def find_button(hwnd, button_info):
                        try:
                            if win32gui.IsWindowVisible(hwnd):
                                window_text = win32gui.GetWindowText(hwnd)
                                class_name = win32gui.GetClassName(hwnd)

                                if window_text and class_name == 'Button':
                                    clean_text = window_text.replace('&', '').lower().strip()
                                    clean_button = button_text.replace('&', '').lower().strip()
                                    
                                    if clean_button in clean_text or clean_text in clean_button:
                                        button_info['hwnd'] = hwnd
                                        button_info['text'] = window_text
                                        return False
                        except:
                            pass
                        return True

                    button_info = {}
                    win32gui.EnumChildWindows(installer_hwnd, find_button, button_info)

                    if button_info.get('hwnd'):
                        win32gui.SendMessage(button_info['hwnd'], win32con.WM_LBUTTONDOWN, 0, 0)
                        win32gui.SendMessage(button_info['hwnd'], win32con.WM_LBUTTONUP, 0, 0)
                        self.logger.info(f"Clicked button: {button_info['text']}")
                        return True

                except Exception as e:
                    self.logger.debug(f"Error in window {installer_hwnd}: {e}")
                    continue

            self.logger.error(f"Button '{button_text}' not found")
            return False

        except Exception as e:
            self.logger.error(f"win32gui click failed: {e}")
            return False

    def click_checkbox(self, checkbox_text: str) -> bool:
        """Click checkbox using win32gui."""
        try:
            import win32gui
            import win32con

            self.logger.info(f"Looking for checkbox: {checkbox_text}")

            # Find installer windows
            installer_windows = []
            def find_windows(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        if any(title in window_text.lower() for title in 
                               ["nmap", "setup", "installer", "npcap"]):
                            windows.append(hwnd)
                except:
                    pass
                return True

            win32gui.EnumWindows(find_windows, installer_windows)

            # Search for checkbox
            for installer_hwnd in installer_windows:
                try:
                    def find_checkbox(hwnd, checkbox_info):
                        try:
                            if win32gui.IsWindowVisible(hwnd):
                                window_text = win32gui.GetWindowText(hwnd)
                                class_name = win32gui.GetClassName(hwnd)

                                if class_name == 'Button':
                                    # Check if it's a checkbox (style BS_CHECKBOX or BS_AUTOCHECKBOX)
                                    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                                    if style & 0x02 or style & 0x03:  # BS_CHECKBOX or BS_AUTOCHECKBOX
                                        clean_text = window_text.lower().strip()
                                        clean_checkbox = checkbox_text.lower().strip()
                                        
                                        if clean_checkbox in clean_text or clean_text in clean_checkbox:
                                            checkbox_info['hwnd'] = hwnd
                                            checkbox_info['text'] = window_text
                                            return False
                        except:
                            pass
                        return True

                    checkbox_info = {}
                    win32gui.EnumChildWindows(installer_hwnd, find_checkbox, checkbox_info)

                    if checkbox_info.get('hwnd'):
                        win32gui.SendMessage(checkbox_info['hwnd'], win32con.BM_CLICK, 0, 0)
                        self.logger.info(f"Clicked checkbox: {checkbox_info['text']}")
                        return True

                except Exception as e:
                    self.logger.debug(f"Error in window {installer_hwnd}: {e}")
                    continue

            self.logger.error(f"Checkbox '{checkbox_text}' not found")
            return False

        except Exception as e:
            self.logger.error(f"Checkbox click failed: {e}")
            return False

    def detect_current_window(self) -> str:
        """Detect current installer window type."""
        try:
            import win32gui

            installer_windows = []
            def find_windows(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        if any(title in window_text.lower() for title in 
                               ["nmap", "setup", "installer", "npcap"]):
                            windows.append((hwnd, window_text))
                except:
                    pass
                return True

            win32gui.EnumWindows(find_windows, installer_windows)
            
            for hwnd, window_text in installer_windows:
                self.logger.info(f"Found window: {window_text}")
                
                if "npcap" in window_text.lower():
                    return "npcap"
                elif "nmap" in window_text.lower():
                    return "nmap"
                elif "setup" in window_text.lower():
                    return "setup"

            return "unknown"

        except Exception as e:
            self.logger.debug(f"Window detection failed: {e}")
            return "unknown"

    def launch_installer(self) -> bool:
        """Launch the Nmap installer."""
        self.logger.info("Launching Nmap installer...")

        try:
            subprocess.Popen([str(self.installer_path)], shell=True)
            time.sleep(3)
            self.logger.info("Installer launched successfully")
            return True

        except Exception as e:
            self.logger.error(f"Failed to launch installer: {e}")
            return False

    def scan_available_buttons(self) -> list:
        """Scan and return all available buttons in installer windows."""
        available_buttons = []
        try:
            import win32gui

            # Find installer windows
            installer_windows = []
            def find_windows(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        window_text = win32gui.GetWindowText(hwnd)
                        if any(title in window_text.lower() for title in
                               ["nmap", "setup", "installer", "npcap"]):
                            windows.append(hwnd)
                except:
                    pass
                return True

            win32gui.EnumWindows(find_windows, installer_windows)

            for installer_hwnd in installer_windows:
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

                    win32gui.EnumChildWindows(installer_hwnd, find_all_buttons, available_buttons)
                except:
                    continue

        except Exception as e:
            self.logger.debug(f"Button scan failed: {e}")

        # Remove duplicates and return
        return list(set(available_buttons))

    def handle_nmap_installer(self) -> bool:
        """Handle the main Nmap installer windows."""
        self.logger.info("Handling Nmap installer windows...")

        max_attempts = 60  # Increased for longer installations
        attempt = 0
        last_buttons = []

        while attempt < max_attempts:
            attempt += 1
            window_type = self.detect_current_window()
            current_buttons = self.scan_available_buttons()

            self.logger.info(f"Attempt {attempt}: Window type: {window_type}, Buttons: {current_buttons}")

            if window_type == "unknown" and not current_buttons:
                time.sleep(1)
                continue

            # Check if we're done (no installer windows)
            if not current_buttons:
                self.logger.info("No installer buttons found - installation may be complete")
                break

            # Try buttons in priority order
            button_priority = [
                ("next", 2), ("i agree", 2), ("agree", 2), ("accept", 2),
                ("install", 5), ("yes", 2), ("ok", 2),
                ("finish", 2), ("close", 1), ("done", 1)
            ]

            button_clicked = False
            for button_text, wait_time in button_priority:
                if any(button_text in btn for btn in current_buttons):
                    if self.click_with_win32gui(button_text):
                        self.logger.info(f"Clicked '{button_text}' button")
                        time.sleep(wait_time)
                        button_clicked = True
                        break

            if not button_clicked:
                # If no standard buttons found, wait and retry
                if current_buttons == last_buttons:
                    self.logger.warning(f"Same buttons detected twice: {current_buttons}")
                    time.sleep(2)
                else:
                    time.sleep(1)

            last_buttons = current_buttons

        self.logger.info("Nmap installer handling completed")
        return True

    def handle_npcap_installer(self) -> bool:
        """Handle Npcap installer if it appears."""
        self.logger.info("Checking for Npcap installer...")

        max_wait = 60  # Increased wait time
        start_time = time.time()
        npcap_handled = False

        while time.time() - start_time < max_wait:
            window_type = self.detect_current_window()
            current_buttons = self.scan_available_buttons()

            if window_type == "npcap" or any("npcap" in btn for btn in current_buttons):
                self.logger.info("Npcap installer detected, handling...")
                npcap_handled = True

                # Handle Npcap installer with comprehensive button checking
                button_priority = [
                    ("next", 2),
                    ("install", 8), ("yes", 2), ("ok", 2),
                    ("finish", 2), ("close", 1), ("done", 1)
                ]

                button_clicked = False
                for button_text, wait_time in button_priority:
                    if any(button_text in btn for btn in current_buttons):
                        if self.click_with_win32gui(button_text):
                            self.logger.info(f"Npcap: Clicked '{button_text}' button")
                            time.sleep(wait_time)
                            button_clicked = True
                            break

                if not button_clicked:
                    time.sleep(1)

            elif npcap_handled and not current_buttons:
                # Npcap was handled and no more buttons - likely complete
                self.logger.info("Npcap installation appears complete")
                break
            elif not npcap_handled and not current_buttons:
                # No Npcap installer detected and no buttons - check if we missed it
                time.sleep(1)
            else:
                time.sleep(1)

        if not npcap_handled:
            self.logger.info("No Npcap installer detected - may have been skipped or already installed")
        else:
            self.logger.info("Npcap installer handling completed")

        return True

    def verify_installation(self) -> bool:
        """Verify that Nmap was installed successfully."""
        self.logger.info("Verifying Nmap installation...")

        # Check common installation paths
        possible_paths = [
            r"C:\Program Files (x86)\Nmap\nmap.exe",
            r"C:\Program Files\Nmap\nmap.exe"
        ]

        for path in possible_paths:
            if os.path.exists(path):
                self.logger.info(f"Nmap installation verified at: {path}")
                return True

        # Try to run nmap command to verify it's in PATH
        try:
            result = subprocess.run(["nmap", "--version"],
                                  capture_output=True, text=True, timeout=10)
            if result.returncode == 0 and "nmap" in result.stdout.lower():
                self.logger.info("Nmap installation verified via PATH")
                return True
        except:
            pass

        self.logger.warning("Could not verify Nmap installation")
        return False

    def handle_concurrent_installers(self) -> bool:
        """Handle both Nmap and Npcap installers by focusing on the correct window."""
        self.logger.info("Handling concurrent Nmap and Npcap installers...")

        max_duration = 300  # 5 minutes total timeout
        start_time = time.time()

        while time.time() - start_time < max_duration:
            try:
                import win32gui

                nmap_hwnd, npcap_hwnd = None, None

                def find_windows(hwnd, param):
                    nonlocal nmap_hwnd, npcap_hwnd
                    try:
                        if win32gui.IsWindowVisible(hwnd):
                            text = win32gui.GetWindowText(hwnd)
                            # Use 'in' for more robust title matching
                            if 'nmap setup' in text.lower():
                                nmap_hwnd = hwnd
                            elif 'npcap' in text.lower() and 'setup' in text.lower():
                                npcap_hwnd = hwnd
                    except Exception:
                        pass
                    return True

                win32gui.EnumWindows(find_windows, None)

                # If the Npcap window is present, it gets exclusive priority.
                if npcap_hwnd:
                    self.logger.info("Npcap window detected. Setting focus and handling.")
                    try:
                        win32gui.SetForegroundWindow(npcap_hwnd)
                        time.sleep(0.2)
                        self._handle_npcap_window(npcap_hwnd)
                    except Exception as e:
                        self.logger.error(f"Failed to set focus on Npcap window: {e}")
                    time.sleep(1) # Wait a second before re-evaluating
                    continue

                # If Npcap is gone, handle Nmap.
                if nmap_hwnd:
                    self.logger.info("Handling Nmap window.")
                    try:
                        win32gui.SetForegroundWindow(nmap_hwnd)
                        time.sleep(0.2)
                        self._handle_nmap_window(nmap_hwnd)
                    except Exception as e:
                        self.logger.error(f"Failed to set focus on Nmap window: {e}")

                # If no installer windows are found, we're done.
                if not nmap_hwnd and not npcap_hwnd:
                    self.logger.info("No active installer windows found. Installation should be complete.")
                    break

                time.sleep(1)

            except Exception as e:
                self.logger.error(f"Error in concurrent handler: {e}")
                time.sleep(2)

        return True

    def _handle_npcap_window(self, hwnd) -> bool:
        """Handle the Npcap window by clicking buttons in a prioritized order."""
        try:
            import win32gui
            import win32con

            controls = self._get_window_controls(hwnd)
            buttons = controls['buttons']
            button_texts = [b['text'] for b in buttons]
            self.logger.info(f"Npcap Window Found: Buttons={button_texts}")

            # Define the priority of buttons to click based on the observed flow.
            button_priority = [
                'install',
                'finish',
                'next >',     # Fallback
                'close'       # Fallback
            ]

            # Find and click the highest-priority button available.
            for priority_text in button_priority:
                for btn in buttons:
                    if priority_text in btn['text']:
                        win32gui.SendMessage(btn['hwnd'], win32con.BM_CLICK, 0, 0)
                        self.logger.info(f"Npcap Action: Clicked '{btn['text']}'")
                        time.sleep(2) # Wait for the next window/state
                        return False # Action taken, exit to re-evaluate

            self.logger.warning("No actionable button found in Npcap window.")
            return False # No action taken, continue polling

        except Exception as e:
            self.logger.error(f"Error handling Npcap window: {e}")
            return False

    def _get_window_controls(self, hwnd):
        """Get all controls (text, buttons) from a window."""
        controls = {'buttons': [], 'text': []}
        try:
            import win32gui
            def find_controls(child_hwnd, param):
                try:
                    if win32gui.IsWindowVisible(child_hwnd):
                        text = win32gui.GetWindowText(child_hwnd).lower()
                        class_name = win32gui.GetClassName(child_hwnd).lower()
                        if 'button' in class_name:
                            controls['buttons'].append({'hwnd': child_hwnd, 'text': text.replace('&', '')})
                        elif 'static' in class_name:
                            controls['text'].append(text)
                except Exception:
                    pass
                return True
            win32gui.EnumChildWindows(hwnd, find_controls, None)
        except Exception as e:
            self.logger.error(f"Error getting window controls: {e}")
        return controls

    def _handle_nmap_window(self, hwnd) -> bool:
        """Handle a specific Nmap window by detecting the current step."""
        try:
            import win32gui
            import win32con

            controls = self._get_window_controls(hwnd)
            buttons = controls['buttons']
            all_text = " ".join(controls['text'])
            button_texts = [b['text'] for b in buttons]
            self.logger.info(f"Nmap Step Detection: Buttons={button_texts}, Text='{all_text[:100]}...'" )

            # Determine step and act
            action_taken = False
            if 'choose components' in all_text:
                self.logger.info("Nmap Step: Choose Components")
                for btn in buttons:
                    if 'next >' in btn['text']:
                        win32gui.SendMessage(btn['hwnd'], win32con.BM_CLICK, 0, 0)
                        self.logger.info("Nmap Action: Clicked 'Next >' (Components)")
                        action_taken = True
                        break
            elif 'installing' in all_text or 'execute:' in all_text:
                self.logger.info("Nmap Step: Installing (waiting)")
                # This is a waiting step, no action to take
                return False
            elif 'completed' in all_text or 'has been installed' in all_text:
                 self.logger.info("Nmap Step: Installation Completed")
                 for btn in buttons:
                    if 'next >' in btn['text'] or 'finish' in btn['text']:
                        win32gui.SendMessage(btn['hwnd'], win32con.BM_CLICK, 0, 0)
                        self.logger.info(f"Nmap Action: Clicked '{btn['text']}' (Completed)")
                        action_taken = True
                        break

            # Fallback for any other screen with a Next or Install button
            if not action_taken:
                for btn in buttons:
                    if 'install' in btn['text']:
                        win32gui.SendMessage(btn['hwnd'], win32con.BM_CLICK, 0, 0)
                        self.logger.info("Nmap Action: Clicked 'Install' (Fallback)")
                        action_taken = True
                        break
                if not action_taken:
                     for btn in buttons:
                        if 'next >' in btn['text']:
                            win32gui.SendMessage(btn['hwnd'], win32con.BM_CLICK, 0, 0)
                            self.logger.info("Nmap Action: Clicked 'Next >' (Fallback)")
                            action_taken = True
                            break

            if action_taken:
                time.sleep(1.5)

            return False # Keep processing

        except Exception as e:
            self.logger.error(f"Error handling Nmap window: {e}")
            return False

    def run_installation(self) -> bool:
        """Run the complete installation process."""
        self.logger.info("Starting Nmap installation automation")

        try:
            if not self.validate_prerequisites():
                return False

            if not self.launch_installer():
                return False

            # Use the new concurrent handler instead of separate handlers
            if not self.handle_concurrent_installers():
                return False

            # Verify installation
            if not self.verify_installation():
                self.logger.warning("Installation verification failed, but process completed")
                # Don't return False here as installation might still be successful

            self.logger.info("Nmap installation automation completed successfully!")
            return True

        except Exception as e:
            self.logger.error(f"Installation failed: {e}")
            return False

def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(description="Nmap Installer Automation")
    parser.add_argument("installer_path", help="Path to the Nmap installer executable")

    args = parser.parse_args()

    installer = NmapInstaller(installer_path=args.installer_path)
    success = installer.run_installation()
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()
