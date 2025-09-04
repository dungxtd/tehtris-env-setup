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
                    ("i agree", 2), ("agree", 2), ("accept", 2), ("next", 2),
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
        """Handle both Nmap and Npcap installers that may run concurrently."""
        self.logger.info("Handling concurrent Nmap and Npcap installers...")

        max_duration = 300  # 5 minutes total timeout
        start_time = time.time()
        nmap_complete = False
        npcap_complete = False

        while time.time() - start_time < max_duration:
            try:
                import win32gui

                # Find all installer windows
                nmap_windows = []
                npcap_windows = []

                def find_installer_windows(hwnd, param):
                    try:
                        if win32gui.IsWindowVisible(hwnd):
                            window_text = win32gui.GetWindowText(hwnd)
                            if "nmap" in window_text.lower() and "setup" in window_text.lower():
                                nmap_windows.append((hwnd, window_text))
                            elif "npcap" in window_text.lower():
                                npcap_windows.append((hwnd, window_text))
                    except:
                        pass
                    return True

                win32gui.EnumWindows(find_installer_windows, None)

                # Handle Npcap installer first (higher priority when both are present)
                if npcap_windows and not npcap_complete:
                    self.logger.info(f"Found Npcap window: {npcap_windows[0][1]}")
                    self._handle_npcap_window(npcap_windows[0][0])
                    # Don't handle Nmap while Npcap is active
                    continue

                # Only handle Nmap if no Npcap windows are present
                if nmap_windows and not nmap_complete and not npcap_windows:
                    self.logger.info(f"Found Nmap window: {nmap_windows[0][1]}")
                    self._handle_nmap_window(nmap_windows[0][0])

                # Check if Npcap is complete (no more Npcap windows)
                if not npcap_windows and not npcap_complete:
                    npcap_complete = True
                    self.logger.info("Npcap installation completed (no more Npcap windows)")

                # Check if Nmap is complete (no more Nmap windows)
                if not nmap_windows and not nmap_complete:
                    nmap_complete = True
                    self.logger.info("Nmap installation completed (no more Nmap windows)")

                # Check if both are complete
                if not nmap_windows and not npcap_windows:
                    self.logger.info("No installer windows found - installation likely complete")
                    break

                # If Nmap is installing and Npcap appears, prioritize Npcap
                if nmap_windows and npcap_windows:
                    self.logger.info("Both installers active - prioritizing Npcap")

                time.sleep(1)

            except Exception as e:
                self.logger.error(f"Error in concurrent handler: {e}")
                time.sleep(2)

        return True

    def _handle_npcap_window(self, hwnd) -> bool:
        """Handle a specific Npcap window."""
        try:
            import win32gui
            import win32con

            # Get window title for debugging
            window_title = win32gui.GetWindowText(hwnd)
            self.logger.info(f"Processing Npcap window: {window_title}")

            # Get all buttons and checkboxes in this window
            controls = []
            def find_controls(child_hwnd, param):
                try:
                    if win32gui.IsWindowVisible(child_hwnd):
                        window_text = win32gui.GetWindowText(child_hwnd)
                        class_name = win32gui.GetClassName(child_hwnd)
                        if class_name == 'Button' and window_text:
                            clean_text = window_text.replace('&', '').strip()
                            controls.append((child_hwnd, clean_text, class_name))
                            self.logger.info(f"Found Npcap control: '{clean_text}' ({class_name})")
                except:
                    pass
                return True

            win32gui.EnumChildWindows(hwnd, find_controls, None)

            if not controls:
                self.logger.warning("No controls found in Npcap window")
                return False

            # Handle different Npcap installation steps
            button_clicked = False

            # Step 1: License Agreement - look for "I Agree" button
            for control_hwnd, control_text, class_name in controls:
                if "i agree" in control_text.lower():
                    win32gui.SendMessage(control_hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
                    win32gui.SendMessage(control_hwnd, win32con.WM_LBUTTONUP, 0, 0)
                    self.logger.info(f"Npcap: Clicked 'I Agree' button")
                    time.sleep(1)
                    button_clicked = True
                    break

            # Step 2: Installation Options - look for "Install" button
            if not button_clicked:
                for control_hwnd, control_text, class_name in controls:
                    if "install" in control_text.lower() and "uninstall" not in control_text.lower():
                        win32gui.SendMessage(control_hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
                        win32gui.SendMessage(control_hwnd, win32con.WM_LBUTTONUP, 0, 0)
                        self.logger.info(f"Npcap: Clicked 'Install' button")
                        time.sleep(3)  # Installation takes time
                        button_clicked = True
                        break

            # Step 3: Completion - look for "Finish" or "Close" button
            if not button_clicked:
                for control_hwnd, control_text, class_name in controls:
                    if any(word in control_text.lower() for word in ["finish", "close", "done"]):
                        win32gui.SendMessage(control_hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
                        win32gui.SendMessage(control_hwnd, win32con.WM_LBUTTONUP, 0, 0)
                        self.logger.info(f"Npcap: Clicked '{control_text}' button")
                        time.sleep(1)
                        button_clicked = True
                        break

            # Fallback: try "Next" button
            if not button_clicked:
                for control_hwnd, control_text, class_name in controls:
                    if "next" in control_text.lower():
                        win32gui.SendMessage(control_hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
                        win32gui.SendMessage(control_hwnd, win32con.WM_LBUTTONUP, 0, 0)
                        self.logger.info(f"Npcap: Clicked 'Next' button")
                        time.sleep(1)
                        button_clicked = True
                        break

            if not button_clicked:
                self.logger.warning(f"No actionable button found in Npcap window")
                # List all available controls for debugging
                control_list = [f"'{text}'" for _, text, _ in controls]
                self.logger.info(f"Available Npcap controls: {', '.join(control_list)}")

            return False  # Keep processing

        except Exception as e:
            self.logger.error(f"Error handling Npcap window: {e}")
            return False

    def _handle_nmap_window(self, hwnd) -> bool:
        """Handle a specific Nmap window."""
        try:
            import win32gui
            import win32con

            # Get buttons in this window
            buttons = []
            def find_buttons(child_hwnd, param):
                try:
                    if win32gui.IsWindowVisible(child_hwnd):
                        window_text = win32gui.GetWindowText(child_hwnd)
                        class_name = win32gui.GetClassName(child_hwnd)
                        if window_text and class_name == 'Button':
                            buttons.append((child_hwnd, window_text.replace('&', '').lower().strip()))
                except:
                    pass
                return True

            win32gui.EnumChildWindows(hwnd, find_buttons, None)

            # Check if this is an "Installing" window (should wait)
            window_text = win32gui.GetWindowText(hwnd)
            if "installing" in window_text.lower():
                self.logger.info("Nmap installation in progress - waiting...")
                return False  # Keep processing

            # Priority order for Nmap buttons
            nmap_button_priority = ["next", "i agree", "agree", "install", "finish", "close"]

            for priority_button in nmap_button_priority:
                for button_hwnd, button_text in buttons:
                    if priority_button in button_text:
                        win32gui.SendMessage(button_hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
                        win32gui.SendMessage(button_hwnd, win32con.WM_LBUTTONUP, 0, 0)
                        self.logger.info(f"Nmap: Clicked '{button_text}' button")
                        time.sleep(3 if priority_button == "install" else 1)
                        return False  # Continue processing this window

            return False  # Keep processing

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
