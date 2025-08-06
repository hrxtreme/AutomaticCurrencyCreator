import sys
import time
import pandas as pd
import traceback
import json
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QComboBox, QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog,
    QMessageBox, QCheckBox
)
from PyQt6.QtGui import QFont
from PyQt6.QtCore import Qt, QThread, QObject, pyqtSignal

from playwright.sync_api import sync_playwright, Page, expect

CONFIG_FILE = "config.json"

# =============================================================================
# Browser Driver Class
# =============================================================================
class BrowserDriver:
    """
    Manages all browser interactions using Playwright.
    """
    def __init__(self, progress_callback):
        self.playwright = None
        self.browser = None
        self.page = None
        self.progress_callback = progress_callback

    def launch(self):
        self.progress_callback("Launching browser...")
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(headless=False, slow_mo=50)
        self.page = self.browser.new_page()
        self.progress_callback("Browser launched successfully.")

    def login(self, url, email, password):
        if not self.page:
            raise Exception("Browser is not launched. Call launch() first.")

        self.progress_callback(f"Navigating to login page: {url}")
        self.page.goto(url, timeout=60000)

        self.progress_callback("Entering email...")
        email_locator = self.page.locator('#input28')
        expect(email_locator).to_be_visible(timeout=30000)
        email_locator.fill(email)
        
        self.progress_callback("Clicking 'Next' button...")
        next_button_locator = self.page.locator('input[value="Next"]')
        expect(next_button_locator).to_be_visible(timeout=30000)
        next_button_locator.click()
        self.progress_callback("Email submitted.")

        self.progress_callback("Entering password...")
        password_locator = self.page.locator('#input29')
        expect(password_locator).to_be_visible(timeout=30000)
        password_locator.fill(password)
        
        self.progress_callback("Clicking 'Verify' button...")
        verify_button_locator = self.page.locator('input[value="Verify"]')
        expect(verify_button_locator).to_be_visible(timeout=30000)
        verify_button_locator.click()
        self.progress_callback("Password submitted.")

        self.progress_callback("Looking for MFA options...")
        try:
            push_option_locator = self.page.locator('[aria-label="Select to get a push notification to the Okta Verify app."]')
            expect(push_option_locator).to_be_visible(timeout=30000)
            self.progress_callback("Push notification option found. Clicking it.")
            push_option_locator.click()

            self.progress_callback("Waiting for push sent confirmation...")
            push_sent_locator = self.page.get_by_text("We've sent a push notification", exact=False)
            expect(push_sent_locator).to_be_visible(timeout=15000)
            self.progress_callback("Confirmation received: Push notification sent.")
        except Exception as e:
            self.progress_callback(f"[INFO] Did not find MFA selection screen, or an error occurred. Assuming push was sent by default. Details: {e}")

        self.progress_callback("Waiting for Multi-Factor Authentication (MFA)...")
        self.progress_callback(">>> Please approve the notification on your phone. <<<")
        
        self.page.wait_for_url(lambda url: "okta.com" not in url, timeout=120000)
        self.progress_callback("MFA approved. Login successful!")

    def wait_for_page_to_settle(self):
        """
        Waits for the page to be fully loaded and network idle.
        This prevents race conditions after login or navigation.
        """
        self.progress_callback("Waiting for page to fully load...")
        self.page.wait_for_load_state("networkidle", timeout=30000)
        self.progress_callback("Page has settled.")

    def create_standard_user(self, base_url, user_details, user_password):
        if not self.page:
            raise Exception("Browser is not launched.")

        create_url = f"{base_url}/CreateUserAccount"
        self.progress_callback(f"Navigating to Create User page: {create_url}")
        self.page.goto(create_url, timeout=60000)
        self.wait_for_page_to_settle()

        username = f"{user_details['Username']}{user_details['Postfix']}"
        currency_code = str(user_details['Currency']).split(' ')[0]

        self.progress_callback(f"Creating user: {username} with currency {currency_code}")

        # 1. Player Type is skipped as requested.
        self.progress_callback("  - Player Type: Skipped.")

        # 2. Market Selection
        self.progress_callback("  - Setting Market...")
        market_dropdown_button = self.page.locator('//*[@id="accountForm"]/div/div[2]/div/div/input')
        expect(market_dropdown_button).to_be_visible(timeout=15000)
        market_dropdown_button.click()
        # --- MODIFIED: Using a more specific locator to find the visible <span> ---
        self.page.locator("span", has_text="DEF (No Regulated Market)").click()
        
        # 3. Product Selection
        self.progress_callback("  - Setting Product...")
        product_dropdown_button = self.page.locator('//*[@id="accountForm"]/div/div[3]/div/div/input')
        expect(product_dropdown_button).to_be_visible(timeout=15000)
        product_dropdown_button.click()
        # --- MODIFIED: Using a more specific locator to find the visible <span> ---
        self.page.locator("span", has_text="Island Paradise Mobile (5007)").click()
        
        # 4. Username
        self.progress_callback("  - Filling Username...")
        username_locator = self.page.locator("#username")
        expect(username_locator).to_be_visible(timeout=15000)
        username_locator.fill(username)
        
        # 5. Password
        self.progress_callback("  - Filling Password...")
        password_locator = self.page.locator("#password")
        expect(password_locator).to_be_visible(timeout=15000)
        password_locator.fill(user_password)
        
        # --- These locators are still placeholders ---
        self.progress_callback("  - Setting Currency...")
        currency_locator = self.page.locator("#placeholder-currency")
        expect(currency_locator).to_be_visible(timeout=30000)
        currency_locator.select_option(label=currency_code)
        
        self.progress_callback("  - Clicking 'Create Account'...")
        create_button_locator = self.page.locator("#placeholder-create-button")
        expect(create_button_locator).to_be_visible(timeout=30000)
        create_button_locator.click()

        success_locator = self.page.locator("text=User account created successfully")
        expect(success_locator).to_be_visible(timeout=30000)
        self.progress_callback(f"Successfully created user: {username}")

    def close(self):
        if self.browser:
            self.browser.close()
            self.progress_callback("Browser closed.")
        if self.playwright:
            self.playwright.stop()

# =============================================================================
# Helper Functions
# =============================================================================
def parse_user_data(file_path):
    required_sheet = 'Sheet1'
    try:
        xls = pd.ExcelFile(file_path)
    except FileNotFoundError:
        raise FileNotFoundError(f"The file could not be found at the path: {file_path}")
    except Exception as e:
        raise IOError(f"The file at {file_path} could not be opened or is corrupted. Details: {e}")
    if required_sheet not in xls.sheet_names:
        raise ValueError(f"A required sheet named '{required_sheet}' was not found in the Excel file.")
    df = pd.read_excel(file_path, sheet_name=required_sheet, header=None)
    lvc_df = df.iloc[:, 0:3].copy()
    lvc_df.columns = lvc_df.iloc[0]
    lvc_df = lvc_df[1:].reset_index(drop=True)
    lvc_df.rename(columns={'LVC Currency': 'Currency'}, inplace=True)
    lvc_df.dropna(subset=['Currency'], inplace=True)
    standard_df = df.iloc[:, 4:7].copy()
    standard_df.columns = standard_df.iloc[0]
    standard_df = standard_df[1:].reset_index(drop=True)
    standard_df.dropna(subset=['Currency'], inplace=True)
    required_columns = ['Currency', 'Username', 'Postfix']
    if not all(col in lvc_df.columns for col in required_columns):
        raise ValueError("The LVC section (Columns A-C) is missing required headers: 'LVC Currency', 'Username', 'Postfix'.")
    if not all(col in standard_df.columns for col in required_columns):
        raise ValueError("The Standard section (Columns E-G) is missing required headers: 'Currency', 'Username', 'Postfix'.")
    lvc_users = lvc_df.to_dict('records')
    standard_users = standard_df.to_dict('records')
    return lvc_users, standard_users

def parse_credentials_file(file_path):
    """Reads email and password from a simple text file."""
    with open(file_path, 'r') as f:
        lines = f.readlines()
    if len(lines) < 2:
        raise ValueError("Credentials file must have at least two lines (email and password).")
    email = lines[0].strip()
    password = lines[1].strip()
    return email, password

# =============================================================================
# Automation Worker (Controller)
# =============================================================================
class AutomationWorker(QObject):
    progress_update = pyqtSignal(str)
    automation_error = pyqtSignal(str)
    automation_finished = pyqtSignal()

    def __init__(self, gtp_url, email, password, lvc_users, standard_users, user_password, debug_mode):
        super().__init__()
        self.gtp_url = gtp_url
        self.email = email
        self.password = password
        self.lvc_users = lvc_users
        self.standard_users = standard_users
        self.user_password = user_password
        self.debug_mode = debug_mode
        self.is_running = True
        self.driver = None

    def run(self):
        self.driver = BrowserDriver(progress_callback=self.progress_update.emit)
        try:
            self.progress_update.emit("Automation thread started.")
            if self.debug_mode:
                self.progress_update.emit("--- DEBUG MODE ENABLED: Processing first LVC user only. ---")
            if not self.is_running: return

            self.driver.launch()
            if not self.is_running: return
            
            self.driver.login(self.gtp_url, self.email, self.password)
            if not self.is_running: return
            
            self.driver.wait_for_page_to_settle()
            if not self.is_running: return

            self.progress_update.emit("\nStep 2: Processing LVC users...")
            for user in self.lvc_users:
                if not self.is_running: break
                self.driver.create_standard_user(self.gtp_url, user, self.user_password)
                if self.debug_mode:
                    self.progress_update.emit("--- DEBUG MODE: Halting after first LVC user. ---")
                    break
            
            if not self.debug_mode:
                if not self.is_running: return
                self.progress_update.emit("\nStep 3: Processing Standard users...")
                for user in self.standard_users:
                    if not self.is_running: break
                    self.driver.create_standard_user(self.gtp_url, user, self.user_password)

            if self.is_running:
                self.progress_update.emit("\nAutomation complete!")

        except Exception as e:
            print("--- A CRITICAL ERROR OCCURRED IN THE AUTOMATION WORKER ---")
            traceback.print_exc()
            print("---------------------------------------------------------")
            
            error_type = type(e).__name__
            error_details = str(e)
            full_error_message = (
                f"A critical error stopped the automation.\n\n"
                f"Error Type: {error_type}\n"
                f"Details: {error_details}"
            )
            self.progress_update.emit(f"[ERROR] {full_error_message}")
            self.automation_error.emit(full_error_message)
        finally:
            if self.driver:
                self.driver.close()
            self.automation_finished.emit()

    def stop(self):
        self.progress_update.emit("Stopping process...")
        self.is_running = False

# =============================================================================
# Main Application Window (UI)
# =============================================================================
class MainWindow(QMainWindow):
    GTP_VERSIONS = {
        "GTP 640": "https://admin-app1-gtp640.installprogram.eu",
        "GTP 641": "https://admin-app1-gtp641.installprogram.eu",
        "GTP 642": "https://admin-app1-gtp642.installprogram.eu",
        "GTP 643": "https://admin-app1-gtp643.installprogram.eu",
        "GTP 644": "https://admin-app1-gtp644.installprogram.eu",
    }

    def __init__(self):
        super().__init__()
        self.automation_thread = None
        self.worker = None
        self.config = {}
        self.setWindowTitle("GTP User Automation Tool v1.0")
        self.setGeometry(100, 100, 700, 550)
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setSpacing(15)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        header_font = QFont()
        header_font.setPointSize(12)
        header_font.setBold(True)
        self.load_config()
        self.setup_ui_elements(header_font)
        self.setup_connections()
        self.apply_config()

    def setup_ui_elements(self, header_font):
        config_header = QLabel("1. GTP Configuration")
        config_header.setFont(header_font)
        self.main_layout.addWidget(config_header)
        self.gtp_dropdown = QComboBox()
        self.gtp_dropdown.addItems(self.GTP_VERSIONS.keys())
        self.main_layout.addWidget(self.gtp_dropdown)
        
        login_header = QLabel("2. File & Password Selection")
        login_header.setFont(header_font)
        self.main_layout.addWidget(login_header)

        cred_layout = QHBoxLayout()
        self.select_cred_button = QPushButton("Select Credentials File (.txt)")
        self.cred_path_label = QLabel("No file selected.")
        self.cred_path_label.setStyleSheet("font-style: italic; color: #555;")
        cred_layout.addWidget(self.select_cred_button)
        cred_layout.addWidget(self.cred_path_label, 1)
        self.main_layout.addLayout(cred_layout)

        user_layout = QHBoxLayout()
        self.select_user_file_button = QPushButton("Select User Data File (.xlsx)")
        self.user_file_path_label = QLabel("No file selected.")
        self.user_file_path_label.setStyleSheet("font-style: italic; color: #555;")
        user_layout.addWidget(self.select_user_file_button)
        user_layout.addWidget(self.user_file_path_label, 1)
        self.main_layout.addLayout(user_layout)
        
        user_pass_label = QLabel("Default Password for New User Accounts:")
        self.main_layout.addWidget(user_pass_label)
        self.user_password_input = QLineEdit()
        self.user_password_input.setText("snow") # Pre-filled as requested
        self.main_layout.addWidget(self.user_password_input)

        self.debug_mode_checkbox = QCheckBox("Debug Mode (Process first LVC user only)")
        self.main_layout.addWidget(self.debug_mode_checkbox)
        
        self.start_button = QPushButton("Start Automation")
        self.start_button.setStyleSheet(
            "QPushButton { background-color: #4CAF50; color: white; padding: 10px; border-radius: 5px; font-size: 14px; }"
            "QPushButton:hover { background-color: #45a049; }"
        )
        self.main_layout.addWidget(self.start_button)
        
        log_header = QLabel("Status Log")
        log_header.setFont(header_font)
        self.main_layout.addWidget(log_header)
        self.status_log = QTextEdit()
        self.status_log.setReadOnly(True)
        self.main_layout.addWidget(self.status_log)

    def setup_connections(self):
        self.select_cred_button.clicked.connect(self.select_credentials_file)
        self.select_user_file_button.clicked.connect(self.select_user_data_file)
        self.start_button.clicked.connect(self.start_automation)

    def start_automation(self):
        gtp_selection = self.gtp_dropdown.currentText()
        gtp_url = self.GTP_VERSIONS[gtp_selection]
        cred_path = self.cred_path_label.text()
        user_data_path = self.user_file_path_label.text()
        user_password = self.user_password_input.text()
        debug_mode = self.debug_mode_checkbox.isChecked()

        errors = []
        if "No file selected" in cred_path: errors.append("You must select a credentials file.")
        if "No file selected" in user_data_path: errors.append("You must select a user data file.")
        if not user_password: errors.append("The password for new users cannot be empty.")
        if errors:
            QMessageBox.warning(self, "Input Error", "\n".join(errors))
            return

        self.log_message("="*50)
        self.log_message("Starting pre-flight checks...")
        
        try:
            self.log_message(f"Parsing credentials file: {cred_path}")
            email, password = parse_credentials_file(cred_path)
            self.log_message("  - Credentials parsed successfully.")
            
            self.log_message(f"Parsing user data file: {user_data_path}")
            lvc_users, standard_users = parse_user_data(user_data_path)
            self.log_message(f"  - Validation successful: Found {len(lvc_users)} LVC and {len(standard_users)} Standard users.")
        except Exception as e:
            print("--- A FILE PARSING ERROR OCCURRED ---")
            traceback.print_exc()
            print("-------------------------------------")
            error_message = f"Failed to read or validate an input file.\n\nError: {e}"
            self.log_message(f"[ERROR] {error_message}")
            QMessageBox.critical(self, "File Error", error_message)
            return

        self.log_message("Pre-flight checks passed. Starting automation process...")
        self.toggle_controls(False)

        self.automation_thread = QThread()
        self.worker = AutomationWorker(gtp_url, email, password, lvc_users, standard_users, user_password, debug_mode)
        self.worker.moveToThread(self.automation_thread)

        self.worker.progress_update.connect(self.log_message)
        self.worker.automation_error.connect(self.on_automation_error)
        self.worker.automation_finished.connect(self.on_automation_finished)

        self.automation_thread.started.connect(self.worker.run)
        self.automation_thread.finished.connect(self.automation_thread.deleteLater)
        
        self.automation_thread.start()

    def on_automation_error(self, error_message):
        QMessageBox.critical(self, "Automation Error", error_message)

    def on_automation_finished(self):
        self.log_message("Process finished.")
        if self.automation_thread is not None:
            self.automation_thread.quit()
            self.automation_thread.wait()
        self.toggle_controls(True)
        self.automation_thread = None
        self.worker = None

    def select_credentials_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Credentials File", "", "Text Files (*.txt)")
        if file_path:
            self.cred_path_label.setText(file_path)
            self.cred_path_label.setStyleSheet("font-style: normal; color: #000;")
            self.log_message(f"Selected credentials file: {file_path}")
            self.config['credentials_path'] = file_path
            self.save_config()

    def select_user_data_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select User Data File", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.user_file_path_label.setText(file_path)
            self.user_file_path_label.setStyleSheet("font-style: normal; color: #000;")
            self.log_message(f"Selected user data file: {file_path}")
            self.config['user_data_path'] = file_path
            self.save_config()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                self.config = json.load(f)
        else:
            self.config = {}

    def save_config(self):
        with open(CONFIG_FILE, 'w') as f:
            json.dump(self.config, f, indent=4)

    def apply_config(self):
        """Applies loaded configuration to the UI."""
        cred_path = self.config.get('credentials_path')
        if cred_path and os.path.exists(cred_path):
            self.cred_path_label.setText(cred_path)
            self.cred_path_label.setStyleSheet("font-style: normal; color: #000;")
        
        user_data_path = self.config.get('user_data_path')
        if user_data_path and os.path.exists(user_data_path):
            self.user_file_path_label.setText(user_data_path)
            self.user_file_path_label.setStyleSheet("font-style: normal; color: #000;")

    def log_message(self, message):
        self.status_log.append(message)

    def toggle_controls(self, enabled):
        self.gtp_dropdown.setEnabled(enabled)
        self.select_cred_button.setEnabled(enabled)
        self.select_user_file_button.setEnabled(enabled)
        self.user_password_input.setEnabled(enabled)
        self.start_button.setEnabled(enabled)
        self.debug_mode_checkbox.setEnabled(enabled)
        
    def closeEvent(self, event):
        self.save_config() # Save config on close
        if self.automation_thread and self.automation_thread.isRunning():
            self.worker.stop()
            self.automation_thread.quit()
            self.automation_thread.wait()
        event.accept()

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
