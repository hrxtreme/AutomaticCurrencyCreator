import sys
import time
import pandas as pd
import traceback
import json
import os
import logging
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QComboBox, QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog,
    QMessageBox, QCheckBox, QRadioButton, QGroupBox
)
from PyQt6.QtGui import QFont, QIntValidator
from PyQt6.QtCore import Qt, QThread, QObject, pyqtSignal

from playwright.sync_api import sync_playwright, Page, expect
from openpyxl import load_workbook

CONFIG_FILE = "config.json"
LOG_FILE = "automation_log.txt"

# =============================================================================
# Logging Setup
# =============================================================================
def setup_logging():
    """Sets up logging to a file."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(message)s',
        filename=LOG_FILE,
        filemode='w'  # Overwrite the log file on each run
    )
setup_logging()

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
        
        self.progress_callback("Pausing briefly for dashboard to initialize...")
        self.page.wait_for_timeout(3000)

    def wait_for_page_to_settle(self):
        self.progress_callback("Waiting for page to fully load...")
        self.page.wait_for_load_state("networkidle", timeout=30000)
        self.progress_callback("Page has settled.")

    def navigate_to_create_user_page(self, base_url):
        if not self.page:
            raise Exception("Browser is not launched.")
        create_url = f"{base_url}/CreateUserAccount"
        self.progress_callback(f"Navigating to Create User page: {create_url}")
        self.page.goto(create_url, timeout=60000)
        self.wait_for_page_to_settle()

    def fill_user_creation_form(self, user_details, user_password, postfix):
        username = f"{user_details['Username']}{postfix}"
        
        # --- MODIFIED: More robust currency code extraction from username ---
        username_from_file = str(user_details['Username'])
        # Assuming the currency code is the first 3 letters of the username string (e.g., "KRWx" -> "KRW")
        currency_code = username_from_file[:3]

        self.progress_callback(f"--- Creating user account: {username} ---")

        self.progress_callback("  - Setting Market...")
        market_dropdown_button = self.page.locator('//*[@id="accountForm"]/div/div[2]/div/div/input')
        expect(market_dropdown_button).to_be_visible(timeout=15000)
        market_dropdown_button.click()
        self.page.locator("span", has_text="DEF (No Regulated Market)").click()
        
        self.progress_callback("  - Setting Product...")
        product_dropdown_button = self.page.locator('//*[@id="accountForm"]/div/div[3]/div/div/input')
        expect(product_dropdown_button).to_be_visible(timeout=15000)
        product_dropdown_button.click()
        self.page.locator("span", has_text="Island Paradise Mobile (5007)").click()
        
        self.progress_callback("  - Filling Username...")
        username_locator = self.page.locator("#username")
        expect(username_locator).to_be_visible(timeout=15000)
        username_locator.fill(username)
        
        self.progress_callback("  - Filling Password...")
        password_locator = self.page.locator("#password")
        expect(password_locator).to_be_visible(timeout=15000)
        password_locator.fill(user_password)
        
        self.progress_callback("  - Setting Currency...")
        currency_dropdown_button = self.page.locator('//div[9]//div[1]//div[1]//input[1]')
        expect(currency_dropdown_button).to_be_visible(timeout=15000)
        currency_dropdown_button.click()
        # --- MODIFIED: Use the unique currency code to find the currency in the dropdown ---
        self.page.locator(f"span:has-text('({currency_code})')").click()
        
        self.progress_callback("  - Clicking 'Create Account'...")
        create_button_locator = self.page.locator("#submit")
        expect(create_button_locator).to_be_visible(timeout=30000)
        create_button_locator.click()

        try:
            start_time = time.time()
            while time.time() - start_time < 30:
                success_locator = self.page.locator(".card-panel.green")
                error_locator = self.page.locator(".card-panel.red")

                if success_locator.is_visible():
                    self.progress_callback(f"  - Successfully created user: {username}")
                    return True
                
                if error_locator.is_visible():
                    self.progress_callback("  - [ERROR] Creation failed. Checking logs...")
                    log_header = self.page.locator("div.collapsible-header:has-text('Log')")
                    expect(log_header).to_be_visible(timeout=10000)
                    log_header.click()
                    
                    log_message_locator = self.page.locator("#resultContainerMessage")
                    expect(log_message_locator).to_be_visible(timeout=10000)
                    error_details = log_message_locator.inner_text()
                    
                    self.progress_callback(f"  - Detailed Error: {error_details.strip()}")
                    return False
                
                self.page.wait_for_timeout(500)
            
            raise Exception("Timeout: Neither success nor error panel became visible after 30 seconds.")

        except Exception as e:
            self.progress_callback(f"  - [CRITICAL] Could not determine creation status. Error: {e}")
            return False

    def migrate_user_to_lvc(self, base_url, username):
        self.progress_callback(f"--- Starting LVC Migration for user: {username} ---")
        
        user_accounts_url = f"{base_url}/useraccounts"
        self.progress_callback(f"Navigating to User Accounts page: {user_accounts_url}")
        self.page.goto(user_accounts_url, timeout=60000)
        self.wait_for_page_to_settle()

        self.progress_callback("Clicking search icon to reveal search bar...")
        search_button_locator = self.page.locator("//button[@aria-label='Search']")
        expect(search_button_locator).to_be_visible(timeout=15000)
        search_button_locator.click()

        self.progress_callback(f"Searching for user: {username}")
        search_input_locator = self.page.locator("//input[@type='text']").first
        expect(search_input_locator).to_be_visible(timeout=15000)
        search_input_locator.fill(username)
        self.page.keyboard.press("Enter")
        
        self.progress_callback("Waiting for search results...")
        self.page.wait_for_timeout(1500)
        self.wait_for_page_to_settle()
        self.progress_callback("Search complete.")

        self.progress_callback("Locating user in table...")
        user_row_locator = self.page.locator(f"tr:has-text('{username}')")
        expect(user_row_locator).to_be_visible(timeout=15000)
        self.progress_callback("User row found.")

        self.progress_callback("Clicking 'three dots' menu...")
        three_dots_button = user_row_locator.locator("span.MuiIconButton-label")
        expect(three_dots_button).to_be_visible(timeout=10000)
        three_dots_button.click()
        
        self.progress_callback("Clicking 'Migrate LVCS' option...")
        migrate_option = self.page.get_by_role("menuitem", name="Migrate LVCS")
        expect(migrate_option).to_be_visible(timeout=10000)
        migrate_option.click()

        self.progress_callback("Confirming migration...")
        confirmation_button = self.page.get_by_role("button", name="Migrate")
        expect(confirmation_button).to_be_visible(timeout=10000)
        confirmation_button.click()

        self.progress_callback("Pausing for 1 second to allow migration to complete...")
        self.page.wait_for_timeout(1000)
        self.progress_callback(f"Successfully migrated {username} to LVC.")

    def add_balance_to_user(self, base_url, username, amount):
        self.progress_callback(f"--- Adding balance to user: {username} ---")
        
        balance_url = f"{base_url}/BalanceUserAccount"
        self.progress_callback(f"Navigating to Balance page: {balance_url}")
        self.page.goto(balance_url, timeout=60000)
        self.wait_for_page_to_settle()

        self.progress_callback(f"Filling balance form for {username}...")
        
        try:
            error_locator = self.page.locator("#loginName-error")
            if error_locator.is_visible(timeout=1000):
                self.progress_callback(f"  - [ERROR] User '{username}' not found on balance page. Skipping.")
                return
        except:
            pass

        username_locator = self.page.locator("#loginName")
        expect(username_locator).to_be_visible(timeout=15000)
        username_locator.fill(username)

        amount_locator = self.page.locator("#amount")
        expect(amount_locator).to_be_visible(timeout=15000)
        amount_locator.fill(amount)

        set_balance_button = self.page.locator("#submit")
        expect(set_balance_button).to_be_visible(timeout=15000)
        set_balance_button.click()

        try:
            start_time = time.time()
            while time.time() - start_time < 30:
                success_locator = self.page.locator(".card-panel.green")
                if success_locator.is_visible():
                    self.progress_callback(f"  - Successfully added balance to {username}")
                    return
                self.page.wait_for_timeout(500)
            raise Exception("Timeout: Could not find success message after setting balance.")
        except Exception as e:
            self.progress_callback(f"  - [CRITICAL] Could not determine balance status. Error: {e}")

    def close(self):
        if self.browser:
            self.browser.close()
            self.progress_callback("Browser closed.")
        if self.playwright:
            self.playwright.stop()

# =============================================================================
# Helper Functions
# =============================================================================
def parse_user_data(file_path, mode):
    """
    Parses user data from an Excel or JSON file based on the selected mode.
    """
    lvc_users, standard_users = [], []
    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension in ['.xlsx', '.xls']:
        # --- Excel Parsing Logic ---
        required_sheet = 'Sheet1'
        try:
            xls = pd.ExcelFile(file_path)
        except FileNotFoundError:
            raise FileNotFoundError(f"The file could not be found at the path: {file_path}")
        except Exception as e:
            raise IOError(f"The file at {file_path} could not be opened or is corrupted. Details: {e}")
        if required_sheet not in xls.sheet_names:
            raise ValueError(f"A required sheet named '{required_sheet}' was not found in the Excel file.")
        
        lvc_required_columns = ['lvc currency', 'username', 'lvc_username']
        standard_required_columns = ['std_currency', 'std_username']

        if mode in ["all", "lvc_only"]:
            lvc_df = pd.read_excel(file_path, sheet_name=required_sheet, header=0, usecols="A,B,D")
            lvc_df.columns = [str(col).strip().lower() for col in lvc_df.columns]
            if not all(col in lvc_df.columns for col in lvc_required_columns):
                raise ValueError(f"The LVC section is missing required headers: LVC Currency, Username, LVC_username.")
            lvc_df.rename(columns={'lvc currency': 'Currency', 'username': 'Username', 'lvc_username': 'LVC_username'}, inplace=True)
            lvc_df.dropna(subset=['Currency'], inplace=True)
            lvc_users = lvc_df.to_dict('records')

        if mode in ["all", "standard_only"]:
            standard_df = pd.read_excel(file_path, sheet_name=required_sheet, header=0, usecols="E,F")
            standard_df.columns = [str(col).strip().lower() for col in standard_df.columns]
            if not all(col in standard_df.columns for col in standard_required_columns):
                raise ValueError(f"The Standard section is missing required headers: std_currency, std_username.")
            standard_df.rename(columns={'std_currency': 'Currency', 'std_username': 'Username'}, inplace=True)
            standard_df.dropna(subset=['Currency'], inplace=True)
            standard_users = standard_df.to_dict('records')
            
    elif file_extension == '.json':
        # --- JSON Parsing Logic (MODIFIED for flexibility) ---
        with open(file_path, 'r') as f:
            data = {k.lower(): v for k, v in json.load(f).items()} # Make top-level keys case-insensitive
        
        # Determine the correct keys for user lists
        lvc_list_key = 'lvc_users' if 'lvc_users' in data else 'lvc_currencies'
        std_list_key = 'standard_users' if 'standard_users' in data else 'std_currencies'

        if mode in ["all", "lvc_only"] and lvc_list_key in data:
            # Normalize keys within each user dictionary to be case-insensitive
            normalized_lvc_users = []
            for user in data[lvc_list_key]:
                normalized_user = {k.lower().replace('_', ' '): v for k, v in user.items()}
                # Standardize to the keys the application expects
                final_user = {
                    'Currency': normalized_user.get('lvc currency', ''),
                    'Username': normalized_user.get('username', ''),
                    'LVC_username': normalized_user.get('lvc username', '')
                }
                normalized_lvc_users.append(final_user)
            lvc_users = normalized_lvc_users
        
        if mode in ["all", "standard_only"] and std_list_key in data:
            normalized_std_users = []
            for user in data[std_list_key]:
                normalized_user = {k.lower().replace('_', ' '): v for k, v in user.items()}
                final_user = {
                    'Currency': normalized_user.get('std currency', ''),
                    'Username': normalized_user.get('std username', '')
                }
                normalized_std_users.append(final_user)
            standard_users = normalized_std_users
    else:
        raise ValueError(f"Unsupported file type: {file_extension}. Please select an Excel or JSON file.")
        
    return lvc_users, standard_users

def parse_credentials_file(file_path):
    with open(file_path, 'r') as f:
        lines = f.readlines()
    if len(lines) < 2:
        raise ValueError("Credentials file must have at least two lines (email and password).")
    email = lines[0].strip()
    password = lines[1].strip()
    return email, password

def parse_gtp_list_file(file_path):
    """Reads and validates the GTP list from a JSON file."""
    with open(file_path, 'r') as f:
        data = json.load(f)
    if not isinstance(data, dict) or not data:
        raise ValueError("GTP list file must contain a non-empty JSON object.")
    return data

def update_excel_with_lvc_names(file_path, migrated_user, postfix):
    """
    Updates the LVC_username column in the original Excel file using openpyxl.
    """
    try:
        workbook = load_workbook(filename=file_path)
        sheet = workbook['Sheet1']
        
        original_username = migrated_user['Username']
        original_currency = migrated_user['Currency']
        lvc_username = f"LVC_{original_username}{postfix}"

        # Find the row to update by matching currency and original username
        for row in range(2, sheet.max_row + 1): # Start from row 2 to skip header
            if sheet.cell(row=row, column=1).value == original_currency and sheet.cell(row=row, column=2).value == original_username:
                # Update the LVC_username column (column D)
                sheet.cell(row=row, column=4).value = lvc_username
                break
        
        workbook.save(filename=file_path)
        print(f"Successfully updated LVC_username for {original_username} in {file_path}")

    except Exception as e:
        print(f"Error updating Excel file: {e}")

def update_json_with_lvc_names(file_path, migrated_user, postfix):
    """
    Updates the LVC_username key in the original JSON file (case-insensitively).
    """
    try:
        with open(file_path, 'r') as f:
            data = json.load(f)
        
        original_username = migrated_user['Username']
        original_currency = migrated_user['Currency']
        lvc_username = f"LVC_{original_username}{postfix}"

        # Find the correct list key (lvc_users or lvc_currencies)
        lvc_list_key = None
        for key in data.keys():
            if key.lower() in ['lvc_users', 'lvc_currencies']:
                lvc_list_key = key
                break
        
        if lvc_list_key:
            for user in data[lvc_list_key]:
                # Find user by matching currency and username case-insensitively
                user_keys_lower = {k.lower().replace('_', ' '): k for k in user.keys()}
                
                currency_key = user_keys_lower.get('lvc currency')
                username_key = user_keys_lower.get('username')
                lvc_username_key = user_keys_lower.get('lvc username')

                if currency_key and username_key and lvc_username_key:
                    if user[currency_key] == original_currency and user[username_key] == original_username:
                        user[lvc_username_key] = lvc_username
                        break
        
        with open(file_path, 'w') as f:
            json.dump(data, f, indent=2)
        print(f"Successfully updated LVC_username for {original_username} in {file_path}")

    except Exception as e:
        print(f"Error updating JSON file: {e}")

# =============================================================================
# Automation Worker (Controller)
# =============================================================================
class AutomationWorker(QObject):
    progress_update = pyqtSignal(str)
    automation_error = pyqtSignal(str)
    automation_finished = pyqtSignal()

    def __init__(self, gtp_url, email, password, lvc_users, standard_users, user_password, user_data_path, mode, lvc_balance, standard_balance, postfix):
        super().__init__()
        self.gtp_url = gtp_url
        self.email = email
        self.password = password
        self.lvc_users = lvc_users
        self.standard_users = standard_users
        self.user_password = user_password
        self.user_data_path = user_data_path
        self.mode = mode
        self.lvc_balance = lvc_balance
        self.standard_balance = standard_balance
        self.postfix = postfix
        self.is_running = True
        self.driver = None

    def run(self):
        self.driver = BrowserDriver(progress_callback=self.progress_update.emit)
        try:
            self.progress_update.emit("Automation thread started.")
            if not self.is_running: return

            self.driver.launch()
            if not self.is_running: return
            
            self.driver.login(self.gtp_url, self.email, self.password)
            if not self.is_running: return
            
            # --- LVC User Processing ---
            if self.mode in ["all", "lvc_only"]:
                self.progress_update.emit("\n--- STEP A: Creating LVC User Accounts ---")
                created_lvc_users = []
                if self.lvc_users:
                    self.driver.navigate_to_create_user_page(self.gtp_url)
                    for user in self.lvc_users:
                        if not self.is_running: break
                        success = self.driver.fill_user_creation_form(user, self.user_password, self.postfix)
                        if success:
                            created_lvc_users.append(user)
                if not self.is_running: return

                self.progress_update.emit("\n--- STEP B: Migrating Users to LVC ---")
                for user in created_lvc_users:
                    if not self.is_running: break
                    initial_username = f"{user['Username']}{self.postfix}"
                    self.driver.migrate_user_to_lvc(self.gtp_url, initial_username)
                    if self.user_data_path.lower().endswith(('.xlsx', '.xls')):
                        update_excel_with_lvc_names(self.user_data_path, user, self.postfix)
                    elif self.user_data_path.lower().endswith('.json'):
                        update_json_with_lvc_names(self.user_data_path, user, self.postfix)

                if not self.is_running: return

                self.progress_update.emit("\n--- STEP C: Adding Balance to LVC Users ---")
                for user in created_lvc_users:
                    if not self.is_running: break
                    lvc_username = f"LVC_{user['Username']}{self.postfix}"
                    self.driver.add_balance_to_user(self.gtp_url, lvc_username, self.lvc_balance)
                if not self.is_running: return

            # --- Standard User Processing ---
            if self.mode in ["all", "standard_only"]:
                self.progress_update.emit("\n--- Creating Standard User Accounts ---")
                created_standard_users = []
                if self.standard_users:
                    self.driver.navigate_to_create_user_page(self.gtp_url)
                    for user in self.standard_users:
                        if not self.is_running: break
                        success = self.driver.fill_user_creation_form(user, self.user_password, self.postfix)
                        if success:
                            created_standard_users.append(user)
                if not self.is_running: return

                self.progress_update.emit("\n--- Adding Balance to Standard Users ---")
                for user in created_standard_users:
                    if not self.is_running: break
                    standard_username = f"{user['Username']}{self.postfix}"
                    self.driver.add_balance_to_user(self.gtp_url, standard_username, self.standard_balance)

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
    def __init__(self):
        super().__init__()
        self.automation_thread = None
        self.worker = None
        self.config = {}
        self.GTP_VERSIONS = {}
        self.setWindowTitle("GTP User Automation Tool v1.1")
        self.setGeometry(100, 100, 700, 680) 
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
        self.check_start_button_state()

    def setup_ui_elements(self, header_font):
        config_header = QLabel("1. GTP Configuration")
        config_header.setFont(header_font)
        self.main_layout.addWidget(config_header)

        gtp_layout = QHBoxLayout()
        self.select_gtp_list_button = QPushButton("Select GTP List File (.json)")
        self.gtp_list_path_label = QLabel("No file selected.")
        self.gtp_list_path_label.setStyleSheet("font-style: italic; color: #555;")
        gtp_layout.addWidget(self.select_gtp_list_button)
        gtp_layout.addWidget(self.gtp_list_path_label, 1)
        self.main_layout.addLayout(gtp_layout)

        self.gtp_dropdown = QComboBox()
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
        self.select_user_file_button = QPushButton("Select User Data File (excel or .json)")
        self.user_file_path_label = QLabel("No file selected.")
        self.user_file_path_label.setStyleSheet("font-style: italic; color: #555;")
        user_layout.addWidget(self.select_user_file_button)
        user_layout.addWidget(self.user_file_path_label, 1)
        self.main_layout.addLayout(user_layout)
        
        user_pass_label = QLabel("Password for New User Accounts:")
        self.main_layout.addWidget(user_pass_label)
        self.user_password_input = QLineEdit()
        self.user_password_input.setText("snow")
        self.main_layout.addWidget(self.user_password_input)

        postfix_label = QLabel("Postfix for Usernames:")
        self.main_layout.addWidget(postfix_label)
        self.postfix_input = QLineEdit()
        self.postfix_input.setText("x1")
        self.main_layout.addWidget(self.postfix_input)


        balance_groupbox = QGroupBox("Balance Amounts")
        balance_layout = QVBoxLayout()
        
        lvc_balance_layout = QHBoxLayout()
        lvc_balance_label = QLabel("Balance for LVC user:")
        self.lvc_balance_input = QLineEdit()
        self.lvc_balance_input.setText("7000000000000")
        self.lvc_balance_input.setMaxLength(13)
        lvc_balance_layout.addWidget(lvc_balance_label)
        lvc_balance_layout.addWidget(self.lvc_balance_input)
        balance_layout.addLayout(lvc_balance_layout)

        standard_balance_layout = QHBoxLayout()
        standard_balance_label = QLabel("Balance for Standard user:")
        self.standard_balance_input = QLineEdit()
        self.standard_balance_input.setText("9999999")
        self.standard_balance_input.setMaxLength(7)
        self.standard_balance_input.setValidator(QIntValidator(0, 9999999))
        standard_balance_layout.addWidget(standard_balance_label)
        standard_balance_layout.addWidget(self.standard_balance_input)
        balance_layout.addLayout(standard_balance_layout)
        
        balance_groupbox.setLayout(balance_layout)
        self.main_layout.addWidget(balance_groupbox)

        mode_groupbox = QGroupBox("Processing Mode")
        mode_layout = QHBoxLayout()
        self.radio_all = QRadioButton("LVC + Standard Currency")
        self.radio_lvc = QRadioButton("Only LVC")
        self.radio_standard = QRadioButton("Only Standard Currency")
        self.radio_all.setChecked(True)
        mode_layout.addWidget(self.radio_all)
        mode_layout.addWidget(self.radio_lvc)
        mode_layout.addWidget(self.radio_standard)
        mode_groupbox.setLayout(mode_layout)
        self.main_layout.addWidget(mode_groupbox)
        
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
        self.select_gtp_list_button.clicked.connect(self.select_gtp_list_file)
        self.select_cred_button.clicked.connect(self.select_credentials_file)
        self.select_user_file_button.clicked.connect(self.select_user_data_file)
        self.start_button.clicked.connect(self.start_automation)

    def start_automation(self):
        gtp_selection = self.gtp_dropdown.currentText()
        gtp_url = self.GTP_VERSIONS[gtp_selection]
        cred_path = self.cred_path_label.text()
        user_data_path = self.user_file_path_label.text()
        user_password = self.user_password_input.text()
        lvc_balance = self.lvc_balance_input.text()
        standard_balance = self.standard_balance_input.text()
        postfix = self.postfix_input.text()
        
        mode = "all"
        if self.radio_lvc.isChecked():
            mode = "lvc_only"
        elif self.radio_standard.isChecked():
            mode = "standard_only"

        errors = []
        if "No file selected" in cred_path: errors.append("You must select a credentials file.")
        if "No file selected" in user_data_path: errors.append("You must select a user data file.")
        if not user_password: errors.append("The password for new users cannot be empty.")
        if not postfix: errors.append("The postfix cannot be empty.")
        if not lvc_balance.isdigit() or not standard_balance.isdigit():
            errors.append("Balance amounts must be valid numbers.")
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
            lvc_users, standard_users = parse_user_data(user_data_path, mode)
            self.log_message(f"  - Validation successful: Found {len(lvc_users)} LVC and {len(standard_users)} Standard users for mode '{mode}'.")
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
        self.worker = AutomationWorker(gtp_url, email, password, lvc_users, standard_users, user_password, user_data_path, mode, lvc_balance, standard_balance, postfix)
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

    def select_gtp_list_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select GTP List File", "", "JSON Files (*.json)")
        if file_path:
            self.load_gtp_list_from_path(file_path)

    def select_credentials_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Credentials File", "", "Text Files (*.txt)")
        if file_path:
            self.cred_path_label.setText(file_path)
            self.cred_path_label.setStyleSheet("font-style: normal; color: #000;")
            self.log_message(f"Selected credentials file: {file_path}")
            self.config['credentials_path'] = file_path
            self.save_config()
            self.check_start_button_state()

    def select_user_data_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select User Data File", "", "Data Files (*.xlsx *.xls *.json)")
        if file_path:
            self.user_file_path_label.setText(file_path)
            self.user_file_path_label.setStyleSheet("font-style: normal; color: #000;")
            self.log_message(f"Selected user data file: {file_path}")
            self.config['user_data_path'] = file_path
            self.save_config()
            self.check_start_button_state()
            
    def load_gtp_list_from_path(self, file_path):
        try:
            self.GTP_VERSIONS = parse_gtp_list_file(file_path)
            self.gtp_dropdown.clear()
            self.gtp_dropdown.addItems(self.GTP_VERSIONS.keys())
            self.gtp_list_path_label.setText(file_path)
            self.gtp_list_path_label.setStyleSheet("font-style: normal; color: #000;")
            self.log_message(f"Successfully loaded {len(self.GTP_VERSIONS)} GTP versions.")
            self.config['gtp_list_path'] = file_path
            self.save_config()
        except Exception as e:
            self.GTP_VERSIONS = {}
            self.gtp_dropdown.clear()
            QMessageBox.critical(self, "File Error", f"Failed to load GTP list file.\n\nError: {e}")
        self.check_start_button_state()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                self.config = json.load(f)
        else:
            self.config = {}

    def save_config(self):
        self.config['postfix'] = self.postfix_input.text()
        self.config['last_gtp_selection'] = self.gtp_dropdown.currentText()
        with open(CONFIG_FILE, 'w') as f:
            json.dump(self.config, f, indent=4)

    def apply_config(self):
        """Applies loaded configuration to the UI."""
        gtp_path = self.config.get('gtp_list_path')
        if gtp_path and os.path.exists(gtp_path):
            self.load_gtp_list_from_path(gtp_path)
            # Restore last selection
            last_selection = self.config.get('last_gtp_selection')
            if last_selection:
                index = self.gtp_dropdown.findText(last_selection)
                if index != -1:
                    self.gtp_dropdown.setCurrentIndex(index)


        cred_path = self.config.get('credentials_path')
        if cred_path and os.path.exists(cred_path):
            self.cred_path_label.setText(cred_path)
            self.cred_path_label.setStyleSheet("font-style: normal; color: #000;")
        
        user_data_path = self.config.get('user_data_path')
        if user_data_path and os.path.exists(user_data_path):
            self.user_file_path_label.setText(user_data_path)
            self.user_file_path_label.setStyleSheet("font-style: normal; color: #000;")

        postfix = self.config.get('postfix', 'x1')
        self.postfix_input.setText(postfix)


    def check_start_button_state(self):
        """Enables the start button only if all required files are selected."""
        gtp_loaded = bool(self.GTP_VERSIONS)
        creds_loaded = "No file selected" not in self.cred_path_label.text()
        users_loaded = "No file selected" not in self.user_file_path_label.text()
        
        if gtp_loaded and creds_loaded and users_loaded:
            self.start_button.setEnabled(True)
        else:
            self.start_button.setEnabled(False)

    def log_message(self, message):
        self.status_log.append(message)
        logging.info(message) # Also log to file

    def toggle_controls(self, enabled):
        self.gtp_dropdown.setEnabled(enabled)
        self.select_gtp_list_button.setEnabled(enabled)
        self.select_cred_button.setEnabled(enabled)
        self.select_user_file_button.setEnabled(enabled)
        self.user_password_input.setEnabled(enabled)
        self.postfix_input.setEnabled(enabled)
        self.lvc_balance_input.setEnabled(enabled)
        self.standard_balance_input.setEnabled(enabled)
        self.start_button.setEnabled(enabled)
        self.radio_all.setEnabled(enabled)
        self.radio_lvc.setEnabled(enabled)
        self.radio_standard.setEnabled(enabled)
        
    def closeEvent(self, event):
        self.save_config()
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

