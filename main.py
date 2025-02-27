import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException

from webdriver_manager.chrome import ChromeDriverManager

from google.oauth2.credentials import Credentials as UserCredentials
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from googleapiclient.discovery import build

from googleapiclient.errors import HttpError

import tkinter as tk
from tkinter import ttk, messagebox

import datetime
import time

import re
import sys

import random
import os
import json
import platform
import subprocess

import requests
from requests.exceptions import RequestException

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

SPREADSHEET_ID = '13AK7mmpuRHUc7yyP9l7VwEJKPLJ26MPja5o05Ib-J88'
DOMAIN_SHEET_ID = '1DCUQw7c92AEKk0pURXegYp_Vy2TklDVijaoz0OMnTpw'
INTERN_SHEET_ID = 'spreadsheets/d/1e51xKqxg55XcFCC0w9ySKTU4QvNZXTT6-OpxWV_HX6U'

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.compose',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.settings.basic'
]

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36"
]

def check_bot_detection(driver):
    print("Checking for bot detection...")
    try:
        if "security_check" in driver.current_url:
            print("Bot detection: Security check in URL")
            return True
        
        driver.find_element(By.ID, 'captcha-internal')
        print("Bot detection: CAPTCHA found")
        return True
    except NoSuchElementException:
        return False


# Human-like typing simulation
def human_type(element, text):
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(0.02, 0.2))


#------------------------------------------------------------------------------------------------------------------------------------------------------#

# Hunter API
class HunterIOAPI:
    def __init__(self, api_key):
        self.base_url = "https://api.hunter.io/v2"
        self.api_key = api_key
        
    def _make_request(self, endpoint, params):
        try:
            params = {k: v for k, v in params.items() if v is not None}
            params['api_key'] = self.api_key

            response = requests.get(
                f"{self.base_url}/{endpoint}",
                params=params,
                timeout=30
            )
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            print(f"Hunter API Error: {str(e)}")
            # Use e.response to check the status code, if available
            if e.response is not None and e.response.status_code == 429:
                print("Rate limit exceeded. Waiting 60 seconds...")
                time.sleep(60)
                return self._make_request(endpoint, params)
            return None

    def find_email(self, first_name, last_name, company=None, domain=None, full_name=None, max_duration=10):
        """Use Email Finder endpoint"""
        params = {
            'first_name': first_name,
            'last_name': last_name,
            'full_name': full_name if full_name else f"{first_name} {last_name}",
            'company': company,
            'domain': domain,
            'max_duration': max_duration
        }
        data = self._make_request('email-finder', params)
        
        if data and data.get('data'):
            return {
                'email': data['data']['email'],
                'confidence': data['data']['score'],
                'sources': [s['uri'] for s in data['data'].get('sources', [])]
            }
        return None

    def verify_email(self, email):
        """Use Email Verifier endpoint"""
        params = {'email': email}
        data = self._make_request('email-verifier', params)
        
        if data and data.get('data'):
            return data['data']['status']
        return None

    def domain_search(self, domain, company=None, limit=10):
        """Use Domain Search endpoint"""
        params = {
            'domain': domain,
            'company': company,
            'limit': limit
        }
        data = self._make_request('domain-search', params)
        
        if data and data.get('data'):
            return [email['value'] for email in data['data']['emails']]
        return []


#------------------------------------------------------------------------------------------------------------------------------------------------------#

# GUI Implementation
class LinkedInApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("1200x800")
        self.driver = None
        self._processing = False
        self.current_url_index = 0
        self.hunter_authenticated = False
        self.extraction_window = None
        
        # Initialize services with separate credentials
        try:
            # Sheets authentication using service account
            sheets_creds = ServiceAccountCredentials.from_service_account_file(
                'credentials.json', 
                scopes=SCOPES
            )
            self.sheet_service = build('sheets', 'v4', credentials=sheets_creds)
            
            # Gmail authentication using OAuth token
            gmail_creds = None
            if os.path.exists('token.json'):
                gmail_creds = UserCredentials.from_authorized_user_file('token.json', SCOPES)
            
            if not gmail_creds or not gmail_creds.valid:
                if gmail_creds and gmail_creds.expired and gmail_creds.refresh_token:
                    gmail_creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
                    gmail_creds = flow.run_local_server(port=0)
                with open('token.json', 'w') as token:
                    token.write(gmail_creds.to_json())
                    
            self.gmail_service = build('gmail', 'v1', credentials=gmail_creds)
            
        except Exception as e:
            messagebox.showerror("API Error", f"Failed to initialize services: {str(e)}")
            self.root.destroy()

        self.create_gui()
        self.create_review_treeview()

    def on_close(self):
        """Ensure proper cleanup when closing"""
        if self._processing:
            if not messagebox.askyesno("Confirm Exit", 
                "Processing is still running. Are you sure you want to exit?"):
                return
        
        if self.driver:
            try:
                self.driver.quit()
            except Exception as e:
                print(f"Error closing driver: {e}")
        
        # Save any pending data
        try:
            self.save_to_sheets()
        except Exception as e:
            print(f"Error saving data: {e}")
        
        self.root.destroy()

    def create_gui(self):
        # Main Window
        self.root.title("LinkedIn Manager")
        # self.root.attributes('-topmost', True)
        
        # Input Frame
        self.input_frame = ttk.Frame(self.root, padding=20)
        self.input_frame.pack()
        
        # URL Input
        ttk.Label(self.input_frame, text="Enter LinkedIn Profile URLs:").pack()
        self.url_text = tk.Text(self.input_frame, height=10, width=60)
        self.url_text.pack()
        
        # Control Buttons
        ttk.Button(self.input_frame, text="Process Profiles", 
                 command=self.start_processing).pack(pady=10)
        
        # Status Bar
        self.status = ttk.Label(self.root, text="Ready", relief=tk.SUNKEN)
        self.status.pack(fill=tk.X)

        ttk.Button(self.input_frame, text="Send Scheduled Emails",
             command=self.prepare_emails).pack(pady=10)

    def create_review_treeview(self):
        # Treeview Frame
        self.review_frame = ttk.Frame(self.root)
        self.review_frame.pack(fill='both', expand=True)
        
        # Treeview with Scrollbars - SPECIFY PARENT FRAME
        columns = ('Name', 'Email', 'Company', 'Position', 'Education', 'College', 'Highschool', 'Thai', 'URL', 'Intern')
        self.tree = ttk.Treeview(self.review_frame, columns=columns, show="headings")
        
        # Configure columns first
        for col in columns:
            self.tree.heading(col, text=col)
            if col in ['Name', 'Email', 'Company', 'Thai', 'Intern']:
                self.tree.column(col, width=120)
            else:
                self.tree.column(col, width=0, stretch=tk.NO)

        # Create scrollbars WITHIN THE REVIEW FRAME
        vsb = ttk.Scrollbar(self.review_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.review_frame, orient="horizontal", command=self.tree.xview)
        
        # Link scrollbars to treeview
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Grid layout WITHIN REVIEW FRAME
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        # Configure column headings
        self.tree.heading('Name', text='Name')
        self.tree.heading('Email', text='Email')
        self.tree.heading('Company', text='Company')
        self.tree.heading('Position', text='Position')
        self.tree.heading('Education', text='Education')
        self.tree.heading('College', text='College')
        self.tree.heading('Highschool', text='Highschool')
        self.tree.heading('Thai', text='Thai')
        self.tree.heading('URL', text='URL')
        
        # Configure grid weights for proper resizing
        self.review_frame.grid_rowconfigure(0, weight=1)
        self.review_frame.grid_columnconfigure(0, weight=1)
        
        # Edit binding
        self.tree.bind('<Double-1>', self.on_cell_edit)

    def on_cell_edit(self, event):
        # Cell editing implementation
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            col = self.tree.identify_column(event.x)
            col_index = int(col[1:]) - 1
            item = self.tree.identify_row(event.y)
            values = list(self.tree.item(item, 'values'))
            # If the "Thai" column (index 7) is clicked, toggle its value.
            if col_index == 7:
                new_value = not (values[7] in [True, 'True'])
                values[7] = new_value
                self.tree.item(item, values=values)
            else:
                # For other columns, allow normal editing.
                current_value = values[col_index]
                edit_win = tk.Toplevel()
                edit_win.title("Edit Value")
                entry = ttk.Entry(edit_win)
                entry.insert(0, current_value)
                entry.pack(padx=10, pady=10)
                ttk.Button(edit_win, text="Save",
                           command=lambda: self.save_edited_value(item, col_index, entry.get(), edit_win)
                           ).pack()
            
    def handle_bot_detection(self):
        self.driver.save_screenshot('bot_detection.png')
        self._processing = False  # Pause processing
        
        verify_win = tk.Toplevel(self.root)
        verify_win.grab_set()  # Make window modal
        verify_win.title("Verification Required")
        
        ttk.Label(verify_win, 
                text="LinkedIn requires manual verification\n"
                    "1. Complete the security check in the browser\n"
                    "2. Click Continue when done").pack(padx=20, pady=10)
        
        ttk.Button(verify_win, text="Continue", 
                command=lambda: self.resume_processing(verify_win)).pack(pady=10)

    def resume_processing(self, window):
        window.destroy()
        try:
            self.driver.refresh()
            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.TAG_NAME, 'body'))
            )
        except:
            pass
        self._processing = True
        self.process_next_profile()
    
    def start_processing(self):
        if self._processing:
            messagebox.showinfo("Already running", "Processing is already in progress")
            return

        existing_urls = {self.tree.item(item)['values'][8] for item in self.tree.get_children()}
        raw_urls = [url.strip() for url in self.url_text.get("1.0", tk.END).splitlines() 
                    if url.startswith('https://www.linkedin.com/in/')]
        self.profile_urls = list(dict.fromkeys([u for u in raw_urls if u not in existing_urls]))
        
        if not self.profile_urls:
            messagebox.showwarning("No URLs", "No valid LinkedIn URLs found")
            return
        
        # Clear and disable URL input during extraction.
        print("Beginning Extraction")
        self.url_text.delete("1.0", tk.END)
        self.url_text.config(state=tk.DISABLED)

        # Initialize driver if needed
        if not self.driver:
            chrome_options = Options()
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_argument(f"--user-agent={random.choice(USER_AGENTS)}")
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument("--disable-dev-shm-usage")

            try:
                # Get Chrome version
                chrome_version = subprocess.check_output(
                    '/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --version',
                    shell=True
                ).decode().split()[-1]  # Corrected to [-1]

                # Get matching ChromeDriver version
                service = Service(ChromeDriverManager(driver_version=chrome_version).install())

                # Apple Silicon specific config
                if sys.platform == 'darwin' and platform.machine() == 'arm64':
                    chrome_options.binary_location = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
                    
                    # Add ARM-specific flags
                    chrome_options.add_argument("--use-angle=vulkan")
                    chrome_options.add_argument("--disable-features=UseChromeOSDirectVideoDecoder")

                self.driver = webdriver.Chrome(
                    service=service,
                    options=chrome_options
                )
                self.linkedin_login()
            except Exception as e:
                messagebox.showerror("Driver Error", 
                    f"Critical browser initialization error: {str(e)}\n\n"
                    "Troubleshooting steps:\n"
                    "1. Run in terminal: pip install --upgrade webdriver-manager\n"
                    "2. Delete cached drivers: rm -rf ~/.wdm/drivers\n"
                    "3. Verify Chrome path: /Applications/Google Chrome.app")
                return
        
        self.current_url_index = 0
        self._processing = True
        self.process_next_profile()

        # Clear the text area and re-enable input
        self.url_text.config(state=tk.NORMAL)
        print("Processing complete - ready for new URLs")


    def process_next_profile(self):
        if not self._processing or self.current_url_index >= len(self.profile_urls):
            self._processing = False
            self.update_ui_status("Processing complete")
            self.bring_to_front()
            print("Processing complete - ready for new URLs")
            return

        url = self.profile_urls[self.current_url_index]
        try:
            print(f"\n=== Processing profile {self.current_url_index+1}/{len(self.profile_urls)} ===")
            print(f"Navigating to: {url}")
            
            if check_bot_detection(self.driver):
                self.handle_bot_detection()
                return

            # Process profile with error wrapping
            try:
                data = get_linkedin_profile_info(self, self.driver, url)
                if data is None:  # Email lookup failed
                    print(f"Skipping profile with no email: {url}")
                    self.current_url_index += 1
                    self.root.after(1000, self.process_next_profile)
                    return
                
                self.add_to_treeview(data, url)
                self.root.after(int(random.uniform(15000, 30000)), self.process_next_profile)
            except Exception as e:
                print(f"Critical error processing profile: {str(e)}")
                raise

            self.current_url_index += 1
            self.root.after(int(random.uniform(15000, 30000)), self.process_next_profile)

        except Exception as e:
            print(f"Error processing profile: {str(e)}")
            # Restart driver if connection lost
            try:
                self.driver.current_url  # Test driver responsiveness
            except:
                print("Driver crashed! Reinitializing...")
                self.driver.quit()
                self.driver = None
                self.start_processing()
                return
            
            if "404" in str(e):
                self.current_url_index += 1
                
            if self.current_url_index >= len(self.profile_urls):
                self._processing = False
                self.update_ui_status("Processing complete")
                self.bring_to_front()
                print("Processing complete - ready for new URLs")
                return

            self.handle_scrape_error(url, str(e))
            self.root.after(45000, self.process_next_profile)

    def update_ui_status(self, text, progress=None):
        """Update status bar and progress"""
        self.root.after(0, lambda: self.status.config(text=text))
        self.status.config(text=text)
        # Immediately update the GUI
        self.root.update_idletasks()

    def add_to_treeview(self, data, url):
        """Direct treeview update"""
        self.tree.insert('', 'end', values=(
            data['name'],
            data['email'],
            data['company'],
            data['position'],
            data['education'],
            data['college'],
            data['highschool'],
            data['thai'],
            url,
            data['intern']
        ))
        # Immediately update the GUI
        self.root.update_idletasks()

    def handle_scrape_error(self, url, error_msg):
        """Handle scraping errors"""
        messagebox.showerror("Error", f"Failed to process {url}: {error_msg}")

    def linkedin_login(self):
        try:
            with open("config.json") as f:
                config = json.load(f)
                
            print("Navigating to LinkedIn login...")
            self.driver.get("https://www.linkedin.com/login")
                    
            # Email field
            email_field = WebDriverWait(self.driver, 25).until(
                EC.presence_of_element_located((By.ID, "username"))
            )
            print("Entering email...")
            human_type(email_field, config['email'])
            time.sleep(random.uniform(0.5, 1.5))
            
            # Password field
            password_field = self.driver.find_element(By.ID, "password")
            print("Entering password...")
            human_type(password_field, config['password'])
            time.sleep(random.uniform(0.5, 1.5))
            
            # Click login button
            print("Submitting login...")
            self.driver.find_element(By.XPATH, "//button[@type='submit']").click()
            
            # Handle recovery email check
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "input__email_verification_pin"))
                )
                messagebox.showwarning("Recovery Check Needed", "...")
            except TimeoutException:
                pass
                
            # Verify login by waiting for profile dropdown
            print("Verifying login success...")
            WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.global-nav__me"))
            )
            print(f"Login successful! Current URL: {self.driver.current_url}")
            time.sleep(2)

        except Exception as e:
            print(f"Login failed with error: {str(e)}")
            self.driver.save_screenshot('login_error.png')
            messagebox.showerror("Login Error", "...")
            self.on_close()

    def save_edited_value(self, item, col_index, new_value, window):
        values = list(self.tree.item(item, 'values'))
        values[col_index] = new_value
        self.tree.item(item, values=values)
        window.destroy()

    def get_links_from_sheet(sheet_service, spreadsheet_id, range):
        try:
            result = sheet_service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id, range=range
            ).execute()
            return [row[0] for row in result.get('values', []) if row]
        except Exception as e:
            print(f"Error fetching links: {str(e)}")
            return []

    def save_to_sheets(self):
        # Fetch existing links from both sheets
        print("Saving to sheets...")
        existing_main = self.get_links_from_sheet(self.sheet_service, SPREADSHEET_ID, 'Sheet1!K:K')
        existing_intern = self.get_links_from_sheet(self.sheet_service, INTERN_SHEET_ID, 'Sheet1!K:K')
        all_existing = set(existing_main + existing_intern)

        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            link = values[8]
            if link in all_existing:
                continue

            is_intern = values[9]
            spreadsheet_id = INTERN_SHEET_ID if is_intern else SPREADSHEET_ID

            # Append to the appropriate sheet
            append_row_to_sheet(
                self.sheet_service, spreadsheet_id,
                link=link,
                name=values[0],
                company=values[2],
                position=values[3],
                email=values[1],
                education=values[4],
                same_college=values[5],
                same_highschool=values[6],
                thai=values[7]
            )

    def prepare_emails(self):
        # Disable GUI elements during email processing
        self.url_text.config(state=tk.DISABLED)
        
        # Get emails from treeview instead of sheets
        emailSendList = []
        
        # First, get all existing LinkedIn URLs from the sheet
        try:
            result = self.sheet_service.spreadsheets().values().get(
                spreadsheetId=SPREADSHEET_ID,
                range='Sheet1!M:M'  # LinkedIn URLs column
            ).execute()
            existing_urls = set(row[0] for row in result.get('values', [])[1:] if row)  # Skip header
        except Exception as e:
            messagebox.showerror("Sheet Error", f"Failed to check existing records: {str(e)}")
            self.url_text.config(state=tk.NORMAL)
            return
        
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            linkedin_url = values[8]  # URL from treeview
            
            # Skip if URL already exists in sheet
            if linkedin_url in existing_urls:
                print(f"Skipping duplicate LinkedIn profile: {linkedin_url}")
                continue
            
            # Create a new row in the sheet for tracking
            try:
                current_date = datetime.datetime.now().strftime('%Y-%m-%d')
                
                # Prepare row values according to specified columns
                row_values = [
                    'N/A',                  # A: Offered help on
                    'Not Yet',              # B: Status
                    values[0],              # C: Name
                    values[2],              # D: Company
                    '',                     # E: Location (empty)
                    '',                     # F: Industry (empty)
                    values[4],              # G: Education
                    values[3],              # H: Position
                    'Yes' if values[7] else 'No',  # I: Connection (Thai)
                    'Yes' if values[5] else 'No',  # J: Connection (College)
                    'Yes' if values[6] else 'No',  # K: Connection (Highschool)
                    values[1],              # L: Gmail
                    linkedin_url,           # M: LinkedIn URL
                    'Yes',                  # N: Coffee Chat Invite Sent?
                    'No',                   # O: Coffee Chat Yet?
                    'No',                   # P: Thank You Message Sent?
                    'No',                   # Q: Connect on LinkedIn?
                    'No',                   # R: Follow Up 1
                    '',                     # S: Coffee Chat Notes
                    'Process',              # T: Process with Bot?
                    current_date            # U: Last Updated Date
                ]
                
                result = self.sheet_service.spreadsheets().values().append(
                    spreadsheetId=SPREADSHEET_ID,
                    range='Sheet1!A2',
                    valueInputOption='USER_ENTERED',
                    body={'values': [row_values]}
                ).execute()
                
                # Get the row number of the newly inserted row
                row_number = len(result.get('updates', {}).get('updatedRange', '').split('!')[1].split(':')[0])
                
                filtered_row = {
                    'Name': values[0],
                    'Company': values[2],
                    'Position': values[3],
                    'Email': values[1],
                    'Education': values[4],
                    'College': values[5],
                    'Highschool': values[6],
                    'Thai': values[7],
                    'row_number': row_number
                }
                emailSendList.append(filtered_row)
                
            except Exception as e:
                messagebox.showerror("Sheet Error", f"Failed to add record: {str(e)}")
                self.url_text.config(state=tk.NORMAL)
                return
        
        if emailSendList:
            print("\n=== Sending Emails ===")
            self.send_email(emailSendList)
        else:
            messagebox.showinfo("No Emails", "No new emails to send")
            self.url_text.config(state=tk.NORMAL)

    def send_email(self, emailSendList):
        try:
            # Validate emails first
            invalid_emails = []
            valid_entries = []
            
            # Verify resume exists
            resume_path = os.path.join('Resources', 'NopparujVongpatarakulResume.pdf')
            if not os.path.exists(resume_path):
                print(f"Resume not found at: {resume_path}")
                resume_path = os.path.join('resources', 'NopparujVongpatarakulResume.pdf')  # Try lowercase folder
                if not os.path.exists(resume_path):
                    messagebox.showerror("Resume Error", "Resume not found in Resources or resources folder")
                    return
            
            for item in emailSendList:
                email = item.get('Email', '')
                if not re.fullmatch(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b', email):
                    invalid_emails.append(f"{item.get('Name', 'Unknown')}: {email}")
                else:
                    valid_entries.append(item)

            # Show invalid emails warning
            if invalid_emails:
                messagebox.showwarning(
                    "Invalid Emails",
                    f"Skipped {len(invalid_emails)} invalid emails:\n" + 
                    "\n".join(invalid_emails)
                )

            if not valid_entries:
                messagebox.showinfo("No Valid Emails", "No valid emails to send")
                return

            # Process valid entries
            success_count = 0
            failed_emails = []
            base_delay = 2  # Base delay between emails
            
            for idx, item in enumerate(valid_entries):
                try:
                    # Calculate send time
                    now = datetime.datetime.now(datetime.timezone.utc).astimezone()
                    send_time = self.calculate_send_time(now)

                    # Prepare email content
                    msg = MIMEMultipart()
                    msg['From'] = "taz2547@gmail.com"
                    msg['To'] = item['Email']
                    msg['Subject'] = "UIUC student who would love to connect!"
                    
                    # Generate personalized message
                    message_content = self.generate_email_content(item, now)
                    msg.attach(MIMEText(message_content, 'plain'))

                    # Attach resume
                    print(f"Attaching resume from: {resume_path}")
                    with open(resume_path, 'rb') as f:
                        resume = MIMEApplication(f.read(), _subtype='pdf')
                        resume.add_header('Content-Disposition', 'attachment', 
                                        filename='Nopparuj_Vongpatarakul_Resume.pdf')
                        msg.attach(resume)

                    # Create draft message
                    raw = base64.urlsafe_b64encode(msg.as_string().encode()).decode()
                    
                    # Create draft with labels
                    draft = self.gmail_service.users().drafts().create(
                        userId='me',
                        body={
                            'message': {
                                'raw': raw
                            }
                        }
                    ).execute()

                    print(f"Created draft for {item['Email']} with ID: {draft['id']}")
                    
                    success_count += 1

                    # Adaptive delay with randomization
                    delay = random.uniform(base_delay, base_delay * 2)
                    if idx > 0 and idx % 5 == 0:
                        delay += random.uniform(5, 10)  # Extra delay every 5 emails
                    time.sleep(delay)

                except Exception as e:
                    error_msg = f"{item['Email']} - {str(e)}"
                    print(f"Error creating draft: {error_msg}")
                    failed_emails.append(error_msg)
                    time.sleep(random.uniform(10, 15))  # Longer delay after error
                    continue

            # Show final results
            result_msg = [
                f"Successfully created drafts: {success_count}/{len(valid_entries)}",
                f"Failed: {len(failed_emails)}"
            ]
            
            if failed_emails:
                result_msg.append("\nFailed emails:\n" + "\n".join(failed_emails))
                
            messagebox.showinfo(
                "Draft Creation Complete",
                "\n".join(result_msg)
            )

        except Exception as e:
            messagebox.showerror("Email Error", f"Critical failure: {str(e)}")

    def calculate_send_time(self, current_time):
        """Calculate optimal send time with timezone awareness"""
        # Convert to local work hours (9 AM to 5 PM)
        if current_time.weekday() in [0, 1, 2]:  # Mon-Wed
            if current_time.hour < 9:
                send_time = current_time.replace(hour=9, minute=0, second=0, microsecond=0)
            else:
                send_time = current_time + datetime.timedelta(days=1)
        else:
            days_until_monday = (0 - current_time.weekday()) % 7
            send_time = current_time + datetime.timedelta(days=days_until_monday)

        # Ensure time is within business hours
        send_time = send_time.replace(hour=random.randint(9, 16), minute=random.choice([0, 15, 30, 45]))
        
        # Never schedule more than 7 days in advance
        if (send_time - current_time).days > 7:
            send_time = current_time + datetime.timedelta(days=7)
            
        return send_time

    def generate_email_content(self, item, current_time):
        """Generate personalized email content with proper time references"""
        # Determine education background
        education = item['Education']
        student_name = ""
        
        if item['Highschool']:
            education = "ISB"
            student_name = "ISB Panther"
        if item['College']:
            education = "UIUC"
            student_name = "Illini"

        # Personalization
        last_name = item['Name'].strip().split()[-1] if ' ' in item['Name'] else item['Name']
        current_year = current_time.year
        time_reference = "this week" if current_time.weekday() < 3 else "next week"

        # Build message template
        templates = {
            'alumni': f"""Hi {last_name},
    I hope this message finds you well. My name is Nopparuj (Taz), a third year Computer Science student at UIUC from Bangkok, Thailand. I'm reaching out because I'm particularly interested in {item['Company']}'s work in your field as a {item['Position']}.

    Having noticed your background at {education}, I'm inspired by your career path from {education} to {item['Company']}. As a fellow {student_name}, I'd greatly appreciate any insights you might share about breaking into this field.

    Would you have {time_reference} for a brief 15-minute chat about your professional experiences? I completely understand if you're busy - even a quick email response would be incredibly helpful.

    Thank you for considering my request, and I hope to connect soon!

    Best regards,
    Taz""",

            'thai': f"""Hi {last_name},
    I hope this email finds you well. My name is Nopparuj (Taz), a third year Computer Science student at UIUC from Bangkok. I'm reaching out because I'm impressed by {item['Company']}'s innovations in your field as a {item['Position']}.

    As a Thai professional working abroad, your experience is particularly inspiring to me. I'd be grateful for any advice you might have about navigating international tech careers.

    Might you have {time_reference} for a short conversation? I'd be happy to adjust to your schedule.

    Thank you for your time and consideration!

    Best regards,
    Taz""",

            'general': f"""Hi {last_name},
    I hope you're doing well! I'm Nopparuj (Taz), a third year CS student at UIUC from Bangkok. I'm reaching out because I'm keenly interested in {item['Company']}'s work in your role as a {item['Position']}.

    Your career journey from {education} to {item['Company']} is exactly the kind of path I aspire to follow. Could I possibly ask for {time_reference} for a brief chat about your experiences?

    I completely understand if you're too busy - any insights you could share via email would also be greatly appreciated.

    Thank you for your time!

    Best regards,
    Taz"""
        }

        # Select template
        if item['Thai']:
            return templates['thai']
        elif item['College'] or item['Highschool']:
            return templates['alumni']
        else:
            return templates['general']

    def update_sheet_status(self, row_number):
        try:
            range_name = f'Sheet1!L{row_number}'
            body = {'values': [['Yes']]}
            self.sheet_service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_name,
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
        except Exception as e:
            messagebox.showerror("Update Error", str(e))

    def bring_to_front(self):
        """Bring window to front and maintain focus"""
        self.root.lift()  # Lift the window
        self.root.focus_force()  # Force focus on the window


#------------------------------------------------------------------------------------------------------------------------------------------------------#

def get_domain_mappings(sheet_service):
    """Fetch company-domain pairs from Google Sheet"""
    try:
        result = sheet_service.spreadsheets().values().get(
            spreadsheetId=DOMAIN_SHEET_ID,
            range='Sheet1!A2:B'  # Column A: Company Name, B: Domain
        ).execute()
        rows = result.get('values', [])
        return {row[0].lower().strip(): row[1].lower().strip() for row in rows if len(row) >= 2}
    except Exception as e:
        print(f"Error fetching domain mappings: {str(e)}")
        return {}

def get_linkedin_profile_info(self, driver, profile_url):
    try:
        if check_bot_detection(driver):
            raise Exception("LinkedIn bot verification required - aborting scrape")
        
        # Initialize default values
        TR_name = "No Name"
        TR_email = "No Email"
        TR_company = "No Company"
        TR_position = "No Position"
        TR_intern = False
        TR_education_list = []
        TR_education = "No Education Listed"
        TR_college = False
        TR_highschool = False
        TR_thai = False
        
        wait = WebDriverWait(driver, 10)
        
        # 1. Get Name from about-this-profile overlay
        driver.get(profile_url + "overlay/about-this-profile/")
        time.sleep(1)
        try:
            name_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1")))
            TR_name = name_element.text.strip()
        except Exception as e:
            print(f"Error extracting name: {str(e)}")
        
        
        # 2. Get Professional Experience
        driver.get(profile_url + "details/experience/")
        time.sleep(1)
        first_position = True
        try:
            experience_sections = driver.find_elements(By.XPATH, "//div[@data-view-name='profile-component-entity']")
            for section in experience_sections:
                try:
                    roles = section.find_elements(By.XPATH, ".//div[@data-view-name='profile-component-entity']")
                    if roles:
                        company_element = section.find_element(By.XPATH, ".//span[not(contains(@class, 'visually-hidden'))][1]")
                        company_name = company_element.text.strip()
                        for role in roles:
                            position_element = role.find_element(By.XPATH, ".//span[not(contains(@class, 'visually-hidden'))][1]")
                            position_name = position_element.text.strip()
                            if first_position:
                                TR_company = company_name
                                TR_position = position_name
                                first_position = False
                                break
                        if not first_position:
                            break
                    else:
                        company_element = section.find_element(By.XPATH, ".//span[contains(@class, 't-14 t-normal')][1]")
                        company_full = company_element.text.strip().split('\n')[0]
                        company_name = company_full.split(' · ')[0]
                        position_element = section.find_element(By.XPATH, ".//span[not(contains(@class, 'visually-hidden'))][1]")
                        position_name = position_element.text.strip()
                        if first_position:
                            TR_company = company_name
                            TR_position = position_name
                            first_position = False
                            break
                except Exception as e:
                    print(f"Error processing experience section: {str(e)}")
                    continue
        except Exception as e:
            print(f"Error locating experience sections: {str(e)}")

        intern_keywords = {'apprentice', 'apprenticeship', 'intern', 'internship', 'trainee'}
        TR_intern = any(keyword in TR_position.lower() for keyword in intern_keywords)

        # 3. Get Email from contact-info overlay
        email_found = False
        print("\n=== Email Extraction ===")
        
        # Attempt 1: Direct contact info extraction
        try:
            driver.get(profile_url + 'overlay/contact-info/')
            email_element = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//a[contains(@href,'mailto:')]"))
            )
            TR_email = email_element.get_attribute('href').split('mailto:')[1]
            print(f"Found direct email: {TR_email}")
            email_found = True
        except Exception as e:
            print("No email in contact info")

        # Attempt 2: Hunter.io lookup (only if we have valid name and company)
        if not email_found and TR_name != "No Name" and TR_company != "No Company":
            print("Attempting Hunter.io API lookup...")
            
            # Split name into first/last
            name_parts = TR_name.split()
            first_name = name_parts[0] if name_parts else ''
            last_name = ' '.join(name_parts[1:]) if len(name_parts) > 1 else ''
            
            # Initialize API client
            with open("config.json") as f:
                config = json.load(f)
            hunter = HunterIOAPI(config['hunter_api_key'])
            
            # Try email finder
            email_data = hunter.find_email(
                first_name=first_name,
                last_name=last_name,
                company=TR_company
            )
            
            if email_data and email_data['confidence'] is not None and email_data['confidence'] > 10:
                TR_email = email_data['email']
                print(f"Found email via API: {TR_email}")
                email_found = True
                
            if not email_found:
                print("Trying Email Finder fallback with domain...")
                domain_mappings = get_domain_mappings(self.sheet_service)
                clean_company = TR_company.lower().strip()
                if clean_company in domain_mappings:
                    domain = domain_mappings[clean_company]
                    print(f"Using mapped domain from sheet: {domain}")
                else:
                    domain = f"{clean_company.replace(' ', '')}.com"
                    print(f"Using generated domain: {domain}")
                
                # Use Email Finder with full details including full_name
                email_data = hunter.find_email(
                    first_name=first_name,
                    last_name=last_name,
                    company=TR_company,
                    domain=domain,
                    full_name=TR_name  # Provide full name for better accuracy
                )
                print(email_data)
                
                if email_data and (email_data.get('confidence') or 0) > 10:
                    TR_email = email_data['email']
                    print(f"Using Email Finder fallback result: {TR_email}")
                    email_found = True
                    
            # Verify email if found
            if email_found:
                status = hunter.verify_email(TR_email)
                if status not in ['valid', 'accept_all']:
                    print(f"Email verification failed ({status}), discarding...")
                    email_found = False


        # Final handling for failed lookup
        if not email_found:
            name_display = TR_name if TR_name != "No Name" else "this profile"
            print(f"\n⚠️ No email found for {name_display}")
            return None  # Signal to skip this entry
        
        # 4. Get Education
        driver.get(profile_url + "details/education/")
        time.sleep(1)
        try:
            edu_sections = driver.find_elements(By.XPATH, "//div[@data-view-name='profile-component-entity']")
            for edu in edu_sections:
                try:
                    school_element = edu.find_element(By.XPATH, ".//span[not(contains(@class, 'visually-hidden'))]")
                    school_name = school_element.text.strip()
                    TR_education_list.append(school_name)  # Add to list
                    if TR_education == "No Education Listed":
                        TR_education = school_name
                    # Check for target schools
                    if "University of Illinois Urbana-Champaign" in school_name:
                        TR_college = True
                    if "International School Bangkok" in school_name:
                        TR_highschool = True
                except Exception as e:
                    print(f"Error processing education entry: {str(e)}")
        except Exception as e:
            print(f"Error locating education sections: {str(e)}")
        
        # 5. Check Thai Language and Education
        driver.get(profile_url + "details/languages/")
        time.sleep(1)
        try:
            language_sections = driver.find_elements(By.XPATH, "//div[@data-view-name='profile-component-entity']")
            for lang in language_sections:
                try:
                    # Get both the language name and proficiency level
                    lang_text = lang.text.strip().split('\n')
                    if len(lang_text) >= 2 and "Thai" in lang_text[0]:
                        proficiency = lang_text[2].lower()
                        print(f"Found Thai language with proficiency: {proficiency}")
                        if "native" in proficiency or "bilingual" in proficiency:
                            TR_thai = True
                            break
                except Exception as e:
                    print(f"Error processing language entry: {str(e)}")
        except Exception as e:
            print(f"Error locating language sections: {str(e)}")

        # Check for Thai schools in education list
        thai_schools = {
            # Universities
            'Chulalongkorn University', 'Mahidol University', 'Kasetsart University', 
            'Thammasat University', 'King Mongkut', 'Assumption University',
            'Bangkok University', 'Rangsit University', 'Silpakorn University',
            'Srinakharinwirot University', 'ABAC', 'KMUTT', 'KMITL',
            # International Schools
            'International School Bangkok', 'Bangkok Patana', 'NIST International School',
            'Ruamrudee International School', 'Shrewsbury International School',
            'Harrow International School', 'Bangkok Prep', 'KIS International School',
            'SISB', 'Wells International School'
        }
        
        TR_thai_education = any(any(school.lower() in edu.lower() for school in thai_schools) 
                              for edu in TR_education_list)
        TR_thai = TR_thai or TR_thai_education
        
    except Exception as e:
        print(f"Critical error during scraping: {str(e)}")
        raise
    
    return {
        'name': TR_name or "N/A",
        'company': TR_company or "N/A",
        'position': TR_position or "N/A",
        'email': TR_email or "N/A",
        'education': TR_education or "N/A",
        'college': TR_college,
        'highschool': TR_highschool,
        'thai': TR_thai,
        'intern': TR_intern
    }


    #------------------------------------------------------------------------------------------------------------------------------------------------------#


def append_row_to_sheet(sheet_service, spreadsheet_id, link, name, company, position, email, education, same_college, same_highschool, thai, force=False):
    #💬 DUPLICATE CHECK: Verify link doesn't exist in column K unless forcing
    if email in ["No Email", "N/A"]:
        print("Skipping sheet entry - no valid email")
        return

    if not force:
        existing = sheet_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range="Sheet1!K:K"  #💬 Column K contains profile links
        ).execute().get('values', [])
        
        existing_links = [item[0] for item in existing if item]
        if link in existing_links:
            raise ValueError(f"Duplicate link: {link}")
        
    # Get the current date in "dd/mm/yyyy" format
    current_date = datetime.datetime.now().strftime("%d/%m/%Y")
    connections = []
    if same_college: connections.append("Same College")
    if same_highschool: connections.append("Same Highschool")
    if thai: connections.append("Thai")
    
    values = [
        ["N/A", "Not Yet", name, company, "", "", education, position, 
         ", ".join(connections) if connections else "N/A", email, link,
         "No", "No", "No", "No", "No", "", "Process", current_date]
    ]

    # Append to sheet
    body = {'values': values}
    sheet_service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range="Sheet1!A:A",
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()


    #------------------------------------------------------------------------------------------------------------------------------------------------------#


def sortSheet(sheet_service, row, ascending):
    # Fetch the sheet metadata to get the Sheet ID
    sheet_metadata = sheet_service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = sheet_metadata.get('sheets', '')
    sheet_id = sheets[0].get("properties", {}).get("sheetId")

    if (ascending) :
        order = "ASCENDING"
    else:
        order = "DESCENDING"

    # Prepare the request body for sorting
    request_body = {
        "requests": [
            {
                "sortRange": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 1, # Skip first row which is headers
                    },
                    "sortSpecs": [
                        {
                            "dimensionIndex": row, # Zero-indexed row
                            "sortOrder": order  # Ascending: A to Z, Descending: Z to A
                        }
                    ]
                }
            }
        ]
    }

    # Send the batchUpdate request to sort the sheet
    response = sheet_service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body=request_body
    ).execute()

    print("Sheet sorted:", response)


    #------------------------------------------------------------------------------------------------------------------------------------------------------#

def main():
    try:
        app = LinkedInApp()
        
        app.root.mainloop()
    except Exception as e:
        print(f"An error occurred: {e}")
        if app.driver:
            app.driver.quit()

if __name__ == "__main__":
    main()


