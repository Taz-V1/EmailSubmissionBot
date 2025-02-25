import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

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

from google.oauth2.service_account import Credentials
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

SPREADSHEET_ID = '13AK7mmpuRHUc7yyP9l7VwEJKPLJ26MPja5o05Ib-J88'
DOMAIN_SHEET_ID = '1DCUQw7c92AEKk0pURXegYp_Vy2TklDVijaoz0OMnTpw'

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send'
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
        self.driver = None
        self._processing = False
        self.current_url_index = 0
        self.hunter_authenticated = False
        
        # Initialize services
        try:
            creds = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
            self.sheet_service = build('sheets', 'v4', credentials=creds)
            self.gmail_service = build('gmail', 'v1', credentials=creds)
        except Exception as e:
            messagebox.showerror("API Error", f"Failed to initialize services: {str(e)}")
            self.root.destroy()

        self.create_gui()
        self.create_review_treeview()

    def on_close(self):
        # Clean up resources when closing the app
        if self.driver:
            self.driver.quit()
        self.root.destroy()

    def create_gui(self):
        # Main Window
        self.root.title("LinkedIn Manager")
        
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
        self.tree = ttk.Treeview(self.review_frame, columns=(  # Add parent here
            'Name', 'Email', 'Company', 'Position', 'Education',
            'College', 'Highschool', 'Thai', 'URL'
        ))
        
        # Configure columns first
        for col in ['Position', 'Education', 'College', 'Highschool', 'Thai', 'URL']:
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
            column = self.tree.identify_column(event.x)
            item = self.tree.identify_row(event.y)
            
            # Get current value
            col_index = int(column[1:])-1
            current_value = self.tree.item(item, 'values')[col_index]
            
            # Create edit window
            edit_win = tk.Toplevel()
            edit_win.title("Edit Value")
            
            # Entry widget
            entry = ttk.Entry(edit_win)
            entry.insert(0, current_value)
            entry.pack(padx=10, pady=10)
            
            # Save button
            ttk.Button(edit_win, text="Save", 
                     command=lambda: self.save_edited_value(item, col_index, entry.get(), edit_win)).pack()
            
    # def hunter_login(self):
    #     """Log into Hunter.io service"""
    #     print("Authenticating with Hunter.io...")
    #     try:
    #         with open("config.json") as f:
    #             config = json.load(f)
                
    #         self.driver.get("https://hunter.io/users/sign_in")
            
    #         # Email field
    #         email_field = WebDriverWait(self.driver, 25).until(
    #             EC.presence_of_element_located((By.ID, "email-field"))
    #         )
    #         print("Entering email...")
    #         human_type(email_field, config['hunter_email'])
    #         time.sleep(random.uniform(0.5, 1.5))
            
    #         # Password field
    #         password_field = self.driver.find_element(By.ID, "password-field")
    #         print("Entering password...")
    #         human_type(password_field, config['hunter_password'])
    #         time.sleep(random.uniform(0.5, 1.5))
            
    #         # Click login
    #         signin_button = WebDriverWait(self.driver, 10).until(
    #             EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Sign in')]"))
    #         )
    #         signin_button.click()
            
    #         time.sleep(1000)

    #         # Verify login
    #         WebDriverWait(self.driver, 30).until(
    #             EC.presence_of_element_located((By.XPATH, "//a[contains(@href,'/logout')]"))
    #         )
    #         print("Hunter.io login successful")
    #         self.hunter_authenticated = True
            
    #     except Exception as e:
    #         print(f"Hunter.io login failed: {str(e)}")
    #         self.hunter_authenticated = False
            
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

        self.profile_urls = [url.strip() for url in self.url_text.get("1.0", tk.END).splitlines() 
                            if url.startswith('https://www.linkedin.com/in/')]
        
        if not self.profile_urls:
            messagebox.showwarning("No URLs", "No valid LinkedIn URLs found")
            return

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


    def process_next_profile(self):
        if not self._processing or self.current_url_index >= len(self.profile_urls):
            self._processing = False
            self.update_ui_status("Processing complete")
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
            url
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

    def save_to_sheets(self):
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            try:
                values = self.tree.item(item)['values']
                append_row_to_sheet(
                    self.sheet_service,
                    link=values[8],  # URL remains at index 8 in treeview
                    name=values[0],   # name
                    company=values[2],# company
                    position=values[3],# position
                    email=values[1],  # email
                    education=values[4],# education
                    same_college=values[5],# college
                    same_highschool=values[6],# highschool
                    thai=values[7]    # thai
                )
            except Exception as e:
                messagebox.showerror("Save Error", str(e))

    def prepare_emails(self):
        try:
            result = self.sheet_service.spreadsheets().values().get(
                spreadsheetId=SPREADSHEET_ID, 
                range='Sheet1!A2:L'
            ).execute()
            rows = result.get('values', [])
            
            emailSendList = []
            for row_idx, row in enumerate(rows):
                if len(row) >= 12 and row[11] == 'No' and re.fullmatch(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b', row[9]):
                    filtered_row = {
                        'Name': row[2] if len(row) > 2 else '',
                        'Company': row[3] if len(row) > 3 else '',
                        'Position': row[7] if len(row) > 7 else '',
                        'Email': row[9] if len(row) > 9 else '',
                        'Education': row[6] if len(row) > 6 else '',
                        'College': 'Same College' in row[8] if len(row) > 8 else False,
                        'Highschool': 'Same Highschool' in row[8] if len(row) > 8 else False,
                        'Thai': 'Thai' in row[8] if len(row) > 8 else False,
                        'row_number': row_idx + 2  # +2 because sheet starts at row 2
                    }
                    emailSendList.append(filtered_row)
            
            if emailSendList:
                self.send_email(emailSendList)
            else:
                messagebox.showinfo("No Emails", "No pending emails to send")
                
        except Exception as e:
            messagebox.showerror("Sheet Error", str(e))

    def send_email(self, emailSendList):
        try:
            # Validate emails first
            invalid_emails = []
            valid_entries = []
            
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
            
            for idx, item in enumerate(valid_entries):
                try:
                    # Calculate send time with improved timezone handling
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

                    # Create and send message
                    raw = base64.urlsafe_b64encode(msg.as_string().encode()).decode()
                    body = {'raw': raw}
                    
                    if send_time > now:
                        body['internalDate'] = str(int(send_time.timestamp() * 1000))

                    self.gmail_service.users().messages().send(
                        userId='me',
                        body=body
                    ).execute()

                    # Update sheet status
                    self.update_sheet_status(item['row_number'])
                    success_count += 1

                    # Rate limiting with random delay
                    time.sleep(random.uniform(2, 5))  # More conservative delay

                except Exception as e:
                    failed_emails.append(f"{item['Email']} - {str(e)}")
                    # Continue processing other emails even if one fails

            # Show final results
            result_msg = [
                f"Successfully scheduled: {success_count}/{len(valid_entries)}",
                f"Failed: {len(failed_emails)}"
            ]
            
            if failed_emails:
                result_msg.append("\nFailed emails:\n" + "\n".join(failed_emails))
                
            messagebox.showinfo(
                "Sending Complete",
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
    I hope this message finds you well. My name is Nopparuj (Taz), a second year Computer Science student at UIUC from Bangkok, Thailand.
    I'm reaching out because I'm particularly interested in {item['Company']}'s work in your field as a {item['Position']}.

    Having noticed your background at {education}, I'm inspired by your career path from {education} to {item['Company']}. 
    As a fellow {student_name}, I'd greatly appreciate any insights you might share about breaking into this field.

    Would you have {time_reference} for a brief 15-minute chat about your professional experiences? I completely understand if you're busy - 
    even a quick email response would be incredibly helpful.

    Thank you for considering my request, and I hope to connect soon!

    Best regards,
    Taz""",

            'thai': f"""Hi {last_name},
    I hope this email finds you well. My name is Nopparuj (Taz), a second year Computer Science student at UIUC from Bangkok.
    I'm reaching out because I'm impressed by {item['Company']}'s innovations in your field as a {item['Position']}.

    As a Thai professional working abroad, your experience is particularly inspiring to me. 
    I'd be grateful for any advice you might have about navigating international tech careers.

    Might you have {time_reference} for a short conversation? I'd be happy to adjust to your schedule.

    Thank you for your time and consideration!

    Best regards,
    Taz""",

            'general': f"""Hi {last_name},
    I hope you're doing well! I'm Nopparuj (Taz), a second year CS student at UIUC from Bangkok.
    I'm reaching out because I'm keenly interested in {item['Company']}'s work in your role as a {item['Position']}.

    Your career journey from {education} to {item['Company']} is exactly the kind of path I aspire to follow. 
    Could I possibly ask for {time_reference} for a brief chat about your experiences?

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
                    # Check for nested positions (multiple roles at same company)
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
                        # Handle single position entries
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
            
            if email_data and email_data['confidence'] is not None and email_data['confidence'] > 70:
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
        
        # 5. Check Thai Language
        driver.get(profile_url + "details/languages/")
        time.sleep(1)
        try:
            language_sections = driver.find_elements(By.XPATH, "//div[@data-view-name='profile-component-entity']")
            for lang in language_sections:
                try:
                    lang_element = lang.find_element(By.XPATH, ".//span[not(contains(@class, 'visually-hidden'))]")
                    language = lang_element.text.strip()
                    if "Thai" in language:
                        # Check for native proficiency
                        proficiency_element = lang.find_element(By.XPATH, ".//span[contains(@class, 'visually-hidden')]")
                        proficiency = proficiency_element.get_attribute('innerText').strip()
                        if "Native" in proficiency:
                            TR_thai = True
                            break
                except Exception as e:
                    print(f"Error processing language entry: {str(e)}")
        except Exception as e:
            print(f"Error locating language sections: {str(e)}")
        
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
        'thai': TR_thai
    }


    #------------------------------------------------------------------------------------------------------------------------------------------------------#


def append_row_to_sheet(sheet_service, link, name, company, position, email, education, same_college, same_highschool, thai, force=False):
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

    # Find the next empty row in the sheet
    result = sheet_service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range="Sheet1").execute()
    rows = result.get('values', [])
    next_empty_row = len(rows) + 1

    # Format the "Connection?" column based on boolean arguments
    connections = []
    if same_college:
        connections.append("Same College")
    if same_highschool:
        connections.append("Same Highschool")
    if thai:
        connections.append("Thai")
    connection_value = ", ".join(connections) if connections else "N/A"

    # Prepare the values to be inserted
    values = [
        ["N/A", "Not Yet", name, company, "", "", education, position, connection_value, email, link, "No", "No", "No", "No", "No", "", "Process", current_date]
    ]

    # Append the values to the sheet
    body = {'values': values}
    sheet_service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID, range="Sheet1!A" + str(next_empty_row),
        valueInputOption="USER_ENTERED", body=body).execute()


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

# def extractEmail(first_name, last_name, company):
#     # Set up email server and compose message


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