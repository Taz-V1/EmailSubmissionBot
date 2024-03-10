import smtplib
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

import datetime
import time

SPREADSHEET_ID = '13AK7mmpuRHUc7yyP9l7VwEJKPLJ26MPja5o05Ib-J88'


def get_linkedin_profile_info(driver, profile_url):
    try:
        # Stuff to return
        TR_name = "NoName"
        TR_email = "NoEmail"
        TR_company = "NoCompany"
        TR_position = "NoPosition"
        TR_college = False
        TR_highschool = False
        TR_thai = False


        #------------------------------------------------------------------------------------------------------------------------------------------------------#

        # Get name
        print("Name:")
        driver.get(profile_url)
        time.sleep(1)
        name_xpath = "//h1[contains(@class,'text-heading-xlarge')]"
        name_element = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, name_xpath))
        )
        print(name_element.text)
        TR_name = name_element.text
        print()
        print()


        #------------------------------------------------------------------------------------------------------------------------------------------------------#

        # Get email, check first if email exists in contact
        emailFound = False
        print("Email:")
        driver.get(profile_url + 'overlay/contact-info/')
        time.sleep(1)
        try:
            # Find the email section by class name
            email_section = driver.find_element(By.XPATH, "//section[contains(@class, 'pv-contact-info__contact-type')]//a[contains(@href, 'mailto:')]")

            # Extract the email address from the href attribute of the <a> tag
            email_href = email_section.get_attribute('href')
            email_address = email_href.split(':')[1]  # Split 'mailto:' and get the email part
            print("Email found:", email_address)
            TR_email = email_address
            emailFound = True

        except NoSuchElementException:
            print("Nothing found")

        print()
        print()


        #------------------------------------------------------------------------------------------------------------------------------------------------------#
        
        # Get Professional Experiences, will have to check if aligns with what you want
        print("Experiences:")
        driver.get(profile_url + 'details/experience/')
        time.sleep(1)
        firstElem = True

        # Load page elements
        elements = driver.find_elements(By.XPATH, "//div[@data-view-name='profile-component-entity']")

        # Check if loaded correctly
        if not elements:
            print("Nothing found")
        else:
            # Select all the top-level experience entries
            main_xpath = "//main[@class='scaffold-layout__main']"
            experience_entries_xpath = f"{main_xpath}//div[@data-view-name='profile-component-entity' and not(ancestor::div[@data-view-name='profile-component-entity'])]"
            experience_elements = driver.find_elements(By.XPATH, experience_entries_xpath)

            for element in experience_elements:
                # Check for nested role entries under the current company
                role_entries_xpath = ".//div[@data-view-name='profile-component-entity']"
                role_elements = element.find_elements(By.XPATH, role_entries_xpath)

                if role_elements:
                    # If there are nested roles, the first span is the company name
                    company_name_element = element.find_element(By.XPATH, ".//span[not(contains(@class, 'visually-hidden'))][1]")
                    company_name = company_name_element.get_attribute('innerText').strip()
                    print(f"Company: {company_name}")

                    for role_element in role_elements:
                        # Extract the role name from nested elements
                        role_name_element = role_element.find_element(By.XPATH, ".//span[not(contains(@class, 'visually-hidden'))][1]")
                        role_name = role_name_element.get_attribute('innerText').strip()
                        print(f"Role: {role_name}")
                        if firstElem:
                            TR_company = company_name
                            TR_position = role_name
                            firstElem = False

                else:
                    # If there are no nested roles, the first span is the role and the second span is the company
                    company_name_element = element.find_element(By.XPATH, ".//span[contains(@class, 't-14') and contains(@class, 't-normal') and not(contains(@class, 'visually-hidden'))][1]")
                    company_name_full = company_name_element.get_attribute('innerText').split('\n')[0].strip()

                    # Split the company name if necessary
                    company_name = company_name_full.split(' · ')[0]

                    role_name_element = element.find_element(By.XPATH, ".//span[not(contains(@class, 'visually-hidden'))][1]")
                    role_name = role_name_element.get_attribute('innerText').strip()

                    print(f"Company: {company_name}")
                    print(f"Role: {role_name}")

                    if firstElem:
                        TR_company = company_name
                        TR_position = role_name
                        firstElem = False


        print()
        print()


        #------------------------------------------------------------------------------------------------------------------------------------------------------#
            
        # Get Education
        # NOTE: NEED TO ALSO EXTRACT MAJOR
        print("Education:")
        driver.get(profile_url + 'details/education/')
        time.sleep(1)

        # Xpath, gets from the "main" page component
        main_xpath = "//main[@class='scaffold-layout__main']"
        education_entries_xpath = f"{main_xpath}//div[@data-view-name='profile-component-entity']"

        education_elements = []

        # Check if loaded correctly
        try:
            education_elements = driver.find_elements(By.XPATH, education_entries_xpath)
            if not language_elements:
                print("Nothing found")
            else:
                for element in education_elements:
                    # Update the XPath to be relative to the main element
                    school_name_element = element.find_element(By.XPATH, ".//span[not(contains(@class, 'visually-hidden'))]")
                    school_name = school_name_element.get_attribute('innerText').strip()
                    print(school_name)
                    
                    # Check for specific matching schools
                    if "University of Illinois Urbana-Champaign" in school_name:
                        print("Found matching education:", school_name)
                        TR_college = True
                    if "International School Bangkok" in school_name:
                        print("Found matching education:", school_name)
                        TR_highschool = True
                    
        except Exception:
            print("Nothing found")

        print()
        print()

        
        #------------------------------------------------------------------------------------------------------------------------------------------------------#
                
        # Check if Nationality Matches Through Native Language
        print("Language:")
        driver.get(profile_url + 'details/languages/')
        time.sleep(1)

        # Xpath, gets from the "main" page component
        main_xpath = "//main[@class='scaffold-layout__main']"
        languages_xpath = f"{main_xpath}//div[@data-view-name='profile-component-entity']"

        language_elements = []

        # Check if loaded correctly
        try:
            language_elements = driver.find_elements(By.XPATH, education_entries_xpath)

            if not language_elements:
                print("Nothing found")
            else:
                for element in language_elements:
                    # Update the XPath to be relative to the main element
                    language_element = element.find_element(By.XPATH, ".//span[not(contains(@class, 'visually-hidden'))]")
                    language = language_element.get_attribute('innerText').strip()
                    print(language)

                    # Check if the language is Thai
                    if "Thai" in language:
                        print("Found matching language:", language)
                        
                        # Check for proficiency level
                        proficiency_xpath = ".//span[contains(@class, 'visually-hidden')]"
                        proficiency_elements = element.find_elements(By.XPATH, proficiency_xpath)

                        # Iterate over proficiency elements to find "Native proficiency"
                        for proficiency_element in proficiency_elements:
                            proficiency_text = proficiency_element.get_attribute('innerText').strip()
                            if "Native" in proficiency_text:
                                print("Proficiency level:", proficiency_text)
                                TR_thai = True
        
        except Exception:
            print("Nothing found")

        print()
        print()
        

        #------------------------------------------------------------------------------------------------------------------------------------------------------#
                
        # Get Email if not Found Beforehand
        # if not emailFound:
        #     print("Oof")

    #------------------------------------------------------------------------------------------------------------------------------------------------------#

    except Exception as e:
        print(f"An error occurred: {e}")


    return TR_name, TR_company, TR_position, TR_email, TR_college, TR_highschool, TR_thai


    #------------------------------------------------------------------------------------------------------------------------------------------------------#


def append_row_to_sheet(sheet_service, link, name, company, position, email, same_college, same_highschool, thai):
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
        ["N/A", "Not Yet", name, company, "", "", position, connection_value, email, link, "No", "No", "No", "No", "No", "", current_date]
    ]

    # Append the values to the sheet
    body = {'values': values}
    sheet_service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID, range="Sheet1!A" + str(next_empty_row),
        valueInputOption="USER_ENTERED", body=body).execute()


    #------------------------------------------------------------------------------------------------------------------------------------------------------#

# def send_email(profile_info):
#     # Set up email server and compose message
#     server = smtplib.SMTP('smtp.example.com', 587)
#     server.starttls()
#     server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    
#     msg = MIMEMultipart()
#     msg['From'] = EMAIL_ADDRESS
#     msg['To'] = profile_info['email']
#     msg['Subject'] = f"Coffee chat invitation from {profile_info['name']}"

#     body = f"Hello {profile_info['name']},\n\nI came across your profile and I'm impressed by your work as a {profile_info['job_title']}. I would love to have a coffee chat with you to discuss potential collaboration opportunities. Please let me know if you would be interested.\n\nBest regards,\nYour Name"
#     msg.attach(MIMEText(body, 'plain'))
    
#     server.send_message(msg)
#     server.quit()


def main():

    # Personal info and desired info
    # If turn into app with I/O, will need the following lines to be inputs
    my_college = "University of Illinois Urbana-Champaign"
    my_highschool = "International School Bangkok"
    desired_job = ""

    # LinkedIn profile URLs
    profile_urls = []



    #------------------------------------------------------------------------------------------------------------------------------------------------------#

    # Info to set up and connect to Google API
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SERVICE_ACCOUNT_FILE = '/Users/tazvongpatarakul/Desktop/LinkedIn Bot/velvety-study-415321-4b33a33e683e.json'

    credentials = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    sheet_service = build('sheets', 'v4', credentials=credentials)

    SPREADSHEET_ID = '13AK7mmpuRHUc7yyP9l7VwEJKPLJ26MPja5o05Ib-J88'


    #------------------------------------------------------------------------------------------------------------------------------------------------------#

    # Initialize Chrome options and Driver
    chrome_options = Options()
    # LeadLeaper Extension
    # chrome_options.add_argument('--load-extension=/Users/tazvongpatarakul/Desktop/CodingStuff/LeadLeaper_Extension/7.1.20_0')

    # Initialize WebDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)


    #------------------------------------------------------------------------------------------------------------------------------------------------------#

    # Username: throwaway5621345798@gmail.com
    # Password: ThrowAwayThrowAway
    # time.sleep(60)


    #------------------------------------------------------------------------------------------------------------------------------------------------------#

    # Navigate to the login page
    login_url = 'https://www.linkedin.com/login'
    driver.get(login_url)

    # Enter login credentials
    # Use throwaway account while extracting info
    time.sleep(2)
    username = driver.find_element(By.ID, 'username')
    username.send_keys('taz2547sub@gmail.com')
    password = driver.find_element(By.ID, 'password')
    password.send_keys('1234512345')

    # Submit the form
    password.send_keys(Keys.RETURN)

    # Can have time delay for human verification, until there is a better solution
    time.sleep(60)

    # Wait for login to complete and navigate to the profile page
    time.sleep(1)


    #------------------------------------------------------------------------------------------------------------------------------------------------------#

    names = []
    companies = []
    positions = []
    emails = []
    collegeBool = []
    highschoolBool = []
    thaiBool = []

    for i in range(len(profile_urls)):
        TR_name, TR_company, TR_position, TR_email, TR_college, TR_highschool, TR_thai = get_linkedin_profile_info(driver, profile_urls[i])
        names.append(TR_name)
        companies.append(TR_company)
        positions.append(TR_position)
        emails.append(TR_email)
        collegeBool.append(TR_college)
        highschoolBool.append(TR_highschool)
        thaiBool.append(TR_thai)


    #------------------------------------------------------------------------------------------------------------------------------------------------------#

    for i in range(len(profile_urls)):
        append_row_to_sheet(sheet_service, profile_urls[i], names[i], companies[i], positions[i], emails[i], collegeBool[i], highschoolBool[i], thaiBool[i])


    #------------------------------------------------------------------------------------------------------------------------------------------------------#

    driver.quit()



if __name__ == "__main__":
    main()
