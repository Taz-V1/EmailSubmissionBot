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
        TR_education = "NoEducation"
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

                    if TR_education == "NoEducation":
                        TR_education = school_name
                    
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


    return TR_name, TR_company, TR_position, TR_email, TR_education, TR_college, TR_highschool, TR_thai


    #------------------------------------------------------------------------------------------------------------------------------------------------------#


def append_row_to_sheet(sheet_service, link, name, company, position, email, education, same_college, same_highschool, thai):
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

def send_email(emailSendList, from_email, password):
    # Set up email server and compose message
    server = smtplib.SMTP(host='smtp.gmail.com', port=587)
    server.starttls()
    server.login(from_email, password)

    # Create a message for each sent data
    for item in emailSendList:
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = item['Email']
        msg['Subject'] = "UIUC student who would love to connect!"

        # Variables/specifics for message content
        lastName = ""
        trimmedName = lastName.strip()
        lastSpaceIdx = trimmedName.rfind(' ')
        lastName = trimmedName[lastSpaceIdx + 1:] if lastSpaceIdx != -1 else trimmedName
        
        # NOTE: IF UNKNOWN PRONOUN CHECK, THEN SET LAST NAME TO FULL NAME

        education = item['Education']
        studentName = ""
        if item['College'] == True:
            education = "UIUC"
            studentName = "Illini"
        elif item['Highschool'] == True:
            education = "ISB"
            studentName = "ISB Panther"
        
        Thai = ""
        if item['Thai'] == True:
            Thai = "Thai "


        message = ""

        if (item['College'] == True or item['Education'] == True):
            message = (f"Hi, (insert gender pronoun). {lastName}, \n\n"
                    "I hope this email finds you well. My name is Nopparuj (Taz), and I'm a second year student at UIUC, majoring in computer science from Bangkok, Thailand. "
                    f"I am reaching out because I am really interested in learning more about {item['Company']}. \n\n"

                    f"I was particularly inspired to contact you upon learning that you completed your studies at {education}. "
                    f"It's great to see a fellow {Thai}{studentName} in such an impactful and exciting company. "
                    f"Your journey from {item['Education']} to working on impactful projects as a {item['Position']} at {item['Company']} is the path I aspire to follow. "
                    "As a result, I would greatly appreciate the chance to chat with you about your professional experiences. \n\n"

                    "I understand that you are very busy. If you have any availability in the coming weeks to join me on a brief call, I would love to set something up. "
                    "Thank you so much for your time, and I look forward to hearing from you soon! \n\n"

                    "Best regards, \n"
                    "Taz")
            
        elif (item['Thai'] == True):
            message = (f"Hi, (insert gender pronoun). {lastName}, \n\n"
                    "I hope this email finds you well. My name is Nopparuj (Taz), and I'm a second year student at UIUC, majoring in computer science from Bangkok, Thailand. "
                    f"I am reaching out because I am really interested in learning more about {item['Company']}. \n\n"

                    "I was particularly inspired to contact you upon learning that you are from Thailand. "
                    "It's great to see fellow Thai citizens making significant contributions within such a dynamic organization. "
                    f"Your journey from {item['Education']} to working on impactful projects as a {item['Position']} at {item['Company']} is incredibly inspiring. "
                    "As a result, I would greatly appreciate the chance to chat with you about your professional experiences. \n\n"

                    "I understand that you are very busy. If you have any availability in the coming weeks to join me on a brief call, I would love to set something up. "
                    "Thank you so much for your time, and I look forward to hearing from you soon! \n\n"

                    "Best regards, \n"
                    "Taz")
            
        else:
            message = (f"Hi, (insert gender pronoun). {lastName}, \n\n"
                    f"I hope this email finds you well. My name is Nopparuj (Taz), and I'm a second year student at UIUC, majoring in computer science from Bangkok, Thailand. "
                    f"I am reaching out because I am really interested in learning more about {item['Company']}. \n\n"

                    f"Your journey from {item['Education']} to working on impactful projects as a {item['Position']} at {item['Company']} is incredibly inspiring. "
                    "As a result, I would greatly appreciate the chance to chat with you about your professional experiences. \n\n"

                    "I understand that you are very busy. If you have any availability in the coming weeks to join me on a brief call, I would love to set something up. "
                    "Thank you so much for your time, and I look forward to hearing from you soon! \n\n"

                    "Best regards, \n"
                    "Taz")
        

        msg.attach(MIMEText(message, 'plain'))

        # Send the message
        server.send_message(msg)

    server.quit()

    #CALL TO SET THIS ROW IN SHEETS TO SENT


    #------------------------------------------------------------------------------------------------------------------------------------------------------#

# def extractEmail(first_name, last_name, company):
#     # Set up email server and compose message


    #------------------------------------------------------------------------------------------------------------------------------------------------------#


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

    # UNCOMMENT THIS TO RUN
    # Initialize Chrome options and Driver
    chrome_options = Options()

    # Initialize WebDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)


    #------------------------------------------------------------------------------------------------------------------------------------------------------#

    # Username: throwaway5621345798@gmail.com
    # Password: ThrowAwayThrowAway
    # time.sleep(60)


    #------------------------------------------------------------------------------------------------------------------------------------------------------#

    # # Navigate to the login page
    # login_url = 'https://www.linkedin.com/login'
    # driver.get(login_url)

    # # Enter login credentials
    # # Use throwaway account while extracting info
    # time.sleep(2)
    # username = driver.find_element(By.ID, 'username')
    # username.send_keys('taz2547sub@gmail.com')
    # password = driver.find_element(By.ID, 'password')
    # password.send_keys('1234512345')

    # # Submit the form
    # password.send_keys(Keys.RETURN)

    # # Can have time delay for human verification, until there is a better solution
    # time.sleep(60)

    # # Wait for login to complete and navigate to the profile page
    # time.sleep(1)


    #------------------------------------------------------------------------------------------------------------------------------------------------------#

    # # Arrays to store information extracted from LinkedIn
    # names = []
    # companies = []
    # positions = []
    # emails = []
    # education = []
    # collegeBool = []
    # highschoolBool = []
    # thaiBool = []

    # # Scrape info from linkedIn links (NEXT STEPS, IMPLEMENT SEARCH FUNCTIONALITY, CHECK WITH API IF IN LIST YET)
    # for i in range(len(profile_urls)):
    #     TR_name, TR_company, TR_position, TR_email, TR_education, TR_college, TR_highschool, TR_thai = get_linkedin_profile_info(driver, profile_urls[i])
    #     names.append(TR_name)
    #     companies.append(TR_company)
    #     positions.append(TR_position)
    #     emails.append(TR_email)
    # if (TR_education == "University of Illinois Urbana-Champaign"):
    #     education.append("UIUC")
    # else:
    #     education.append(TR_education)
    #     collegeBool.append(TR_college)
    #     highschoolBool.append(TR_highschool)
    #     thaiBool.append(TR_thai)

    # # Append extracted info into google sheets
    # for i in range(len(profile_urls)):
    #     append_row_to_sheet(sheet_service, profile_urls[i], names[i], companies[i], positions[i], emails[i], education[i], collegeBool[i], highschoolBool[i], thaiBool[i])


    #------------------------------------------------------------------------------------------------------------------------------------------------------#
        
    # Extract info from sheets API

    # Sorts column 10,(Column K: Coffee chat sent), so "no" entries are at the top
    sortSheet(sheet_service, 10, True)

    # Read data from the sheet
    result = sheet_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range='A2:L').execute()
    rows = result.get('values', [])

    # Filter rows where column L (index 11) is 'No' and extract specific columns
    emailSendList = []
    for row in rows:
        if row[11] == 'No':
            # Extract Name (C, index 2), Company (D, index 3), Education (G, index 6), Position (H, index 7), 
            # College? (I, index 8), Highschool? (I, index 8), Thai? (I, index 8), and Gmail (J, index 9)

            # Check if these characteristics exist in "Connection?" column
            college = "Same College" in row[8]
            highschool = "Same Highschool" in row[8]
            thai = "Thai" in row[8]

            filtered_row = {
                'Name': row[2] if len(row) > 2 else '',
                'Company': row[3] if len(row) > 3 else '',
                'Education': row[6] if len(row) > 3 else '',
                'Position': row[7] if len(row) > 6 else '',
                'College': college,
                'Highschool': highschool,
                'Thai': thai,
                'Email': row[9] if len(row) > 8 else ''
            }
            emailSendList.append(filtered_row)
        else:
            break

    # Send out emails
    send_email(listYay, "taz2547sub2@gmail.com", "enzq rtiu hgrq asox")

    
    # TEST CODE: DELETE LATER    
    listYay = []
    filtered_row = {
        'Name': "Taz",
        'Company': "Company A",
        'Education': "Education A",
        'Position': "Position A",
        'College': True,
        'Highschool': True,
        'Thai': True,
        'Email': "taz2547@gmail.com"
    }
    listYay.append(filtered_row)
    send_email(listYay, "taz2547sub2@gmail.com", "enzq rtiu hgrq asox")
        
        #Change status to yes
        



    #------------------------------------------------------------------------------------------------------------------------------------------------------#
        
    # SEPARATE FUNCTIONALITY (STEP 2) (run at the end of the day)
    # NOTE: Find some way to make it run autonomously?
        #Extract email
        #Check unread emails
        #Find Specific Status (INCLUDE HERE)
            #Some kind of separate check to find out if it is a request or appointment
            #If request
                #Call google calendar API
                #Send ALL available times
            #If appointment
                #Check available times
                #And/or select an available time
                #Update google calendar api



    #------------------------------------------------------------------------------------------------------------------------------------------------------#

    driver.quit()



if __name__ == "__main__":
    main()
