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

from webdriver_manager.chrome import ChromeDriverManager

from bs4 import BeautifulSoup
import time


def get_linkedin_profile_info(driver, profile_url):
    try:
        # Get name
        print("Name:")
        driver.get(profile_url)
        name_xpath = "//h1[contains(@class,'text-heading-xlarge')]"
        name_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, name_xpath))
        )
        print(name_element.text)
        print()
        print()


        #------------------------------------------------------------------------------------------------------------------------------------------------------#

        # Get email, incorporate leadleaper plugin, find another solution in the future


        #------------------------------------------------------------------------------------------------------------------------------------------------------#
        
        # Get Professional Experiences, will have to check if aligns with what you want, have to make sure they are CURRENTLY working here
        print("Experiences:")
        driver.get(profile_url + 'details/experience/')
        # Wait for the experience section to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//div[@data-view-name='profile-component-entity']"))
        )

        # Select all the top-level experience entries
        experience_entries_xpath = "//div[@data-view-name='profile-component-entity' and not(ancestor::div[@data-view-name='profile-component-entity'])]"
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

            print()
            print()


        #------------------------------------------------------------------------------------------------------------------------------------------------------#
            
        # Get Education
        print("Education:")
        driver.get(profile_url + 'details/education/')
        education_entries_xpath = "//div[@data-view-name='profile-component-entity']"
        education_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, education_entries_xpath))
        )

        for element in education_elements:
            school_name_element = element.find_element(By.XPATH, ".//span[not(contains(@class, 'visually-hidden'))]")
            school_name = school_name_element.get_attribute('innerText').strip()
            print(school_name)
            if "University of Illinois Urbana-Champaign" in school_name:
                print("Found matching education:", school_name)

        print()
        print()


    except Exception as e:
        print(f"An error occurred: {e}")


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
    my_education = "University of Illinois Urbana-Champaign" # Eventually will need to switch this with array
    desired_job = ""

    # LinkedIn profile URLs
    profile_urls = []
    profile_urls.append('https://www.linkedin.com/in/edward-yo-kang/')
    profile_urls.append('https://www.linkedin.com/in/edkang99/')
    profile_urls.append('https://www.linkedin.com/in/kritamet-bu/')
    # profile_urls.append('https://www.linkedin.com/in/taz-vongpatarakul-66b1a6214/')
    # profile_urls.append('https://www.linkedin.com/in/paphada-rungsinaporn-95b2a5248/')



    #------------------------------------------------------------------------------------------------------------------------------------------------------#

    # Initialize Chrome options and Driver
    chrome_options = Options()
    # LeadLeaper Extension
    # chrome_options.add_argument('--load-extension=/Users/tazvongpatarakul/Desktop/CodingStuff/LeadLeaper_Extension/7.1.20_0')
    # Uncomment the next line if you want Chrome to run headless
    # chrome_options.add_argument("--headless")

    # Initialize WebDriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    # Login to LeadLeaper
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
    username.send_keys('taz2547sub2@gmail.com')
    password = driver.find_element(By.ID, 'password')
    password.send_keys('123123123123')

    # Submit the form
    password.send_keys(Keys.RETURN)

    # Can have time delay for human verification, until there is a better solution
    # time.sleep(60)

    # Wait for login to complete and navigate to the profile page
    time.sleep(1)


    #------------------------------------------------------------------------------------------------------------------------------------------------------#

    get_linkedin_profile_info(driver, profile_urls[0])
    get_linkedin_profile_info(driver, profile_urls[1])
    get_linkedin_profile_info(driver, profile_urls[2])
    
    driver.quit()



if __name__ == "__main__":
    main()
