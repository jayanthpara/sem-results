import os
import string  # Import the string module
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

# Set up paths (replace with your actual paths)
brave_path = r'C:\Users\91934\AppData\Local\BraveSoftware\Brave-Browser\Application\brave.exe'
driver_path = r'C:\Users\91934\Downloads\Results automate\chromedriver-win64\chromedriver.exe'
excel_file = 'student_results.xlsx'

# WebDriver setup
chrome_options = Options()
chrome_options.binary_location = brave_path
service = Service(driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Open results page
print("Opening results page...")
driver.get("https://mrecresults.mrecexams.com/StudentResult/Index?Id=494&ex76brs22fbmm=2hFEFUqms4U8zZYzxE")  #update the link for session control

# Generate roll numbers
def generate_roll_numbers():
    roll_numbers = []
    prefixes = ['j41a67']

    # Generate numbers from 01 to 99
    for i in range(1, 100):
        roll_number = f"23{prefixes[0]}{i:02}"
        roll_numbers.append(roll_number)

    # Generate alphanumeric roll numbers a0 to k3
    for letter in string.ascii_lowercase[0:11]:  # Assuming a0 to k3 corresponds to a to k
        for digit in range(0, 10):  # 0 to 9
            roll_number = f"23{prefixes[0]}{letter}{digit}"
            roll_numbers.append(roll_number)

    return roll_numbers

roll_numbers = generate_roll_numbers()

# Check if the Excel file exists, create if not
if not os.path.exists(excel_file):
    df = pd.DataFrame(columns=['Roll Number', 'Name', 'SGPA', 'CGPA'] + [f'Subject {i+1}' for i in range(20)])  # Adjust for 20 subjects (CIE + SIE)
    df.to_excel(excel_file, index=False)

for roll_number in roll_numbers:
    try:
        print(f"Entering roll number: {roll_number}")

        input_field = driver.find_element(By.NAME, "HallTicketNo")
        input_field.clear()
        input_field.send_keys(roll_number)
        input_field.send_keys(Keys.RETURN)

        print("Waiting for results to load...")

        # Wait for SGPA element to be present
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f"//div[@id='sgpa_{roll_number.upper()}']")))

        # Fetch SGPA and CGPA
        sgpa_element = driver.find_element(By.XPATH, f"//div[@id='sgpa_{roll_number.upper()}']")
        sgpa = sgpa_element.text.strip()
        print(f"Found SGPA for {roll_number}: {sgpa}")

        cgpa_element = driver.find_element(By.XPATH, f"//td[@id='cgpa_{roll_number.upper()}']")
        cgpa = cgpa_element.text.strip()
        print(f"Found CGPA for {roll_number}: {cgpa}")

        # Fetch Name
        name_element = driver.find_element(By.XPATH, "(//span[@style='color:#851fd0; font-weight:bold'])[2]")
        name = name_element.text.strip().title()  # Capitalize first letter of each word
        print(f"Found Name for {roll_number}: {name}")

        # Extract subject-wise CIE and SIE marks
        marks = []
        subject_elements = driver.find_elements(
            By.XPATH, "//td[contains(@class, 'cie-mark')] | //td[contains(@class, 'sie-mark')]"
        )
        for element in subject_elements:
            marks.append(element.text.strip())

        # Prepare data for Excel
        result = pd.DataFrame({'Roll Number': [roll_number], 'Name': [name], 'SGPA': [sgpa], 'CGPA': [cgpa]})
        for i in range(len(marks)):
            result[f'Subject {i+1}'] = marks[i]

        print(f"Data for {roll_number} entered successfully.")

    except Exception as e:
        print(f"Result not found for {roll_number}: {e}")
        name = 'Not Found'
        sgpa = 'Not Found'
        cgpa = 'Not Found'
        marks = ['Not Found'] * 20  # Assuming 20 subjects
        result = pd.DataFrame({'Roll Number': [roll_number], 'Name': [name], 'SGPA': [sgpa], 'CGPA': [cgpa]})

        for i in range(20):
            result[f'Subject {i+1}'] = marks[i]
        print(driver.page_source)  # Print the page source for debugging

    # Append result to Excel immediately after fetching
    existing_df = pd.read_excel(excel_file)
    updated_df = pd.concat([existing_df, result], ignore_index=True)
    updated_df.to_excel(excel_file, index=False)
    print(f"Results saved for roll number: {roll_number}")

# Close the browser
driver.quit()
print("Browser closed.")
