from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# Path to the Chrome WebDriver executable
driver_path = 'path/to/chromedriver'  # Replace with the path to your chromedriver

# Initialize the Chrome WebDriver using Service
service = Service(driver_path)
driver = webdriver.Chrome(service=service)

# Base URL for the results page
base_url = "https://mrecresults.mrecexams.com/StudentResult/Index?Id=494&ex76brs22fbmm=2hFEFUqms4U8zZYzxE"

# Function to fetch the SGPA for a given roll number
def fetch_sgpa(roll_number):
    driver.get(base_url)
    
    try:
        # Wait for the roll number input field to be present and enter the roll number
        roll_number_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "RollNo"))
        )
        roll_number_input.send_keys(roll_number)
        
        # Submit the form or click the button to display the results
        submit_button = driver.find_element(By.ID, 'btnSubmit')  # Replace 'btnSubmit' with the actual ID if different
        submit_button.click()
        
        # Wait for the SGPA element to be visible after the result is loaded
        sgpa_element = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, f"sgpa_{roll_number}"))
        )
        sgpa = sgpa_element.text.strip()
        
        return sgpa

    except Exception as e:
        print(f"Failed to retrieve data for {roll_number}: {e}")
        return "SGPA not found"

# List of roll numbers
roll_numbers = [f"22j41s67{i:02d}" for i in range(1, 100)]  # Modify this range if needed

# Data storage
results = []

# Loop through roll numbers and collect results
for roll_number in roll_numbers:
    sgpa = fetch_sgpa(roll_number)
    results.append({'Roll Number': roll_number, 'SGPA': sgpa})
    print(f"Fetched: {roll_number} - {sgpa}")

    # Pause to avoid overloading the server
    time.sleep(1)

# Close the browser
driver.quit()

# Create a DataFrame and save to Excel
df = pd.DataFrame(results)
df.to_excel('student_results.xlsx', index=False)

print("Results have been saved to 'student_results.xlsx'")
