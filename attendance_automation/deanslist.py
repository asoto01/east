import json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

def deanslist_automation():
    """
    This script automates the process of logging into Deanslist,
    selecting students from an Excel file, and recording their tardiness.
    """

    # Load credentials from the JSON file
    with open('credentials.json', 'r') as f:
        credentials = json.load(f)

    deanslist_credentials = credentials.get('deanslist', {})
    username = deanslist_credentials.get('username')
    password = deanslist_credentials.get('password')

    if not username or not password:
        print("Deanslist credentials not found in credentials.json")
        return

    # Initialize the Chrome driver
    driver = webdriver.Chrome()
    driver.get("https://ednovate.deanslistsoftware.com/login.php?al=%2F")

    try:
        # Find the username and password fields and enter the credentials
        driver.find_element(By.NAME, "username").send_keys(username)
        driver.find_element(By.NAME, "pw").send_keys(password)

        # Click the login button
        driver.find_element(By.NAME, "submit").click()

        # Wait for the "Record Student Data" tab to be clickable
        wait = WebDriverWait(driver, 10)
        record_data_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'Record Student Data')]")))
        record_data_tab.click()

        # Click on "All Students"
        all_students = wait.until(EC.element_to_be_clickable((By.XPATH, "//td[contains(text(),'All Students')]")))
        all_students.click()

        # Click on "Behavior(s)"
        behavior = wait.until(EC.element_to_be_clickable((By.ID, "els-track-head-cont-behavior")))
        behavior.click()
        
        # Click on "Tardy to school"
        tardy_to_school = wait.until(EC.element_to_be_clickable((By.XPATH, "//td[contains(text(),'Tardy to school')]")))
        tardy_to_school.click()

        # Click on "Student(s)"
        students = wait.until(EC.element_to_be_clickable((By.ID, "els-track-head-cont-stu")))
        students.click()

        # Read the student names from the Excel file
        today_date = datetime.now().strftime('%Y-%m-%d')
        excel_file = f'daily_raptor_master_report_{today_date}.xlsx'
        
        try:
            df = pd.read_excel(excel_file)
            student_names = df['Full Name'].unique().tolist()

            # Select each student from the list
            for name in student_names:
                student_element = wait.until(EC.element_to_be_clickable((By.XPATH, f"//td[contains(text(),'{name}')]")))
                student_element.click()

        except FileNotFoundError:
            print(f"Error: The file {excel_file} was not found.")
            return
        except KeyError:
            print("Error: 'Full Name' column not found in the Excel file.")
            return

        # Prompt the user to press Enter to submit the form
        input("Press Enter to submit the form...")

        # Click the "Save" button
        save_button = wait.until(EC.element_to_be_clickable((By.ID, "els-track-save")))
        save_button.click()

        print("Form submitted successfully!")

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        # Close the browser
        driver.quit()

if __name__ == "__main__":
    deanslist_automation()

