import json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time

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
    options = webdriver.ChromeOptions()
    # Uncomment the next line to run in headless mode (without a visible browser window)
    # options.add_argument('--headless')
    driver = webdriver.Chrome(options=options)
    driver.get("https://ednovate.deanslistsoftware.com/login.php?al=%2F")
    driver.maximize_window()

    try:
        wait = WebDriverWait(driver, 20)
        # Find the username and password fields and enter the credentials
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(username)
        driver.find_element(By.NAME, "pw").send_keys(password)

        # Click the login button
        driver.find_element(By.NAME, "submit").click()

        # ---- MODIFIED SECTION START ----
        # Wait for the "Record Student Data" tab to be clickable and click it to open the side panel
        print("Waiting for 'Record Student Data' tab...")
        # Using a more direct XPath and waiting for the element to be present
        record_data_tab = wait.until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'nav-tab') and contains(normalize-space(), 'Record Student Data')]"))
        )
        # Use a JavaScript click, which can be more reliable for elements that are tricky to interact with.
        driver.execute_script("arguments[0].click();", record_data_tab)
        print("'Record Student Data' tab clicked.")

        # Wait for the side panel to open and then click on the "Roster" which is labeled "All Students"
        print("Waiting for 'All Students' roster...")
        all_students_roster = wait.until(EC.element_to_be_clickable((By.ID, "els-track-head-cont-roster")))
        all_students_roster.click()
        print("'All Students' roster clicked.")
        # ---- MODIFIED SECTION END ----

        # Click on "Behavior(s)"
        print("Waiting for 'Behavior(s)' section...")
        behavior = wait.until(EC.element_to_be_clickable((By.ID, "els-track-head-cont-behavior")))
        behavior.click()
        print("'Behavior(s)' section clicked.")
        
        # Click on "Tardy to school"
        print("Waiting for 'Tardy to school' behavior...")
        tardy_to_school = wait.until(EC.element_to_be_clickable((By.XPATH, "//td[contains(@class, 'cp-behavior') and contains(., 'Tardy to school')]")))
        tardy_to_school.click()
        print("'Tardy to school' clicked.")

        # Click on "Student(s)" to bring focus back to the student selection panel
        print("Waiting for 'Student(s)' section...")
        students_section = wait.until(EC.element_to_be_clickable((By.ID, "els-track-head-cont-stu")))
        students_section.click()
        print("'Student(s)' section clicked.")

        # Read the student names from the Excel file
        today_date = datetime.now().strftime('%Y-%m-%d')
        # Use a flexible path to find the file
        excel_file = f'daily_raptor_master_report_{today_date}.xlsx'
        
        try:
            print(f"Reading student names from {excel_file}...")
            df = pd.read_excel(excel_file)
            # Get unique, non-empty student names
            student_names = [name for name in df['Full Name'].unique() if pd.notna(name) and str(name).strip()]
            print(f"Found {len(student_names)} unique students.")

            # Select each student from the list
            for name in student_names:
                try:
                    print(f"Selecting student: {name}")
                    # Use a more specific XPath to find the student by their exact name
                    student_element = wait.until(EC.element_to_be_clickable((By.XPATH, f"//td[contains(@class, 'click') and normalize-space()='{name.strip()}']")))
                    student_element.click()
                except Exception as student_error:
                    print(f"Could not select '{name}'. They might not be in the list or the name is slightly different. Error: {student_error}")
                    continue

        except FileNotFoundError:
            print(f"Error: The file {excel_file} was not found.")
            return
        except KeyError:
            print("Error: 'Full Name' column not found in the Excel file.")
            return

        # Prompt the user to press Enter to submit the form
        input("All students selected. Press Enter to save and submit the form...")

        # Click the "Save" button
        save_button = wait.until(EC.element_to_be_clickable((By.ID, "els-track-save")))
        save_button.click()

        print("Form submitted successfully!")
        # Add a small delay to see the confirmation before the browser closes
        time.sleep(3)

    except Exception as e:
        print(f"An unexpected error occurred: {e}")

    finally:
        # Close the browser
        print("Closing the browser.")
        driver.quit()

if __name__ == "__main__":
    deanslist_automation()
