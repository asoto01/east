import json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time
import os

def deanslist_automation():
    """
    This script automates logging into Deanslist, selecting students from an
    Excel file, and recording their tardiness.
    """
    # --- Load Credentials ---
    try:
        with open('credentials.json', 'r') as f:
            credentials = json.load(f)
        deanslist_credentials = credentials.get('deanslist', {})
        username = deanslist_credentials.get('username')
        password = deanslist_credentials.get('password')

        if not username or not password:
            print("❌ Deanslist username or password not found in credentials.json")
            return
    except FileNotFoundError:
        print("❌ Error: credentials.json file not found.")
        return

    # --- Initialize WebDriver ---
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')  # Uncomment to run without a browser window
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)

    try:
        # --- Login ---
        print("🚀 Starting Deanslist automation...")
        driver.get("https://ednovate.deanslistsoftware.com/login.php?al=%2F")
        driver.maximize_window()
        
        print("🔑 Logging in...")
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(username)
        driver.find_element(By.NAME, "pw").send_keys(password)
        driver.find_element(By.NAME, "submit").click()
        print("✅ Login successful!")

        # --- Navigate the Menu ---
        # 1. Click "Record Student Data" tab

# --- OPTION 1: Find by Text (Most Recommended) ---
# This is often the most reliable method. It finds the div with the class 'nav-tab'
# that also contains the exact text "Record Student Data".
# The `normalize-space()` function handles any extra whitespace in the HTML.
        try:
            record_data_tab = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'nav-tab') and normalize-space()='Record Student Data']")))
        record_data_tab.click()
        print("Successfully clicked the tab using text content.")

            
        except Exception as e:
        print(f"Error with Option 1 (Find by Text): {e}")

        print(" navigating to 'Record Student Data'...")
        record_data_tab = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'nav-tab') and .//i[contains(@class, 'fa-cubes')]]"))
        )
        record_data_tab.click()

        # 2. Click "All Students"
        print(" selecting 'All Students'...")
        all_students_roster = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//td[@class='click' and normalize-space()='All Students']"))
        )
        all_students_roster.click()
        
        # 3. Click "Behavior(s)" section
        print(" selecting 'Behavior(s)'...")
        behavior_section = wait.until(
            EC.element_to_be_clickable((By.ID, "els-track-head-cont-behavior"))
        )
        behavior_section.click()

        # 4. Click "Tardy to school"
        print(" selecting 'Tardy to school'...")
        tardy_to_school = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//td[contains(@class, 'cp-behavior') and normalize-space()='Tardy to school']"))
        )
        tardy_to_school.click()

        # 5. Click back to the "Student(s)" section to focus the student list
        print(" focusing on the 'Student(s)' list...")
        students_section = wait.until(
            EC.element_to_be_clickable((By.ID, "els-track-head-cont-stu"))
        )
        students_section.click()

        # --- Process Students from Excel ---
        today_date = datetime.now().strftime('%Y-%m-%d')
        excel_file = f'daily_raptor_report_master/daily_raptor_master_report_{today_date}.xlsx'
        
        if not os.path.exists(excel_file):
            print(f"❌ Error: The file {excel_file} was not found.")
            driver.quit()
            return
            
        print(f"📄 Reading student names from {excel_file}...")
        df = pd.read_excel(excel_file)
        
        if 'Full Name' not in df.columns:
            print("❌ Error: 'Full Name' column not found in the Excel file.")
            driver.quit()
            return
            
        student_names = [name for name in df['Full Name'].unique() if pd.notna(name) and str(name).strip()]
        print(f"找到了 {len(student_names)} unique students to process.")

        not_found_students = []
        for name in student_names:
            try:
                # Use a specific XPath to find the student by their exact name
                student_element = wait.until(
                    EC.element_to_be_clickable((By.XPATH, f"//td[contains(@class, 'click') and normalize-space()='{name.strip()}']"))
                )
                student_element.click()
                print(f"  - Selected: {name}")
            except Exception:
                print(f"  - ⚠️ Could not find '{name}' in the list. Skipping.")
                not_found_students.append(name)

        if not_found_students:
            print("\n--- The following students were not found ---")
            for name in not_found_students:
                print(f"  - {name}")
            print("------------------------------------------")

        # --- Final User Confirmation ---
        while True:
            user_input = input("\nAll found students have been selected. Press 'Enter' to save and submit, or type 'q' to quit: ").lower()
            if user_input == 'q':
                print("🛑 Submission cancelled by user.")
                return
            if user_input == '':
                break
        
        # --- Save the Form ---
        print("💾 Saving the form...")
        save_button = wait.until(EC.element_to_be_clickable((By.ID, "els-track-save")))
        save_button.click()

        print("🎉 Form submitted successfully!")
        time.sleep(3) # Wait to observe the result

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    
    finally:
        print("Closing the browser.")
        driver.quit()

if __name__ == "__main__":
    deanslist_automation()
