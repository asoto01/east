import json
import time
import os
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from webdriver_manager.chrome import ChromeDriverManager

# --- Configuration ---
CREDENTIALS_FILE = 'credentials.json'
INITIAL_AND_TARGET_REPORTS_URL = 'https://apps.raptortech.com/Reports/Home/VisitorReports'

# --- Helper Function to Load Credentials ---
# Make sure this function definition is present and correctly placed
def load_credentials(filepath):
    """Loads username and password from a JSON file."""
    try:
        with open(filepath, 'r') as f:
            credentials = json.load(f)
            return credentials['username'], credentials['password']
    except FileNotFoundError:
        print(f"Error: Credentials file '{filepath}' not found.")
        return None, None
    except json.JSONDecodeError:
        print(f"Error: Could not decode JSON from '{filepath}'. Make sure it's valid JSON.")
        return None, None
    except KeyError:
        print(f"Error: 'username' or 'password' key missing in '{filepath}'.")
        return None, None

# --- Main Logic ---
def automate_raptor_reports():
    # This line (or around here) was line 22 in your script causing the error
    username, password = load_credentials(CREDENTIALS_FILE)
    if not username or not password:
        return

    print("Setting up WebDriver...")
    driver = None
    current_working_directory = os.getcwd()
    print(f"Files will be downloaded to the current working directory: {current_working_directory}")

    chrome_options = ChromeOptions()
    prefs = {
        "download.default_directory": current_working_directory,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    try:
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
        wait = WebDriverWait(driver, 30)

        # ... (rest of your login logic, clicks, and file handling as previously provided) ...
        print(f"Navigating to initial URL: {INITIAL_AND_TARGET_REPORTS_URL}")
        driver.get(INITIAL_AND_TARGET_REPORTS_URL)
        print(f"Current URL after initial navigation: {driver.current_url}")

        try:
            print("Checking for username field to determine if login is needed...")
            username_field_present = wait.until(EC.visibility_of_element_located((By.ID, "Username")))
            if username_field_present:
                print("Username field found. Proceeding with login steps...")
                username_field_present.send_keys(username)
                print("Username entered.")
                next_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Next']]")))
                next_button.click()
                print("'Next' button clicked.")
                password_field = wait.until(EC.visibility_of_element_located((By.ID, "Password")))
                password_field.send_keys(password)
                print("Password entered.")
                login_button = wait.until(EC.element_to_be_clickable((By.ID, "login-btn")))
                login_button.click()
                print("'Log In' button clicked.")
                print(f"URL after attempting login: {driver.current_url}")
                # !!! IMPORTANT: Update this ID if needed for your application !!!
                expected_element_after_login_id = "dashboard-welcome-message" 
                wait.until(EC.presence_of_element_located((By.ID, expected_element_after_login_id)))
                print(f"Login successful! General post-login element '{expected_element_after_login_id}' found.")
        except Exception as login_step_error:
            print(f"Login steps not performed or failed (could mean already logged in or an issue): {login_step_error}")
            print(f"Current URL is: {driver.current_url}. Assuming session might be active.")

        print(f"Ensuring navigation to target reports page: {INITIAL_AND_TARGET_REPORTS_URL}")
        driver.get(INITIAL_AND_TARGET_REPORTS_URL)
        print(f"Current URL after ensuring reports page: {driver.current_url}")
        print("Waiting for reports page content to load...")
        time.sleep(3) # Allowing time for JS elements like tabs to load

        # --- Step 1: Click "Students" tab ---
        print("Attempting to click 'Students' tab...")
        students_tab_xpath = "//a[@href='#tab3' and normalize-space(.)='Students']"
        students_tab = wait.until(EC.element_to_be_clickable((By.XPATH, students_tab_xpath)))
        students_tab.click()
        print("'Students' tab clicked.")
        time.sleep(1)

        # --- Step 2: Click "Student Sign-In/Sign-Out History" ---
        print("Attempting to click 'Student Sign-In/Sign-Out History'...")
        sign_in_out_history_xpath = "//li[contains(@class, 'item') and .//h3[normalize-space(.)='Student Sign-In/Sign-Out History']]"
        sign_in_out_history_link = wait.until(EC.element_to_be_clickable((By.XPATH, sign_in_out_history_xpath)))
        sign_in_out_history_link.click()
        print("'Student Sign-In/Sign-Out History' link clicked.")
        time.sleep(2) # Allow time for report criteria section to load

        # --- Step 3: Click "Generate Report" button ---
        print("Attempting to click 'Generate Report' button...")
        generate_report_button = wait.until(EC.element_to_be_clickable((By.ID, "generate-report")))
        generate_report_button.click()
        print("'Generate Report' button clicked.")
        print("Waiting for report to generate (e.g., 10 seconds)...") # Adjust as needed
        time.sleep(10) 

        # --- Step 4: Click "Export to Excel" ---
        print("Attempting to click 'Export to Excel' button...")
        btn_excel_export = wait.until(EC.element_to_be_clickable((By.ID, "btnExcelExport")))
        files_before_download = set(os.listdir(current_working_directory))
        print(f"Files in current working directory before export: {len(files_before_download)}")
        btn_excel_export.click()
        print("'Export to Excel' button clicked.")

        # --- Step 5: Handle downloaded file ---
        print("Waiting for download to complete (max 60 seconds)...")
        download_wait_timeout = 60
        download_poll_interval = 1
        time_waited = 0
        downloaded_file_path = None

        while time_waited < download_wait_timeout:
            files_after_download = set(os.listdir(current_working_directory))
            new_files = files_after_download - files_before_download
            
            for file_name in new_files:
                if file_name.endswith(".xlsx") and not file_name.endswith(".crdownload"):
                    print(f"New .xlsx file detected: {file_name}")
                    time.sleep(2) 
                    downloaded_file_path = os.path.join(current_working_directory, file_name)
                    break
            if downloaded_file_path:
                break
            
            time.sleep(download_poll_interval)
            time_waited += download_poll_interval
            print(f"Still waiting for download... ({time_waited}s / {download_wait_timeout}s)")

        if downloaded_file_path:
            print(f"Download complete. File saved as: {downloaded_file_path}")
            today_date_str = datetime.now().strftime("%Y-%m-%d")
            new_file_name = f"student_sign_in_{today_date_str}.xlsx"
            new_file_path = os.path.join(current_working_directory, new_file_name)

            counter = 1
            while os.path.exists(new_file_path):
                new_file_name = f"student_sign_in_{today_date_str}_{counter}.xlsx"
                new_file_path = os.path.join(current_working_directory, new_file_name)
                counter += 1
            
            os.rename(downloaded_file_path, new_file_path)
            print(f"File renamed to: {new_file_path}")
        else:
            print("Error: Download did not complete or .xlsx file not found within the timeout.")
            if driver: driver.save_screenshot('download_error_screenshot.png')

        print("Process complete. Browser will close in 10 seconds...")
        time.sleep(10)

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        if driver:
            driver.save_screenshot('unexpected_error_screenshot.png')
            print("Screenshot 'unexpected_error_screenshot.png' saved for debugging.")
    finally:
        if driver:
            print("Closing the browser.")
            driver.quit()

if __name__ == "__main__":
    automate_raptor_reports()
