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
import pandas as pd
import glob # To easily find the latest downloaded file

# --- Configuration ---
CREDENTIALS_FILE = 'credentials.json'
INITIAL_AND_TARGET_REPORTS_URL = 'https://apps.raptortech.com/Reports/Home/VisitorReports'
POWERSCHOOL_LOGIN_URL = 'https://ednovate.powerschool.com/admin/pw.html'

# --- Helper Function to Load Credentials ---
def load_credentials(filepath, app_name):
    """Loads username and password for a specific application from a JSON file."""
    try:
        with open(filepath, 'r') as f:
            all_credentials = json.load(f)
            app_credentials = all_credentials.get(app_name)
            if app_credentials:
                return app_credentials.get('username'), app_credentials.get('password')
            else:
                print(f"Error: Application '{app_name}' credentials not found in '{filepath}'.")
                return None, None
    except FileNotFoundError:
        print(f"Error: Credentials file '{filepath}' not found.")
        return None, None
    except json.JSONDecodeError:
        print(f"Error: Could not decode JSON from '{filepath}'. Make sure it's valid JSON.")
        return None, None
    except Exception as e:
        print(f"An unexpected error occurred while loading credentials: {e}")
        return None, None

# --- Main Logic ---
def automate_raptor_and_powerschool():
    # Load RaptorTech credentials
    raptor_username, raptor_password = load_credentials(CREDENTIALS_FILE, 'raptor')
    if not raptor_username or not raptor_password:
        return

    # Load PowerSchool credentials
    powerschool_username, powerschool_password = load_credentials(CREDENTIALS_FILE, 'powerschool')
    if not powerschool_username or not powerschool_password:
        return

    print("Setting up WebDriver...")
    driver = None
    current_working_directory = os.getcwd()
    print(f"Files will be downloaded to the current working directory: {current_working_directory}")

    chrome_options = ChromeOptions()
    # Uncomment the line below to run headless (no browser UI)
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox") # Required for some Linux environments
    chrome_options.add_argument("--window-size=1920,1080") # Set a consistent window size for reliability
    chrome_options.add_argument("--start-maximized") # Maximize window to ensure elements are visible

    prefs = {
        "download.default_directory": current_working_directory,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True # Important for PDFs, though not directly used here
    }
    chrome_options.add_experimental_option("prefs", prefs)

    downloaded_excel_file_path = None # To store the path of the downloaded file

    try:
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
        wait = WebDriverWait(driver, 30)

        # --- RaptorTech Automation ---
        print(f"Navigating to initial URL: {INITIAL_AND_TARGET_REPORTS_URL}")
        driver.get(INITIAL_AND_TARGET_REPORTS_URL)
        print(f"Current URL after initial navigation: {driver.current_url}")

        try:
            print("Checking for username field to determine if login is needed for RaptorTech...")
            username_field = wait.until(EC.visibility_of_element_located((By.ID, "Username")))
            
            print("Username field found. Proceeding with RaptorTech login steps...")
            username_field.send_keys(raptor_username)
            print("RaptorTech Username entered.")
            
            next_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Next']]")))
            next_button.click()
            print("'Next' button clicked.")
            
            password_field = wait.until(EC.visibility_of_element_located((By.ID, "Password")))
            password_field.send_keys(raptor_password)
            print("RaptorTech Password entered.")
            
            login_button = wait.until(EC.element_to_be_clickable((By.ID, "login-btn")))
            login_button.click()
            print("'Log In' button clicked.")
            
            # --- Optimized Wait for RaptorTech Landing Page ---
            print(f"Waiting for RaptorTech to redirect to: {INITIAL_AND_TARGET_REPORTS_URL}...")
            # This waits for the URL to specifically match the target reports URL
            wait.until(EC.url_to_be(INITIAL_AND_TARGET_REPORTS_URL))
            print(f"Successfully landed on RaptorTech reports page: {driver.current_url}")
            
        except Exception as login_step_error:
            print(f"RaptorTech login steps not performed or failed (could mean already logged in or an issue): {login_step_error}")
            print(f"Current URL is: {driver.current_url}. Assuming RaptorTech session might be active.")
            # If we are not on the target URL, navigate there directly
            if driver.current_url != INITIAL_AND_TARGET_REPORTS_URL:
                 print(f"Navigating to target RaptorTech reports page: {INITIAL_AND_TARGET_REPORTS_URL}")
                 driver.get(INITIAL_AND_TARGET_REPORTS_URL)
                 # Wait for the URL to change to the target page after direct navigation
                 wait.until(EC.url_to_be(INITIAL_AND_TARGET_REPORTS_URL))
                 print(f"Current URL after ensuring reports page: {driver.current_url}")

        print("Waiting for RaptorTech reports page content to load (e.g., tabs container)...")
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "nav-tabs")))


        # --- Step 1: Click "Students" tab ---
        print("Attempting to click 'Students' tab...")
        students_tab_xpath = "//a[@href='#tab3' and normalize-space(.)='Students']"
        students_tab = wait.until(EC.element_to_be_clickable((By.XPATH, students_tab_xpath)))
        students_tab.click()
        print("'Students' tab clicked.")
        # Wait for an element inside the Students tab to be visible after click
        wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@id='tab3']//h3[normalize-space(.)='Student Sign-In/Sign-Out History']")))


        # --- Step 2: Click "Student Sign-In/Sign-Out History" ---
        print("Attempting to click 'Student Sign-In/Sign-Out History'...")
        sign_in_out_history_xpath = "//li[contains(@class, 'item') and .//h3[normalize-space(.)='Student Sign-In/Sign-Out History']]"
        sign_in_out_history_link = wait.until(EC.element_to_be_clickable((By.XPATH, sign_in_out_history_xpath)))
        sign_in_out_history_link.click()
        print("'Student Sign-In/Sign-Out History' link clicked.")
        # Wait for the report criteria section to load, indicated by 'generate-report' button
        wait.until(EC.presence_of_element_located((By.ID, "generate-report")))


        # --- Step 3: Click "Generate Report" button ---
        print("Attempting to click 'Generate Report' button...")
        generate_report_button = wait.until(EC.element_to_be_clickable((By.ID, "generate-report")))
        generate_report_button.click()
        print("'Generate Report' button clicked.")
        
        # --- Optimized Wait for Report Generation (assuming btnExcelExport becomes clickable) ---
        print("Waiting for report data to load (looking for 'btnExcelExport' to be clickable)...")
        # This is the key optimization for report generation time
        wait.until(EC.element_to_be_clickable((By.ID, "btnExcelExport")))
        print("RaptorTech report appears to have generated.")


        # --- Step 4: Click "Export to Excel" ---
        print("Attempting to click 'Export to Excel' button...")
        btn_excel_export = wait.until(EC.element_to_be_clickable((By.ID, "btnExcelExport")))
        files_before_download = set(os.listdir(current_working_directory))
        print(f"Files in current working directory before export: {len(files_before_download)}")
        btn_excel_export.click()
        print("'Export to Excel' button clicked.")

        # --- Step 5: Handle downloaded file ---
        print("Waiting for RaptorTech download to complete (max 60 seconds)...")
        download_wait_timeout = 60
        download_poll_interval = 1
        time_waited = 0

        while time_waited < download_wait_timeout:
            files_after_download = set(os.listdir(current_working_directory))
            new_files = files_after_download - files_before_download
            
            for file_name in new_files:
                # Check for .xlsx and ensure it's not a partial download
                if file_name.endswith(".xlsx") and not file_name.endswith(".crdownload"):
                    print(f"New .xlsx file detected: {file_name}")
                    # Give it a tiny bit more time to ensure it's fully written
                    time.sleep(1) 
                    downloaded_excel_file_path = os.path.join(current_working_directory, file_name)
                    break
            if downloaded_excel_file_path:
                break
            
            time.sleep(download_poll_interval)
            time_waited += download_poll_interval
            print(f"Still waiting for download... ({time_waited}s / {download_wait_timeout}s)")

        if downloaded_excel_file_path:
            print(f"Download complete. File saved as: {downloaded_excel_file_path}")
            
            # --- Consistent File Renaming with Timestamp (using dashes) ---
            # %I for 12-hour clock, %M for minute, %p for AM/PM
            today_date_str = datetime.now().strftime("%m-%d-%Y-%I-%M-%p") 
            new_file_name_base = f"student-sign-in-{today_date_str}"
            new_file_name = f"{new_file_name_base}.xlsx"
            new_file_path = os.path.join(current_working_directory, new_file_name)

            counter = 1
            while os.path.exists(new_file_path):
                new_file_name = f"{new_file_name_base}-{counter}.xlsx"
                new_file_path = os.path.join(current_working_directory, new_file_name)
                counter += 1
                
            os.rename(downloaded_excel_file_path, new_file_path)
            downloaded_excel_file_path = new_file_path # Update the path to the renamed file
            print(f"File renamed to: {downloaded_excel_file_path}")

            # --- Data Extraction from Excel ---
            print("Extracting ID Numbers from the downloaded Excel file...")
            try:
                df = pd.read_excel(downloaded_excel_file_path)
                
                # Ensure 'Date/Time' column is datetime objects for proper filtering
                if 'Date/Time' in df.columns:
                    df['Date/Time'] = pd.to_datetime(df['Date/Time'], errors='coerce')
                    
                    # Filter for times between 8:30 AM and 9:05 AM
                    # Create time objects for comparison
                    start_time = datetime.strptime("08:30 AM", "%I:%M %p").time()
                    end_time = datetime.strptime("09:05 AM", "%I:%M %p").time()

                    # Filter rows where the time component of 'Date/Time' is within the range
                    filtered_df = df[
                        (df['Date/Time'].dt.time >= start_time) & 
                        (df['Date/Time'].dt.time <= end_time)
                    ]
                    
                    # Extract 'ID Number' column, drop NaNs, and convert to list of strings
                    if 'ID Number' in filtered_df.columns:
                        extracted_ids = filtered_df['ID Number'].dropna().astype(str).tolist()
                        print(f"Extracted {len(extracted_ids)} ID(s) from Excel file.")
                        # Format as a single string, one ID per line
                        ids_to_paste = "\n".join(extracted_ids)
                        print(f"IDs to paste into PowerSchool:\n{ids_to_paste}")
                    else:
                        print("Error: 'ID Number' column not found in the Excel file.")
                        ids_to_paste = ""
                else:
                    print("Error: 'Date/Time' column not found in the Excel file for filtering.")
                    ids_to_paste = ""

            except Exception as excel_error:
                print(f"Error processing Excel file: {excel_error}")
                ids_to_paste = ""

            # --- PowerSchool Automation ---
            if ids_to_paste:
                print(f"Navigating to PowerSchool login page: {POWERSCHOOL_LOGIN_URL}")
                driver.get(POWERSCHOOL_LOGIN_URL)
                
                print("Attempting PowerSchool login...")
                username_field_ps = wait.until(EC.visibility_of_element_located((By.ID, "fieldUsername")))
                password_field_ps = wait.until(EC.visibility_of_element_located((By.ID, "fieldPassword")))
                
                username_field_ps.send_keys(powerschool_username)
                password_field_ps.send_keys(powerschool_password)
                print("PowerSchool credentials entered.")
                
                # Assuming the login button is typically the next clickable element or a form submission
                # If there's no explicit login button, the form might submit on Enter key press.
                # Let's assume there's a submit button or pressing Enter works.
                # A common ID for a login button is "btnSubmit" or similar. If not, we can trigger form submit.
                # For now, let's try to submit the form after entering password.
                password_field_ps.submit() # This attempts to submit the form after typing password
                print("Attempted to log in to PowerSchool.")

                # Wait for the URL to change or a specific element on the PowerSchool dashboard/landing page
                print("Waiting for PowerSchool dashboard to load...")
                # You'll need to find a unique element ID/XPATH that appears ONLY after successful login.
                # For demonstration, I'll wait for the "MultiSelect" link itself, but a dashboard element is better.
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.dialogDivM.custom_link[title='MultiSelect - Students']")))
                print("PowerSchool login successful.")

                # --- Click MultiSelect Link ---
                print("Attempting to click 'MultiSelect - Students' link...")
                multiselect_link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.dialogDivM.custom_link[title='MultiSelect - Students']")))
                multiselect_link.click()
                print("'MultiSelect - Students' link clicked.")

                # --- Paste IDs into Textarea ---
                print("Waiting for MultiSelect dialog to appear and textarea to be visible...")
                multiselect_textarea = wait.until(EC.visibility_of_element_located((By.ID, "multiSelValsStu")))
                multiselect_textarea.clear() # Clear any existing content
                multiselect_textarea.send_keys(ids_to_paste)
                print("IDs pasted into PowerSchool MultiSelect textarea.")
                
                print("PowerSchool process completed. You can now manually proceed in PowerSchool if needed.")
                time.sleep(5) # Keep browser open briefly for manual inspection

            else:
                print("No IDs extracted, skipping PowerSchool automation.")

        else:
            print("Error: RaptorTech download did not complete or .xlsx file not found within the timeout.")
            if driver: driver.save_screenshot('download_error_screenshot.png')

        print("Automation process complete. Browser will close in 5 seconds...")
        time.sleep(5)

    except Exception as e:
        print(f"An unexpected error occurred during automation: {e}")
        if driver:
            driver.save_screenshot('unexpected_error_screenshot.png')
            print("Screenshot 'unexpected_error_screenshot.png' saved for debugging.")
    finally:
        if driver:
            print("Closing the browser.")
            driver.quit()

if __name__ == "__main__":
    automate_raptor_and_powerschool()
