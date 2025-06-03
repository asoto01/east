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
import re # For regular expressions to extract student count
from selenium.webdriver.support.ui import Select # Import Select for dropdowns

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
        "plugins.always_open_pdf_externally": True
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
            
            # --- NEW OPTIMIZATION: Wait for a general post-login element on the dashboard, then force navigate ---
            # Instead of looking for the 'Reports' link directly, wait for a common element on the dashboard.
            # Based on your previous description, an element from the product tiles might be a good indicator.
            # Example: A generic product tile, or a main content div on the dashboard.
            print("Waiting for RaptorTech dashboard/landing page to load (a general post-login indicator)...")
            
            # A more generic element that should be present on the dashboard after login.
            # This XPath looks for any 'product-tile' link. Adjust if your dashboard has a better, more stable indicator.
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.product-tile"))) 
            print(f"RaptorTech dashboard element detected. Current URL: {driver.current_url}")

            # Now, force a direct navigation to the reports URL.
            print(f"Forcing direct navigation to target reports page: {INITIAL_AND_TARGET_REPORTS_URL}")
            driver.get(INITIAL_AND_TARGET_REPORTS_URL)
            
            # Finally, confirm that we've landed on the correct reports URL.
            wait.until(EC.url_to_be(INITIAL_AND_TARGET_REPORTS_URL))
            print(f"Successfully landed on RaptorTech reports page: {driver.current_url}")
            
        except Exception as login_step_error:
            print(f"RaptorTech login steps not performed or failed (could mean already logged in or an issue): {login_step_error}")
            print(f"Current URL is: {driver.current_url}. Attempting to navigate directly to reports URL if not there.")
            if driver.current_url != INITIAL_AND_TARGET_REPORTS_URL:
                 driver.get(INITIAL_AND_TARGET_REPORTS_URL)
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
        wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@id='tab3']//h3[normalize-space(.)='Student Sign-In/Sign-Out History']")))


        # --- Step 2: Click "Student Sign-In/Sign-Out History" ---
        print("Attempting to click 'Student Sign-In/Sign-Out History'...")
        sign_in_out_history_xpath = "//li[contains(@class, 'item') and .//h3[normalize-space(.)='Student Sign-In/Sign-Out History']]"
        sign_in_out_history_link = wait.until(EC.element_to_be_clickable((By.XPATH, sign_in_out_history_xpath)))
        sign_in_out_history_link.click()
        print("'Student Sign-In/Sign-Out History' link clicked.")
        wait.until(EC.presence_of_element_located((By.ID, "generate-report")))


        # --- Step 3: Click "Generate Report" button ---
        print("Attempting to click 'Generate Report' button...")
        generate_report_button = wait.until(EC.element_to_be_clickable((By.ID, "generate-report")))
        generate_report_button.click()
        print("'Generate Report' button clicked.")
        
        # --- Optimized Wait for Report Generation (1 second at most) ---
        print("Waiting for RaptorTech report data to load (looking for 'btnExcelExport' to be clickable, max 1 sec)...")
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
                if file_name.endswith(".xlsx") and not file_name.endswith(".crdownload"):
                    print(f"New .xlsx file detected: {file_name}")
                    time.sleep(1) # Give it a tiny bit more time to ensure it's fully written
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
            timestamp_str = datetime.now().strftime("%m-%d-%Y-%I-%M-%S-%p") # Added seconds
            new_file_name_base = f"student-sign-in-{timestamp_str}"
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
            ids_to_paste = "" # Initialize here to ensure it's always defined
            try:
                df = pd.read_excel(downloaded_excel_file_path)
                
                if 'Date/Time' in df.columns:
                    df['Date/Time'] = pd.to_datetime(df['Date/Time'], errors='coerce')
                    
                    start_time = datetime.strptime("08:30 AM", "%I:%M %p").time()
                    end_time = datetime.strptime("09:05 AM", "%I:%M %p").time()

                    filtered_df = df[
                        (df['Date/Time'].dt.time >= start_time) & 
                        (df['Date/Time'].dt.time <= end_time)
                    ]
                    
                    if 'ID Number' in filtered_df.columns:
                        extracted_ids = filtered_df['ID Number'].dropna().astype(str).tolist()
                        ids_to_paste = "\n".join(extracted_ids)
                        print(f"Extracted {len(extracted_ids)} ID(s) from Excel file. IDs are:\n{ids_to_paste}")
                    else:
                        print("Error: 'ID Number' column not found in the Excel file.")
                else:
                    print("Error: 'Date/Time' column not found in the Excel file for filtering.")

            except Exception as excel_error:
                print(f"Error processing Excel file: {excel_error}")

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
                
                password_field_ps.submit() # Submit the form
                print("Attempted to log in to PowerSchool.")

                # Wait for PowerSchool dashboard to load
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
                print(f"IDs pasted into PowerSchool MultiSelect textarea. Pasted IDs:\n{ids_to_paste}") # Confirm pasted IDs

                # --- Click "Search" button in MultiSelect dialog ---
                print("Attempting to click 'Search' button in MultiSelect dialog...")
                search_button_xpath = "//button[contains(., 'Search') and @onclick=\"MultiSelect.searchType='admin'; MultiSelect.powerScheduler = 'Home'; MultiSelect.collectIDs();\"]"
                search_button = wait.until(EC.element_to_be_clickable((By.XPATH, search_button_xpath)))
                search_button.click()
                print("'Search' button clicked in MultiSelect dialog.")
                
                # --- Wait for student count and print it ---
                print("Waiting for student selection count to appear...")
                # Adjusting to a small fixed wait after visibility to ensure text is stable
                student_count_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//h2[contains(text(), 'Current Student Selection')]")))
                time.sleep(0.5) # Small buffer for text rendering
                
                count_text = student_count_element.text
                match = re.search(r'\((\d+)\)', count_text)
                if match:
                    student_count = match.group(1)
                    print(f"Number of students selected: {student_count}")
                else:
                    print("Could not extract student count from the element text.")
                
                # --- Click the correct dropdown button for "Group Functions" ---
                print("Attempting to click 'Group Functions' dropdown button...")
                group_functions_button = wait.until(EC.element_to_be_clickable((By.ID, "selectFunctionDropdownButtonStudent")))
                group_functions_button.click()
                print("'Group Functions' dropdown button clicked.")
                
                # Add a small explicit wait for the dropdown menu items to become visible
                # This is important before trying to click 'Mass Update Attendance'
                time.sleep(0.5) 

                # --- Click "Mass Update Attendance" link ---
                print("Attempting to click 'Mass Update Attendance' link...")
                mass_update_attendance_link = wait.until(EC.element_to_be_clickable((By.ID, "lnk_studentsMassUpdateAttendance")))
                mass_update_attendance_link.click()
                print("'Mass Update Attendance' link clicked.")

                # --- Wait for redirection to batch update page ---
                print("Waiting for redirection to batch attendance update page...")
                wait.until(EC.url_to_be("https://ednovate.powerschool.com/admin/attendance/record/batch/meetinggroup.html?dothisfor=selected"))
                print(f"Successfully redirected to: {driver.current_url}")

                # --- Select checkbox for "AMA" ---
                print("Attempting to select 'AMA' checkbox...")
                ama_checkbox_xpath = "//tr[td[normalize-space(.)='AMA']]/td/input[@type='checkbox' and @name='cb7;1']"
                ama_checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, ama_checkbox_xpath)))
                if not ama_checkbox.is_selected():
                    ama_checkbox.click()
                    print("'AMA' checkbox selected.")
                else:
                    print("'AMA' checkbox was already selected.")

                # --- Select checkbox for "AMAB" ---
                print("Attempting to select 'AMAB' checkbox...")
                amab_checkbox_xpath = "//tr[td[normalize-space(.)='AMA']]/td/input[@type='checkbox' and @name='cb7;2']"
                amab_checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, amab_checkbox_xpath)))
                if not amab_checkbox.is_selected():
                    amab_checkbox.click()
                    print("'AMAB' checkbox selected.")
                else:
                    print("'AMAB' checkbox was already selected.")
                
                # --- Select "UL" for attendance code ---
                print("Attempting to select 'UL' as attendance code...")
                attendance_code_select = wait.until(EC.element_to_be_clickable((By.NAME, "att_attcodelist")))
                select = Select(attendance_code_select)
                select.select_by_value("UL")
                print("'UL' (Unexcused Late Arrival) selected for attendance code.")

                print("PowerSchool batch update configuration complete. Browser will stay open for 10 seconds for review.")
                time.sleep(10) # Keep browser open for review before closing

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
