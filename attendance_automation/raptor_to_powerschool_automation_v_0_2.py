import json
import time
import os
import sys 
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import re
from selenium.webdriver.support.ui import Select

# --- Configuration ---
CREDENTIALS_FILE = 'credentials.json'
INITIAL_AND_TARGET_REPORTS_URL = 'https://apps.raptortech.com/Reports/Home/VisitorReports'
POWERSCHOOL_LOGIN_URL = 'https://ednovate.powerschool.com/admin/pw.html'

# --- School Period Definitions ---
# Normal Monday-Thursday Schedule
M_TH_NORMAL_PERIODS = [
    {"id": "AMA", "name": "AM Advisory (Normal)", "start_str": "08:30 AM", "end_str": "09:05 AM", "cb_prefix": "cb7"},
    {"id": "P1",  "name": "Period 1 (Normal)",    "start_str": "09:10 AM", "end_str": "10:12 AM", "cb_prefix": "cb1"},
    {"id": "P2",  "name": "Period 2 (Normal)",    "start_str": "10:17 AM", "end_str": "11:19 AM", "cb_prefix": "cb2"},
    {"id": "P3",  "name": "Period 3 (Normal)",    "start_str": "11:24 AM", "end_str": "01:01 PM", "cb_prefix": "cb3"},
    {"id": "P4",  "name": "Period 4 (Normal)",    "start_str": "01:06 PM", "end_str": "02:08 PM", "cb_prefix": "cb4"},
    {"id": "P5",  "name": "Period 5 (Normal)",    "start_str": "02:13 PM", "end_str": "03:15 PM", "cb_prefix": "cb5"},
    {"id": "PMA", "name": "PM Advisory (Normal)", "start_str": "03:20 PM", "end_str": "03:30 PM", "cb_prefix": "cb8"},
]

# Enrichment Monday-Thursday Schedule
M_TH_ENRICHMENT_PERIODS = [
    {"id": "AMA", "name": "AM Advisory (Enrichment)", "start_str": "08:30 AM", "end_str": "09:05 AM", "cb_prefix": "cb7"}, # AMA stays the same
    {"id": "P1",  "name": "Period 1 (Enrichment)",    "start_str": "09:10 AM", "end_str": "10:05 AM", "cb_prefix": "cb1"},
    {"id": "P2",  "name": "Period 2 (Enrichment)",    "start_str": "10:10 AM", "end_str": "11:05 AM", "cb_prefix": "cb2"},
    {"id": "P3",  "name": "Period 3 (Enrichment)",    "start_str": "11:10 AM", "end_str": "12:40 PM", "cb_prefix": "cb3"},
    {"id": "P4",  "name": "Period 4 (Enrichment)",    "start_str": "12:45 PM", "end_str": "01:40 PM", "cb_prefix": "cb4"},
    {"id": "P5",  "name": "Period 5 (Enrichment)",    "start_str": "01:45 PM", "end_str": "02:40 PM", "cb_prefix": "cb5"},
    {"id": "PMA", "name": "PM Advisory (Enrichment)", "start_str": "02:45 PM", "end_str": "03:30 PM", "cb_prefix": "cb8"},
]

ACTUAL_F_PERIODS = [
    {"id": "AMA", "name": "AM Advisory (Fri)", "start_str": "08:30 AM", "end_str": "08:40 AM", "cb_prefix": "cb7"},
    {"id": "P1",  "name": "Period 1 (Fri)",    "start_str": "08:45 AM", "end_str": "09:26 AM", "cb_prefix": "cb1"},
    {"id": "P2",  "name": "Period 2 (Fri)",    "start_str": "09:31 AM", "end_str": "10:12 AM", "cb_prefix": "cb2"},
    {"id": "P3",  "name": "Period 3 (Fri)",    "start_str": "10:17 AM", "end_str": "10:58 AM", "cb_prefix": "cb3"},
    {"id": "P4",  "name": "Period 4 (Fri)",    "start_str": "11:03 AM", "end_str": "11:44 AM", "cb_prefix": "cb4"},
    {"id": "P5",  "name": "Period 5 (Fri)",    "start_str": "11:49 AM", "end_str": "12:30 PM", "cb_prefix": "cb5"},
    {"id": "PMA", "name": "PM Advisory (Fri)", "start_str": "12:35 PM", "end_str": "01:30 PM", "cb_prefix": "cb8"},
]

SPECIAL_DAY_SCHEDULE = [
    {"id": "AMA", "name": "AM Advisory (Special Day)", "start_str": "08:00 AM", "end_str": "04:00 PM", "cb_prefix": "cb7"}
]

# Base configuration - M-Th part will be updated dynamically based on user choice
BASE_PERIODS_CONFIG = {
    "F": ACTUAL_F_PERIODS,
    "S": SPECIAL_DAY_SCHEDULE
}

# --- Helper Function for Safe Input with Exit Option ---
def safe_input(prompt_message):
    full_prompt = f"{prompt_message} (or type 'q', 'quit', 'exit' to exit application): "
    user_response = input(full_prompt).strip().lower()
    if user_response in ['exit', 'quit', 'q']:
        print("\nExiting application as requested...")
        sys.exit(0) 
    return user_response

# --- Helper Function to Load Credentials ---
def load_credentials(filepath, app_name):
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

# --- Helper Function for User Input ---
def get_user_day_and_period_selection(effective_periods_config): # Takes the dynamically built config
    print("\n🗓️ Select the day of the week:")
    days_map = {
        "1": ("M", "Monday"), "2": ("T", "Tuesday"), "3": ("W", "Wednesday"),
        "4": ("H", "Thursday"), "5": ("F", "Friday"), "6": ("S", "Special Day (e.g., Testing)")
    }
    for key, (_, name) in days_map.items():
        print(f"  {key}. {name}")

    selected_day_key = None
    day_choice_num = "" # ensure it's defined for the warning message if needed
    while selected_day_key is None:
        day_choice_num = safe_input("Enter number for the day")
        if day_choice_num in days_map:
            selected_day_key = days_map[day_choice_num][0]
        else:
            print("Invalid selection. Please try again.")
    
    if selected_day_key == "S":
        print("\n⚙️ Special Day selected. This assumes AM Advisory (AMA) processing.")
        current_day_periods = effective_periods_config.get("S") 
        selected_period_obj = current_day_periods[0]

        while True:
            try:
                filter_start_str_input = safe_input("Enter the START time for ID filtering (e.g., 08:00 AM)")
                datetime.strptime(filter_start_str_input, "%I:%M %p") 
                filter_start_str = filter_start_str_input
                break
            except ValueError:
                print("Invalid time format. Please use HH:MM AM/PM (e.g., 08:30 AM or 01:15 PM).")
        while True:
            try:
                filter_end_str_input = safe_input("Enter the END time for ID filtering (e.g., 03:00 PM)")
                datetime.strptime(filter_end_str_input, "%I:%M %p") 
                filter_end_str = filter_end_str_input
                break
            except ValueError:
                print("Invalid time format. Please use HH:MM AM/PM (e.g., 09:00 AM or 04:30 PM).")
        
        print(f"\n🔍 IDs will be filtered from Excel for student sign-ins between: {filter_start_str} and {filter_end_str}")
        print(f"Students in this list will be marked UL (Unexcused Late) for {selected_period_obj['name']}.")
        return selected_day_key, selected_period_obj, current_day_periods, filter_start_str, filter_end_str
    else:
        current_day_periods = effective_periods_config.get(selected_day_key)
        # This fallback should ideally not be needed if effective_periods_config is built correctly
        if not current_day_periods: 
            print(f"Warning: Period data for day key '{selected_day_key}' ({days_map.get(day_choice_num, ('','Unknown'))[1]}) is not fully defined. Defaulting to a standard M-Th Normal schedule structure for this day.")
            current_day_periods = M_TH_NORMAL_PERIODS # Fallback to normal M-Th if something went wrong

        print(f"\n📚 Select the period for which students are arriving (this period will be marked UL - Unexcused Late for these students):")
        for i, period in enumerate(current_day_periods):
            print(f"  {i+1}. {period['name']} ({period['start_str']} - {period['end_str']})")

        period_choice_idx = -1
        while True:
            try:
                period_choice_str = safe_input(f"Enter number for the period (1-{len(current_day_periods)})")
                period_choice_idx = int(period_choice_str) - 1
                if 0 <= period_choice_idx < len(current_day_periods):
                    break 
                else:
                    print("Invalid period number. Please try again.")
            except ValueError:
                print("Invalid input. Please enter a number.")

        selected_period_obj = current_day_periods[period_choice_idx]

        if period_choice_idx == 0: 
            filter_start_time_str = selected_period_obj['start_str']
        else:
            previous_period_obj = current_day_periods[period_choice_idx - 1]
            filter_start_time_str = previous_period_obj['end_str'] 
            
        filter_end_time_str = selected_period_obj['end_str']

        print(f"\n🔍 IDs will be filtered from the Excel sheet for student sign-in times between: {filter_start_time_str} and {filter_end_time_str}")
        print(f"Students in this list will be marked UL (Unexcused Late) for {selected_period_obj['name']}.")
        if period_choice_idx > 0:
            prev_period_names = ", ".join([p['name'] for p in current_day_periods[:period_choice_idx]])
            print(f"Periods prior to {selected_period_obj['name']} (i.e., {prev_period_names}) will be marked AU (Truant Absence).")
        
        return selected_day_key, selected_period_obj, current_day_periods, filter_start_time_str, filter_end_time_str

# --- Main Logic ---
def automate_raptor_and_powerschool(selected_period_object, all_periods_for_day, filter_start_str, filter_end_str):
    raptor_username, raptor_password = load_credentials(CREDENTIALS_FILE, 'raptor')
    if not raptor_username or not raptor_password:
        return

    powerschool_username, powerschool_password = load_credentials(CREDENTIALS_FILE, 'powerschool')
    if not powerschool_username or not powerschool_password:
        return

    print("Setting up WebDriver...")
    driver = None
    current_working_directory = os.getcwd()
    print(f"Files will be downloaded to the current working directory: {current_working_directory}")

    chrome_options = ChromeOptions()
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox") 
    chrome_options.add_argument("--window-size=1920,1080") 
    chrome_options.add_argument("--start-maximized") 

    prefs = {
        "download.default_directory": current_working_directory,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    downloaded_excel_file_path = None 

    try:
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
        wait = WebDriverWait(driver, 45) 

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
            print("Waiting for RaptorTech dashboard/landing page to load...")
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.product-tile")))
            print(f"RaptorTech dashboard element detected. Current URL: {driver.current_url}")
            print(f"Forcing direct navigation to target reports page: {INITIAL_AND_TARGET_REPORTS_URL}")
            driver.get(INITIAL_AND_TARGET_REPORTS_URL)
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
        print("Attempting to click 'Students' tab...")
        students_tab_xpath = "//a[@href='#tab3' and normalize-space(.)='Students']"
        students_tab = wait.until(EC.element_to_be_clickable((By.XPATH, students_tab_xpath)))
        students_tab.click()
        print("'Students' tab clicked.")
        wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@id='tab3']//h3[normalize-space(.)='Student Sign-In/Sign-Out History']")))
        print("Attempting to click 'Student Sign-In/Sign-Out History'...")
        sign_in_out_history_xpath = "//li[contains(@class, 'item') and .//h3[normalize-space(.)='Student Sign-In/Sign-Out History']]"
        sign_in_out_history_link = wait.until(EC.element_to_be_clickable((By.XPATH, sign_in_out_history_xpath)))
        sign_in_out_history_link.click()
        print("'Student Sign-In/Sign-Out History' link clicked.")
        wait.until(EC.presence_of_element_located((By.ID, "generate-report")))
        
        print("Attempting to click 'Generate Report' button...")
        generate_report_button = wait.until(EC.element_to_be_clickable((By.ID, "generate-report")))
        generate_report_button.click()
        print("'Generate Report' button clicked. Allowing time for report data to fully populate before export...")
        time.sleep(2.0) 
        
        print("Waiting for RaptorTech 'Export to Excel' button to be clickable...")
        wait.until(EC.element_to_be_clickable((By.ID, "btnExcelExport")))
        print("RaptorTech 'Export to Excel' button is ready.")
        
        btn_excel_export = driver.find_element(By.ID, "btnExcelExport") 
        files_before_download = set(os.listdir(current_working_directory))
        print(f"Files in current working directory before export: {len(files_before_download)}")
        btn_excel_export.click()
        print("'Export to Excel' button clicked.")
        
        print("Waiting for RaptorTech download to complete (max 60 seconds)...")
        # ... (download handling logic - unchanged) ...
        download_wait_timeout = 60
        download_poll_interval = 1
        time_waited = 0
        while time_waited < download_wait_timeout:
            files_after_download = set(os.listdir(current_working_directory))
            new_files = files_after_download - files_before_download
            for file_name in new_files:
                if file_name.endswith(".xlsx") and not file_name.endswith((".crdownload", ".tmp")):
                    print(f"New .xlsx file detected: {file_name}")
                    time.sleep(2) 
                    downloaded_excel_file_path = os.path.join(current_working_directory, file_name)
                    break
            if downloaded_excel_file_path:
                break
            time.sleep(download_poll_interval)
            time_waited += download_poll_interval
            if time_waited % 10 == 0 : print(f"Still waiting for download... ({time_waited}s / {download_wait_timeout}s)")

        if downloaded_excel_file_path:
            print(f"Download complete. File saved as: {downloaded_excel_file_path}")
            timestamp_str = datetime.now().strftime("%m-%d-%Y-%I-%M-%S-%p")
            new_file_name_base = f"student-sign-in-{timestamp_str}"
            new_file_name = f"{new_file_name_base}.xlsx"
            new_file_path = os.path.join(current_working_directory, new_file_name)
            counter = 1
            while os.path.exists(new_file_path):
                new_file_name = f"{new_file_name_base}-{counter}.xlsx"
                new_file_path = os.path.join(current_working_directory, new_file_name)
                counter += 1
            os.rename(downloaded_excel_file_path, new_file_path)
            downloaded_excel_file_path = new_file_path
            print(f"File renamed to: {downloaded_excel_file_path}")

            print("Extracting ID Numbers from the downloaded Excel file...")
            # ... (ID extraction logic - unchanged) ...
            ids_to_paste = ""
            try:
                df = pd.read_excel(downloaded_excel_file_path)
                if 'Date/Time' in df.columns:
                    df['Date/Time'] = pd.to_datetime(df['Date/Time'], errors='coerce')
                    
                    start_time_dt = datetime.strptime(filter_start_str, "%I:%M %p").time()
                    end_time_dt = datetime.strptime(filter_end_str, "%I:%M %p").time()
                    print(f"Filtering Excel data for times between {start_time_dt.strftime('%I:%M %p')} and {end_time_dt.strftime('%I:%M %p')}")

                    filtered_df = df[
                        (df['Date/Time'].dt.time >= start_time_dt) & 
                        (df['Date/Time'].dt.time <= end_time_dt) &
                        (df['Date/Time'].notna()) 
                    ]
                    
                    if 'ID Number' in filtered_df.columns:
                        extracted_ids = [str(int(float(id_val))) for id_val in filtered_df['ID Number'].dropna() if str(id_val).replace('.', '', 1).isdigit()]
                        ids_to_paste = "\n".join(extracted_ids)
                        print(f"Extracted {len(extracted_ids)} ID(s) from Excel file. IDs are:\n{ids_to_paste if ids_to_paste else 'None'}")
                    else:
                        print("Error: 'ID Number' column not found in the Excel file.")
                else:
                    print("Error: 'Date/Time' column not found in the Excel file for filtering.")
            except Exception as excel_error:
                print(f"Error processing Excel file: {excel_error}")


            if ids_to_paste:
                print(f"Navigating to PowerSchool login page: {POWERSCHOOL_LOGIN_URL}")
                # ... (PowerSchool login and navigation to batch attendance - unchanged) ...
                driver.get(POWERSCHOOL_LOGIN_URL)
                print("Attempting PowerSchool login...")
                username_field_ps = wait.until(EC.visibility_of_element_located((By.ID, "fieldUsername")))
                password_field_ps = wait.until(EC.visibility_of_element_located((By.ID, "fieldPassword")))
                username_field_ps.send_keys(powerschool_username)
                password_field_ps.send_keys(powerschool_password)
                print("PowerSchool credentials entered.")
                password_field_ps.submit()
                print("Attempted to log in to PowerSchool.")
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.dialogDivM.custom_link[title='MultiSelect - Students']")))
                print("PowerSchool login successful.")
                print("Attempting to click 'MultiSelect - Students' link...")
                multiselect_link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.dialogDivM.custom_link[title='MultiSelect - Students']")))
                multiselect_link.click()
                print("'MultiSelect - Students' link clicked.")
                print("Waiting for MultiSelect dialog to appear and textarea to be visible...")
                multiselect_textarea = wait.until(EC.visibility_of_element_located((By.ID, "multiSelValsStu")))
                multiselect_textarea.clear()
                multiselect_textarea.send_keys(ids_to_paste)
                print(f"IDs pasted into PowerSchool MultiSelect textarea.")
                print("Attempting to click 'Search' button in MultiSelect dialog...")
                search_button_xpath = "//button[contains(., 'Search') and @onclick=\"MultiSelect.searchType='admin'; MultiSelect.powerScheduler = 'Home'; MultiSelect.collectIDs();\"]"
                search_button = wait.until(EC.element_to_be_clickable((By.XPATH, search_button_xpath)))
                search_button.click()
                print("'Search' button clicked in MultiSelect dialog.")
                print("Waiting for student selection count to appear...")
                student_count_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//h2[contains(text(), 'Current Student Selection')]")))
                time.sleep(0.5) 
                count_text = student_count_element.text
                match = re.search(r'\((\d+)\)', count_text)
                if match: print(f"Number of students selected: {match.group(1)}")
                else: print("Could not extract student count.")
                
                print("Attempting to click 'Group Functions' dropdown button...")
                group_functions_button = wait.until(EC.element_to_be_clickable((By.ID, "selectFunctionDropdownButtonStudent")))
                group_functions_button.click()
                print("'Group Functions' dropdown button clicked.")
                time.sleep(0.5)
                print("Attempting to click 'Mass Update Attendance' link...")
                mass_update_attendance_link = wait.until(EC.element_to_be_clickable((By.ID, "lnk_studentsMassUpdateAttendance")))
                mass_update_attendance_link.click()
                print("'Mass Update Attendance' link clicked.")
                print("Waiting for redirection to batch attendance update page...")
                expected_url_batch_attendance = "https://ednovate.powerschool.com/admin/attendance/record/batch/meetinggroup.html?dothisfor=selected"
                wait.until(EC.url_to_be(expected_url_batch_attendance))
                print(f"Successfully redirected to: {driver.current_url}")


                target_period_index = -1
                for i, p in enumerate(all_periods_for_day): 
                    if p['id'] == selected_period_object['id']:
                        target_period_index = i
                        break
                
                if target_period_index == -1:
                    print(f"❌ Error: Could not find selected period {selected_period_object['name']} in the period list for processing.")
                else:
                    periods_to_mark_AU = all_periods_for_day[:target_period_index]
                    if periods_to_mark_AU:
                        print(f"\n Marking {len(periods_to_mark_AU)} previous period(s) as AU (Truant Absence)...")
                        # ... (AU marking logic - unchanged) ...
                        for period_obj in periods_to_mark_AU:
                            print(f"  Selecting checkboxes for {period_obj['name']} (A/B columns)...")
                            try:
                                for col in ['1', '2']: 
                                    cb_xpath = f"//input[@type='checkbox' and @name='{period_obj['cb_prefix']};{col}']"
                                    checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, cb_xpath)))
                                    if not checkbox.is_selected():
                                        checkbox.click()
                                print(f"    ✅ Checkboxes for {period_obj['name']} selected.")
                            except Exception as e:
                                print(f"    ❌ Error selecting checkboxes for {period_obj['name']}: {e}")
                        
                        print("  Selecting 'AU' as attendance code...")
                        attendance_code_select_au = wait.until(EC.element_to_be_clickable((By.NAME, "att_attcodelist")))
                        select_au = Select(attendance_code_select_au)
                        select_au.select_by_value("AU")
                        print("    ✅ 'AU' selected.")

                        safe_input("👉 Review selections for AU. Press Enter to SUBMIT and continue")
                        submit_button_ps_au = wait.until(EC.element_to_be_clickable((By.ID, "btnSubmit")))
                        submit_button_ps_au.click()
                        print("    ✅ Submit button clicked for AU marking. Waiting for page to process...")
                        wait.until(EC.staleness_of(submit_button_ps_au))
                        wait.until(EC.element_to_be_clickable((By.NAME, "att_attcodelist"))) 
                        print("    ✅ Page processed after AU submission.")
                        time.sleep(1) 

                    print(f"\n Marking current period ({selected_period_object['name']}) as UL (Unexcused Late)...")
                    # ... (UL marking logic - unchanged) ...
                    try: 
                        clear_button = driver.find_element(By.XPATH, "//a[@name='btnClear' and normalize-space(.)='Clear']")
                        clear_button.click()
                        print("  Clicked 'Clear' button to reset period checkboxes.")
                        time.sleep(0.5) 
                    except Exception:
                        print("  'Clear' button not found or clickable, proceeding with selections.")

                    print(f"  Selecting checkboxes for {selected_period_object['name']} (A/B columns)...")
                    try:
                        for col in ['1', '2']: 
                            cb_xpath_ul = f"//input[@type='checkbox' and @name='{selected_period_object['cb_prefix']};{col}']"
                            checkbox_ul = wait.until(EC.element_to_be_clickable((By.XPATH, cb_xpath_ul)))
                            if not checkbox_ul.is_selected():
                                checkbox_ul.click()
                        print(f"    ✅ Checkboxes for {selected_period_object['name']} selected for UL.")
                    except Exception as e:
                        print(f"    ❌ Error selecting checkboxes for {selected_period_object['name']} for UL: {e}")

                    print("  Selecting 'UL' as attendance code...")
                    attendance_code_select_ul = wait.until(EC.element_to_be_clickable((By.NAME, "att_attcodelist")))
                    select_ul = Select(attendance_code_select_ul)
                    select_ul.select_by_value("UL")
                    print("    ✅ 'UL' selected.")

                    safe_input("👉 Review selections for UL. Press Enter to SUBMIT")
                    submit_button_ps_ul = wait.until(EC.element_to_be_clickable((By.ID, "btnSubmit")))
                    submit_button_ps_ul.click()
                    print("    ✅ Submit button clicked for UL marking.")
                    wait.until(EC.staleness_of(submit_button_ps_ul))
                    print("    ✅ Page processed after UL submission.")

                print("\n🎉 PowerSchool batch update process complete. Browser will stay open for 10 seconds for final review.")
                time.sleep(10)
            else:
                print("No IDs extracted from Excel, skipping PowerSchool automation.")
        else:
            print("❌ Error: RaptorTech download did not complete or .xlsx file not found within the timeout.")
            if driver: driver.save_screenshot('download_error_screenshot.png')

        print("Automation process complete. Browser will close in 5 seconds...")
        time.sleep(5)

    except SystemExit: # Catches sys.exit() calls from safe_input more gracefully at this level too
        print("Application exited by user during automation process.")
    except Exception as e:
        print(f"❌ An unexpected error occurred during automation: {e}")
        if driver:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            error_screenshot_name = f'unexpected_error_screenshot_{timestamp}.png'
            try:
                driver.save_screenshot(error_screenshot_name)
                print(f"📸 Screenshot '{error_screenshot_name}' saved for debugging.")
            except Exception as se:
                print(f"Could not save screenshot: {se}")
    finally:
        if driver:
            print("Closing the browser.")
            driver.quit()

if __name__ == "__main__":
    print("🚀 Welcome to the Attendance Automation Script!")
    print("--------------------------------------------")
    print("📅 Please select the schedule type for Monday-Thursday operations for this session:")
    print("  1. Normal Schedule")
    print("  2. Enrichment Schedule")
    
    active_m_th_schedule = None
    while active_m_th_schedule is None:
        schedule_choice = safe_input("Enter M-Th schedule choice (1 or 2)")
        if schedule_choice == "1":
            active_m_th_schedule = M_TH_NORMAL_PERIODS
            print(" Normal Schedule selected for M-Th operations.")
        elif schedule_choice == "2":
            active_m_th_schedule = M_TH_ENRICHMENT_PERIODS
            print(" Enrichment Schedule selected for M-Th operations.")
        else:
            print("Invalid choice. Please enter 1 or 2.")

    effective_periods_config = {
        "M": active_m_th_schedule,
        "T": active_m_th_schedule,
        "W": active_m_th_schedule,
        "H": active_m_th_schedule,
        "F": ACTUAL_F_PERIODS,    
        "S": SPECIAL_DAY_SCHEDULE 
    }
    
    if not ACTUAL_F_PERIODS or len(ACTUAL_F_PERIODS) < 1: # Check if Friday schedule definition is missing
        print("🚨 WARNING: Friday period schedule (ACTUAL_F_PERIODS) appears to be missing or empty in the script.")

    try:
        selected_day_key, selected_period, all_periods_for_day, filter_start, filter_end = get_user_day_and_period_selection(effective_periods_config)
        
        print(f"\n--- Script Configuration for this Run ---")
        print(f"Selected Day Type: {selected_day_key}")
        if selected_day_key in ['M','T','W','H']:
            if active_m_th_schedule == M_TH_NORMAL_PERIODS:
                print(f"M-Th Schedule Type: Normal")
            else:
                print(f"M-Th Schedule Type: Enrichment")
        print(f"Target Period for UL: {selected_period['name']}")
        print(f"Excel Date/Time Filter: From {filter_start} to {filter_end}")
        print("--- Starting Automation ---")
        
        automate_raptor_and_powerschool(selected_period, all_periods_for_day, filter_start, filter_end)
    except SystemExit: 
        print("Application exited by user.")
    except Exception as main_err:
        print(f"❌ An error occurred in the main execution block: {main_err}")
