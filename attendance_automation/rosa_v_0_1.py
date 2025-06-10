import json
import time
import os
import sys
from datetime import datetime, timedelta
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
import numpy as np

# --- Configuration ---
CREDENTIALS_FILE = 'credentials.json'
INITIAL_AND_TARGET_REPORTS_URL = 'https://apps.raptortech.com/Reports/Home/VisitorReports'
POWERSCHOOL_LOGIN_URL = 'https://ednovate.powerschool.com/admin/pw.html'
POWERSCHOOL_MEETING_ATTENDANCE_URL = 'https://ednovate.powerschool.com/admin/attendance/functions/attendancestatus.meeting.html'

# --- Output Directory Paths ---
RAPTOR_REPORTS_DIR = 'raptor_reports'
DAILY_MASTER_REPORTS_DIR = 'daily_raptor_report_master'
MEET_ATTENDANCE_DIR = 'meet_attendance'

# --- School Period Definitions ---
# (Rest of your period definitions as they were)
M_TH_NORMAL_PERIODS = [
    {"id": "AMA", "name": "AM Advisory (Normal)", "start_str": "08:30 AM", "end_str": "09:05 AM", "cb_prefix": "cb7"},
    {"id": "P1",  "name": "Period 1 (Normal)",    "start_str": "09:10 AM", "end_str": "10:12 AM", "cb_prefix": "cb1"},
    {"id": "P2",  "name": "Period 2 (Normal)",    "start_str": "10:17 AM", "end_str": "11:19 AM", "cb_prefix": "cb2"},
    {"id": "P3",  "name": "Period 3 (Normal)",    "start_str": "11:24 AM", "end_str": "01:01 PM", "cb_prefix": "cb3"},
    {"id": "P4",  "name": "Period 4 (Normal)",    "start_str": "01:06 PM", "end_str": "02:08 PM", "cb_prefix": "cb4"},
    {"id": "P5",  "name": "Period 5 (Normal)",    "start_str": "02:13 PM", "end_str": "03:15 PM", "cb_prefix": "cb5"},
    {"id": "PMA", "name": "PM Advisory (Normal)", "start_str": "03:20 PM", "end_str": "03:30 PM", "cb_prefix": "cb8"},
]

M_TH_ENRICHMENT_PERIODS = [
    {"id": "AMA", "name": "AM Advisory (Enrichment)", "start_str": "08:30 AM", "end_str": "09:05 AM", "cb_prefix": "cb7"},
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

ALL_POSSIBLE_PERIOD_CBS = [
    "cb7;1", "cb7;2", # AMA
    "cb1;1", "cb1;2", # P1
    "cb2;1", "cb2;2", # P2
    "cb3;1", "cb3;2", # P3
    "cb4;1", "cb4;2", # P4
    "cb5;1", "cb5;2", # P5
    "cb8;1", "cb8;2"  # PMA
]

# --- Attendance Code Definitions ---
ABSENCE_CODES = ['A', 'AU', 'AB', 'ABS', 'X'] # Absent, Truant, etc.
# Updated PRESENT_CODES based on user feedback:
# Removed 'P' as it's generally blank. Added '' (empty string) for blank cells.
# Added 'INA' and 'S' as they are codes to look for in PowerSchool implying presence.
PRESENT_CODES = ['T', 'UL', 'LE', 'INA', 'S', ''] # Tardy, Unexcused Late, Excused Late, In Attendance, Signed In, Blank Cell


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

# --- Helper Function for User Input (Raptor Automation) ---
def get_user_day_and_period_selection(effective_periods_config):
    print("\n🗓️ Select the day of the week:")
    days_map = {
        "1": ("M", "Monday"), "2": ("T", "Tuesday"), "3": ("W", "Wednesday"),
        "4": ("H", "Thursday"), "5": ("F", "Friday"), "6": ("S", "Special Day (e.g., Testing)")
    }
    for key, (_, name) in days_map.items():
        print(f"  {key}. {name}")

    selected_day_key = None
    day_choice_num = ""
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
        if not current_day_periods:
            print(f"Warning: Period data for day key '{selected_day_key}' ({days_map.get(day_choice_num, ('','Unknown'))[1]}) is not fully defined. Defaulting to a standard M-Th Normal schedule structure for this day.")
            current_day_periods = M_TH_NORMAL_PERIODS

        # --- UPDATED PERIOD SELECTION PROMPT ---
        period_input_map = {} # Maps user input string to actual period object
        display_options_str = []
        for period_obj in current_day_periods:
            user_input_val = None
            if period_obj['id'] == 'AMA':
                user_input_val = "7"
            elif period_obj['id'].startswith('P') and len(period_obj['id']) == 2: # P1, P2, etc.
                user_input_val = period_obj['id'][1] # Extract the number from P1 -> '1'
            elif period_obj['id'] == 'PMA':
                user_input_val = "8"
            
            if user_input_val:
                period_input_map[user_input_val] = period_obj
                display_options_str.append(f"  {user_input_val}. {period_obj['name']} ({period_obj['start_str']} - {period_obj['end_str']})")
        
        print("\n📚 Select the period for which students are arriving (this period will be marked UL - Unexcused Late for these students):")
        for s in display_options_str:
            print(s)

        selected_period_obj = None
        while selected_period_obj is None:
            period_choice_str = safe_input("Enter number for the period (e.g., 7 for AMA, 1 for P1, 8 for PMA)")
            if period_choice_str in period_input_map:
                selected_period_obj = period_input_map[period_choice_str]
            else:
                print("Invalid period number. Please try again.")
        # --- END UPDATED PERIOD SELECTION PROMPT ---

        # Determine filter times based on the *selected_period_obj* and its position in the *ordered* current_day_periods list
        period_index_in_list = -1
        for i, p in enumerate(current_day_periods):
            if p['id'] == selected_period_obj['id']:
                period_index_in_list = i
                break

        if period_index_in_list == -1:
            print("Error: Could not find the selected period in the defined schedule. Exiting.")
            sys.exit(1)

        if period_index_in_list == 0:
            filter_start_time_str = selected_period_obj['start_str']
        else:
            previous_period_obj = current_day_periods[period_index_in_list - 1]
            filter_start_time_str = previous_period_obj['end_str']

        filter_end_time_str = selected_period_obj['end_str']

        print(f"\n🔍 IDs will be filtered from the Excel sheet for student sign-in times between: {filter_start_time_str} and {filter_end_time_str}")
        print(f"Students in this list will be marked UL (Unexcused Late) for {selected_period_obj['name']}.")
        if period_index_in_list > 0:
            prev_period_names = ", ".join([p['name'] for p in current_day_periods[:period_index_in_list]])
            print(f"Periods prior to {selected_period_obj['name']} (i.e., {prev_period_names}) will be marked AU (Truant Absence).")

        return selected_day_key, selected_period_obj, current_day_periods, filter_start_str, filter_end_str

def setup_webdriver(download_dir):
    print("Setting up WebDriver...")
    chrome_options = ChromeOptions()
    # chrome_options.add_argument("--headless") # Keep commented for visual debugging if needed
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--start-maximized")

    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
    wait = WebDriverWait(driver, 45)
    return driver, wait

def powerschool_login(driver, wait, username, password):
    print(f"Navigating to PowerSchool login page: {POWERSCHOOL_LOGIN_URL}")
    driver.get(POWERSCHOOL_LOGIN_URL)
    print("Attempting PowerSchool login...")
    username_field_ps = wait.until(EC.visibility_of_element_located((By.ID, "fieldUsername")))
    password_field_ps = wait.until(EC.presence_of_element_located((By.ID, "fieldPassword"))) # Corrected: presence_of_element_located
    username_field_ps.send_keys(username)
    password_field_ps.send_keys(password)
    print("PowerSchool credentials entered.")
    password_field_ps.submit()
    print("Attempted to log in to PowerSchool.")
    # Wait for an element that indicates successful login to the main dashboard
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.dialogDivM.custom_link[title='MultiSelect - Students']")))
        print("PowerSchool login successful.")
        return True
    except Exception as e:
        print(f"PowerSchool login failed or main dashboard element not found: {e}")
        driver.save_screenshot('powerschool_login_failure.png')
        return False

# --- UPDATED: Consolidate Attendance Function ---
def consolidate_attendance():
    powerschool_username, powerschool_password = load_credentials(CREDENTIALS_FILE, 'powerschool')
    if not powerschool_username or not powerschool_password:
        return

    os.makedirs(MEET_ATTENDANCE_DIR, exist_ok=True)
    print(f"Ensured '{MEET_ATTENDANCE_DIR}' directory exists.")

    download_directory_for_chrome = os.path.join(os.getcwd(), MEET_ATTENDANCE_DIR)
    print(f"PowerSchool Meeting Attendance files will be downloaded to: {download_directory_for_chrome}")

    # --- UPDATED PROMPT: min_total_absences ---
    min_total_absences = 2 # Default value
    while True:
        try:
            absences_input = safe_input("Enter the minimum number of TOTAL absences (e.g., 2, 3, 4) to flag a student: ")
            min_total_absences = int(absences_input)
            if min_total_absences < 1:
                print("Please enter a positive integer.")
            else:
                break
        except ValueError:
            print("Invalid input. Please enter a number.")
    print(f"Students will be flagged if they have {min_total_absences} or more TOTAL absences AND no presence codes (T, UL, LE, INA, S, or blank) for the day.")
    # --- END UPDATED PROMPT ---

    driver = None
    try:
        driver, wait = setup_webdriver(download_directory_for_chrome)

        if not powerschool_login(driver, wait, powerschool_username, powerschool_password):
            print("Failed to login to PowerSchool. Exiting consolidation process.")
            return

        print(f"Navigating to Meeting Attendance Status page: {POWERSCHOOL_MEETING_ATTENDANCE_URL}")
        driver.get(POWERSCHOOL_MEETING_ATTENDANCE_URL)
        wait.until(EC.url_to_be(POWERSCHOOL_MEETING_ATTENDANCE_URL))
        print("Successfully navigated to Meeting Attendance Status page.")

        # --- Download the Excel file ---
        print("Attempting to click the export dropdown menu...")
        
        # --- REFINED LOCATOR AND WAITS FOR EXPORT DROPDOWN AND OPTIONS ---
        export_dropdown_button_id = "exportDropDownButton"
        export_dropdown_button_locator = (By.ID, export_dropdown_button_id)

        # 1. Wait for the export dropdown button to be clickable. This returns the WebElement.
        print(f"Waiting for export dropdown button (ID: {export_dropdown_button_id}) to be clickable...")
        export_dropdown_button = wait.until(EC.element_to_be_clickable(export_dropdown_button_locator))
        
        # 2. Now that the element is clickable, wait for its 'aria-disabled' attribute to be 'false'.
        #    This ensures the Angular logic has completed and enabled the button.
        print(f"Waiting for export dropdown button (ID: {export_dropdown_button_id}) to be fully enabled (aria-disabled='false')...")
        wait.until(lambda driver: export_dropdown_button.get_attribute("aria-disabled") == "false")
        
        # Now, the 'export_dropdown_button' variable still holds the WebElement, and we know it's enabled.
        export_dropdown_button.click()
        print("Export dropdown button clicked.")
        
        # Add a very short sleep to allow the dropdown to start rendering visually
        time.sleep(0.5) 

        # 2. Wait for the dropdown list (ul) to become visible and its aria-hidden attribute to be false
        print("Waiting for the download options list to become visible...")
        # The ul element often has a distinct class or is a sibling/child of the button's parent.
        # Based on typical PowerSchool/AngularJS patterns, it's often a ul with specific classes that appears/disappears.
        # Let's assume it's still related to 'multiButtonList' and 'groupFunctions' as before,
        # but now we confirm it's visible by checking its 'aria-hidden' attribute.
        download_options_list_xpath = "//ul[contains(@class, 'multiButtonList') and contains(@class, 'groupFunctions') and @aria-hidden='false']"
        wait.until(EC.visibility_of_element_located((By.XPATH, download_options_list_xpath)))
        print("Download options list is visible.")

        # 3. Wait for the specific Excel option within the now-visible list to be clickable
        print("Attempting to click 'Excel Spreadsheet (XLSX)' download option...")
        excel_option_id = "export-option-excelOptionId" # This ID is very specific and reliable if it exists
        excel_download_option = wait.until(EC.element_to_be_clickable((By.ID, excel_option_id)))

        files_before_download = set(os.listdir(download_directory_for_chrome))
        print(f"Files in download directory '{MEET_ATTENDANCE_DIR}' before export: {len(files_before_download)}")
        
        excel_download_option.click()
        print("'Excel Spreadsheet (XLSX)' option clicked.")
        
        # Add a short sleep after clicking the download option to ensure the browser
        # registers the click and initiates the download.
        time.sleep(1.0) 

        downloaded_excel_file_path = None
        download_wait_timeout = 60
        download_poll_interval = 1
        time_waited = 0
        while time_waited < download_wait_timeout:
            files_after_download = set(os.listdir(download_directory_for_chrome))
            new_files = files_after_download - files_before_download
            for file_name in new_files:
                if file_name.endswith(".xlsx") and not file_name.endswith((".crdownload", ".tmp")):
                    print(f"New .xlsx file detected: {file_name}")
                    time.sleep(2) # Give it a moment to finish writing
                    downloaded_excel_file_path = os.path.join(download_directory_for_chrome, file_name)
                    break
            if downloaded_excel_file_path:
                break
            time.sleep(download_poll_interval)
            time_waited += download_poll_interval
            if time_waited % 10 == 0: print(f"Still waiting for download... ({time_waited}s / {download_wait_timeout}s)")

        if not downloaded_excel_file_path:
            print("❌ Error: PowerSchool Meeting Attendance download did not complete or .xlsx file not found within the timeout.")
            driver.save_screenshot('meet_attendance_download_error.png')
            return

        # --- Rename and move the downloaded file ---
        timestamp_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        new_file_name_base = f"meet-attendance-raw_{timestamp_str}"
        new_file_name = f"{new_file_name_base}.xlsx"
        final_meet_attendance_path = os.path.join(download_directory_for_chrome, new_file_name)

        counter = 1
        while os.path.exists(final_meet_attendance_path): # Ensure unique name if somehow same timestamp
            new_file_name = f"{new_file_name_base}-{counter}.xlsx"
            final_meet_attendance_path = os.path.join(download_directory_for_chrome, new_file_name)
            counter += 1

        os.rename(downloaded_excel_file_path, final_meet_attendance_path)
        print(f"Raw Meeting Attendance report renamed and stored: {final_meet_attendance_path}")

        # --- Process the Excel file for total absences ---
        print("Processing Meeting Attendance report for total absences...")
        try:
            df_meet_attendance = pd.read_excel(final_meet_attendance_path)

            # The min_total_absences variable is now set by user input above

            # The script will process all rows in the downloaded Excel file.
            print("Date filtering for Meeting Attendance report is disabled. Processing all rows.")

            # Dynamically identify period columns (e.g., 'AMA', '1', '2', '3', '4', '5', 'PMA')
            period_cols = []
            for col in df_meet_attendance.columns:
                if col in ['AMA', 'PMA']:
                    period_cols.append(col)
                elif re.fullmatch(r'\d+', str(col)): # Check if column name is purely numeric
                    try:
                        num_col = int(col)
                        if 1 <= num_col <= 5: # Assuming periods 1 through 5
                            period_cols.append(col)
                    except ValueError:
                        pass
            
            # Sort period columns, ensuring 'AMA' comes before numbers and 'PMA' comes after numbers
            def sort_key(col_name):
                if col_name == 'AMA':
                    return 0 # Comes first
                elif re.fullmatch(r'\d+', str(col_name)):
                    return int(col_name) # Sort by number
                elif col_name == 'PMA':
                    return 999 # Comes last (assuming max period number is less than 999)
                return 500 # For any other unrecognized column names that might accidentally be in the list

            period_cols.sort(key=sort_key)


            if not period_cols:
                print("Warning: No recognized period columns found (e.g., 'AMA', '1', '2', '3', '4', '5', 'PMA'). Cannot check for total absences.")
                total_absence_ids = set() # No IDs to process
            else:
                print(f"Identified period columns (in processing order): {period_cols}")
                
                # --- UPDATED LOGIC FOR TOTAL ABSENCES ---
                total_absence_ids = set() # Final set of students to flag
                
                # Ensure 'Student Number' column exists
                if 'Student Number' not in df_meet_attendance.columns:
                    print("Error: 'Student Number' column not found in the Meeting Attendance report. Cannot process for absences.")
                    return # Exit function if essential column is missing

                # Iterate through each student (row) in the DataFrame
                for index, row in df_meet_attendance.iterrows():
                    student_number = str(row['Student Number']).strip()
                    if not student_number or pd.isna(student_number):
                        continue # Skip rows with no valid student number

                    total_absences_for_day = 0
                    has_any_present_code_today = False # Track if student was present at any point today

                    for period_col in period_cols:
                        period_status = str(row.get(period_col, '')).strip().upper()
                        
                        if period_status in ABSENCE_CODES:
                            total_absences_for_day += 1
                        elif period_status in PRESENT_CODES: # Any present code (T, UL, LE, INA, S, or blank)
                            has_any_present_code_today = True # Mark that student was present
                    
                    # After checking all periods for a student:
                    # Only add student if they hit the total absence threshold AND were never marked present today
                    if total_absences_for_day >= min_total_absences and not has_any_present_code_today:
                        total_absence_ids.add(student_number)
                        print(f"  Student Number {student_number} identified with {min_total_absences} or more total absences and no presence codes.") # For debugging
                # --- END UPDATED LOGIC ---
                    
            ids_to_paste = "\n".join(list(total_absence_ids))
            print(f"Identified {len(total_absence_ids)} student(s) with {min_total_absences} or more total absences AND no presence codes for the day.")
            if not ids_to_paste:
                print("No students found with the specified total absences criteria. Exiting consolidation.")
                return

        except Exception as excel_process_error:
            print(f"Error processing Meeting Attendance Excel file: {excel_process_error}")
            driver.save_screenshot('meet_attendance_process_error.png')
            return

        # --- PowerSchool MultiSelect and Mass Update for Consolidation ---
        print("Navigating to PowerSchool MultiSelect for identified students...")
        driver.get("https://ednovate.powerschool.com/admin/home.html") # Go to home page to access MultiSelect
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.dialogDivM.custom_link[title='MultiSelect - Students']")))

        print("Attempting to click 'MultiSelect - Students' link...")
        multiselect_link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.dialogDivM.custom_link[title='MultiSelect - Students']")))
        multiselect_link.click()
        print("'MultiSelect - Students' link clicked.")

        print("Waiting for MultiSelect dialog to appear and textarea to be visible...")
        multiselect_textarea = wait.until(EC.visibility_of_element_located((By.ID, "multiSelValsStu")))
        multiselect_textarea.clear()
        multiselect_textarea.send_keys(ids_to_paste)
        print(f"IDs pasted into PowerSchool MultiSelect textarea for {len(total_absence_ids)} students.")

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

        # --- Mark all relevant periods as AU ---
        print("\nMarking all standard periods (AMA, P1-P5, PMA) as AU (Truant Absence)...")
        # Ensure the date is set to today if it's not already, though PowerSchool typically defaults to today.
        # Check if the date field is present and set it if necessary.
        # For simplicity, assuming the current date is automatically set by PowerSchool for mass update.

        for cb_name in ALL_POSSIBLE_PERIOD_CBS:
            print(f"  Selecting checkbox for {cb_name}...")
            try:
                cb_xpath = f"//input[@type='checkbox' and @name='{cb_name}']"
                checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, cb_xpath)))
                if not checkbox.is_selected():
                    checkbox.click()
                print(f"    ✅ Checkbox {cb_name} selected.")
            except Exception as e:
                print(f"    ❌ Error selecting checkbox {cb_name}: {e}")

        print("  Selecting 'AU' as attendance code...")
        attendance_code_select_au = wait.until(EC.element_to_be_clickable((By.NAME, "att_attcodelist")))
        select_au = Select(attendance_code_select_au)
        select_au.select_by_value("AU")
        print("    ✅ 'AU' selected.")

        safe_input("👉 Review ALL selected periods for AU. Press Enter to SUBMIT attendance update.")
        submit_button_ps_au = wait.until(EC.element_to_be_clickable((By.ID, "btnSubmit")))
        submit_button_ps_au.click()
        print("    ✅ Submit button clicked for AU marking. Waiting for page to process...")
        wait.until(EC.staleness_of(submit_button_ps_au))
        # Wait for some element to reappear or page to stabilize after submission
        wait.until(EC.presence_of_element_located((By.NAME, "att_attcodelist")))
        print("    ✅ Page processed after AU submission.")

        print("\n🎉 PowerSchool consolidation process complete. Browser will stay open for 10 seconds for final review.")
        time.sleep(10)

    except SystemExit:
        print("Application exited by user during consolidation process.")
    except Exception as e:
        print(f"❌ An unexpected error occurred during consolidation automation: {e}")
        if driver:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            error_screenshot_name = f'consolidation_error_screenshot_{timestamp}.png'
            try:
                driver.save_screenshot(error_screenshot_name)
                print(f"📸 Screenshot '{error_screenshot_name}' saved for debugging.")
            except Exception as se:
                print(f"Could not save screenshot: {se}")
    finally:
        if driver:
            print("Closing the browser.")
            driver.quit()

# --- Main Logic for Raptor Attendance Automation ---
def automate_raptor_and_powerschool(selected_period_object, all_periods_for_day, filter_start_str, filter_end_str):
    raptor_username, raptor_password = load_credentials(CREDENTIALS_FILE, 'raptor')
    if not raptor_username or not raptor_password:
        return

    powerschool_username, powerschool_password = load_credentials(CREDENTIALS_FILE, 'powerschool')
    if not powerschool_username or not powerschool_password:
        return

    # --- NEW: Ensure output directories exist ---
    os.makedirs(RAPTOR_REPORTS_DIR, exist_ok=True)
    os.makedirs(DAILY_MASTER_REPORTS_DIR, exist_ok=True)
    print(f"Ensured '{RAPTOR_REPORTS_DIR}' and '{DAILY_MASTER_REPORTS_DIR}' directories exist.")

    print("Setting up WebDriver...")
    driver = None

    # --- Set download directory to the raptor_reports folder ---
    download_directory_for_chrome = os.path.join(os.getcwd(), RAPTOR_REPORTS_DIR)
    print(f"RaptorTech Excel files will be downloaded to: {download_directory_for_chrome}")

    driver, wait = setup_webdriver(download_directory_for_chrome)
    downloaded_excel_file_path = None

    try:
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
        files_before_download = set(os.listdir(download_directory_for_chrome))
        print(f"Files in download directory '{RAPTOR_REPORTS_DIR}' before export: {len(files_before_download)}")
        btn_excel_export.click()
        print("'Export to Excel' button clicked.")

        print("Waiting for RaptorTech download to complete (max 60 seconds)...")
        download_wait_timeout = 60
        download_poll_interval = 1
        time_waited = 0
        while time_waited < download_wait_timeout:
            files_after_download = set(os.listdir(download_directory_for_chrome))
            new_files = files_after_download - files_before_download
            for file_name in new_files:
                if file_name.endswith(".xlsx") and not file_name.endswith((".crdownload", ".tmp")):
                    print(f"New .xlsx file detected: {file_name}")
                    time.sleep(2)
                    downloaded_excel_file_path = os.path.join(download_directory_for_chrome, file_name)
                    break
            if downloaded_excel_file_path:
                break
            time.sleep(download_poll_interval)
            time_waited += download_poll_interval
            if time_waited % 10 == 0 : print(f"Still waiting for download... ({time_waited}s / {download_wait_timeout}s)")

        if downloaded_excel_file_path:
            print(f"Download complete. File saved as: {downloaded_excel_file_path}")
            timestamp_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            new_file_name_base = f"raptor-sign-in-raw_{timestamp_str}"
            new_file_name = f"{new_file_name_base}.xlsx"
            final_raptor_report_path = os.path.join(download_directory_for_chrome, new_file_name)

            counter = 1
            while os.path.exists(final_raptor_report_path):
                new_file_name = f"{new_file_name_base}-{counter}.xlsx"
                final_raptor_report_path = os.path.join(download_directory_for_chrome, new_file_name)
                counter += 1

            os.rename(downloaded_excel_file_path, final_raptor_report_path)
            downloaded_excel_file_path = final_raptor_report_path
            print(f"Raw Raptor report renamed and stored: {downloaded_excel_file_path}")

            print("Extracting ID Numbers and Full Names from the downloaded Excel file for processing...")
            ids_to_paste = ""
            students_for_master_report = pd.DataFrame()
            try:
                df_raptor = pd.read_excel(downloaded_excel_file_path)

                expected_cols = ['Date/Time', 'ID Number', 'First Name', 'Last Name']
                if not all(col in df_raptor.columns for col in expected_cols):
                    print(f"Error: Missing one or more expected columns ({expected_cols}) in the Raptor report.")
                    if 'ID Number' not in df_raptor.columns:
                        raise ValueError("Required 'ID Number' column missing.")
                    if 'Date/Time' not in df_raptor.columns:
                        raise ValueError("Required 'Date/Time' column missing for filtering.")

                df_raptor['Date/Time'] = pd.to_datetime(df_raptor['Date/Time'], errors='coerce')

                start_time_dt = datetime.strptime(filter_start_str, "%I:%M %p").time()
                end_time_dt = datetime.strptime(filter_end_str, "%I:%M %p").time()
                print(f"Filtering Excel data for times between {start_time_dt.strftime('%I:%M %p')} and {end_time_dt.strftime('%I:%M %p')}")

                filtered_df = df_raptor[
                    (df_raptor['Date/Time'].dt.time >= start_time_dt) &
                    (df_raptor['Date/Time'].dt.time <= end_time_dt) &
                    (df_raptor['Date/Time'].notna())
                ].copy()

                if 'ID Number' in filtered_df.columns:
                    filtered_df['ID Number'] = filtered_df['ID Number'].dropna().apply(lambda x: str(int(float(x))) if str(x).replace('.', '', 1).isdigit() else np.nan)
                    filtered_df.dropna(subset=['ID Number'], inplace=True)

                    extracted_ids = filtered_df['ID Number'].tolist()
                    ids_to_paste = "\n".join(extracted_ids)
                    print(f"Extracted {len(extracted_ids)} unique ID(s) for PowerSchool processing.")

                    if not filtered_df.empty:
                        students_for_master_report = filtered_df[['ID Number', 'First Name', 'Last Name', 'Date/Time']].copy()
                        students_for_master_report['Timestamp Processed'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        students_for_master_report['Marked Period'] = selected_period_object['name']
                else:
                    print("Error: 'ID Number' column not found in the filtered Excel data.")
            except Exception as excel_error:
                print(f"Error processing Excel file for IDs: {excel_error}")

            # --- Update the Daily Master Report ---
            if not students_for_master_report.empty:
                today_str = datetime.now().strftime("%Y-%m-%d")
                master_file_name = f"daily_raptor_master_report_{today_str}.xlsx"
                master_file_path = os.path.join(DAILY_MASTER_REPORTS_DIR, master_file_name)

                existing_master_df = pd.DataFrame()
                if os.path.exists(master_file_path):
                    try:
                        existing_master_df = pd.read_excel(master_file_path)
                        print(f"Loaded existing master report for today: {master_file_path}")
                    except Exception as e:
                        print(f"Error loading existing master report: {e}. Starting a new one.")

                combined_df = pd.concat([existing_master_df, students_for_master_report], ignore_index=True)
                final_master_df = combined_df.drop_duplicates(subset=['ID Number', 'Marked Period'], keep='first')

                if 'First Name' in final_master_df.columns and 'Last Name' in final_master_df.columns:
                    final_master_df['Full Name'] = final_master_df['First Name'].fillna('') + ' ' + final_master_df['Last Name'].fillna('')
                else:
                    final_master_df['Full Name'] = ''
                    print("Warning: 'First Name' or 'Last Name' columns not found for Full Name creation.")

                # Define the desired column order. Ensure all columns present in final_master_df are included.
                # If a column doesn't exist in final_master_df, it will be skipped by this selection.
                desired_columns_order = [
                    'ID Number', 'First Name', 'Last Name', 'Full Name',
                    'Date/Time', 'Marked Period', 'Timestamp Processed'
                ]
                # Filter to only columns that actually exist in final_master_df
                final_master_df = final_master_df[[col for col in desired_columns_order if col in final_master_df.columns]]

                try:
                    writer = pd.ExcelWriter(master_file_path, engine='xlsxwriter')
                    final_master_df.to_excel(writer, sheet_name='Daily Report', index=False)
                    writer.close()
                    print(f"Daily master report updated: {master_file_path}")

                except Exception as master_report_error:
                    print(f"Error updating daily master report '{master_file_path}': {master_report_error}")

            if ids_to_paste:
                # --- PowerSchool Login for Raptor Automation ---
                if not powerschool_login(driver, wait, powerschool_username, powerschool_password):
                    print("Failed to login to PowerSchool. Exiting Raptor automation.")
                    return

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
                                print(f"    ❌ Error selecting checkbox {cb_name}: {e}")

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

    except SystemExit:
        print("Application exited by user during automation process.")
    except Exception as e:
        print(f"❌ An unexpected error occurred during Raptor automation: {e}")
        if driver:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            error_screenshot_name = f'raptor_automation_error_screenshot_{timestamp}.png'
            try:
                driver.save_screenshot(error_screenshot_name)
                print(f"📸 Screenshot '{error_screenshot_name}' saved for debugging.")
            except Exception as se:
                print(f"Could not save screenshot: {se}")
    finally:
        if driver:
            print("Closing the browser.")
            driver.quit()

# --- Main Menu Prompt ---
def main_menu_prompt():
    print("\n--- Main Menu ---")
    print("Select an operation:")
    print("  1. Consolidate Absences (from Meeting Attendance Report)")
    print("  2. Raptor Attendance (Daily Sign-in / Late Arrivals)")
    choice = safe_input("Enter choice (1 or 2)").strip()
    return choice

if __name__ == "__main__":
    print("🚀 Welcome to the EDNOVATE Attendance Automation Suite!")
    print("-----------------------------------------------------")

    # Ensure all necessary directories exist at startup
    os.makedirs(RAPTOR_REPORTS_DIR, exist_ok=True)
    os.makedirs(DAILY_MASTER_REPORTS_DIR, exist_ok=True)
    os.makedirs(MEET_ATTENDANCE_DIR, exist_ok=True)
    print(f"Ensured '{RAPTOR_REPORTS_DIR}', '{DAILY_MASTER_REPORTS_DIR}', and '{MEET_ATTENDANCE_DIR}' directories exist.")

    try:
        operation_choice = main_menu_prompt()

        if operation_choice == '1':
            print("\nStarting Consolidate Absences workflow...")
            consolidate_attendance()
        elif operation_choice == '2':
            print("\nStarting Raptor Attendance workflow...")
            # Prompt for M-Th schedule type first
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

            if not ACTUAL_F_PERIODS or len(ACTUAL_F_PERIODS) < 1:
                print("🚨 WARNING: Friday period schedule (ACTUAL_F_PERIODS) appears to be missing or empty in the script.")

            selected_day_key, selected_period, all_periods_for_day, filter_start, filter_end = get_user_day_and_period_selection(effective_periods_config)

            print(f"\n--- Script Configuration for this Run ---")
            print(f"Selected Day Type: {selected_day_key}")
            if selected_day_key in ['M','T','W','H']:#
                if active_m_th_schedule == M_TH_NORMAL_PERIODS:
                    print(f"M-Th Schedule Type: Normal")
                else:
                    print(f"M-Th Schedule Type: Enrichment")
            print(f"Target Period for UL: {selected_period['name']}")
            print(f"Excel Date/Time Filter: From {filter_start} to {filter_end}")
            print("--- Starting Automation ---")

            automate_raptor_and_powerschool(selected_period, all_periods_for_day, filter_start, filter_end)
        else:
            print("Invalid main menu choice. Exiting.")

    except SystemExit:
        print("Application exited by user.")
    except Exception as main_err:
        print(f"❌ An unexpected error occurred in the main execution block: {main_err}")
