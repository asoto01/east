# EAST

A small collection of automation scripts to streamline school operations. This repository simplifies generating reports from Raptor and automating attendance tasks in PowerSchool by leveraging API requests or Selenium for web interactions. Sensitive data like passwords is handled securely through user prompts, ensuring no credentials are stored in the repo.

## Repository Structure
- **attendance_automation/**: Contains scripts for automating attendance-related tasks.
  - `rosa_v_0_3.py`: A Python script for automating attendance processes (e.g., integrating with PowerSchool or similar systems). See the section below for details.
- `requirements.txt`: Lists the Python dependencies required for the scripts in this repository.
- Other files/directories: (Add more as needed based on your repo's contents, e.g., report generation scripts).

## Setup Instructions (for Mac OS)
To set up and run the scripts in this repository on macOS, we recommend using a virtual environment (`venv`) to isolate dependencies. This prevents conflicts with your system's Python packages.

### Prerequisites
- Python 3.x installed (macOS comes with Python, but ensure it's version 3.8+; you can install via Homebrew with `brew install python3` if needed).
- Git installed (for cloning the repo).

### Steps
1. **Clone the Repository**:
   ```
   git clone https://github.com/asoto01/east.git
   cd east
   ```

2. **Create and Activate a Virtual Environment**:
   Navigate to the root of the repository and create a virtual environment:
   ```
   python3 -m venv venv
   ```
   Activate it:
   ```
   source venv/bin/activate
   ```
   (You'll see `(venv)` in your terminal prompt when activated. To deactivate later, run `deactivate`.)

3. **Install Dependencies**:
   Install the required packages from `requirements.txt`:
   ```
   pip install -r requirements.txt
   ```

4. **Prepare Credentials for Scripts**:
   Some scripts, like those in `attendance_automation/`, require sensitive credentials. Do not commit these to the repo. Instead:
   - Create the necessary files in the script's directory (e.g., `credentials_json` for `rosa_v_0_3.py`).
   - Obtain your credentials (e.g., API keys or service account JSON) from the relevant service (e.g., Google Cloud, PowerSchool API).
   - Place the file securely in the directory without adding it to Git.

### Running rosa_v_0_3.py
This script is located in the `attendance_automation/` directory and automates attendance tasks (likely integrating with systems like PowerSchool). It requires a `credentials_json` file for authentication (e.g., a service account JSON for API access).

#### Setup Specific to rosa_v_0_3.py
1. Navigate to the directory:
   ```
   cd attendance_automation
   ```
2. Place your `credentials_json` file in this directory. This file should contain your API credentials in JSON format. Example structure (do not use this; generate your own):
   ```
   {
     "type": "service_account",
     "project_id": "your-project-id",
     "private_key_id": "your-private-key-id",
     // ... other fields ...
   }
   ```
   Ensure this file is not added to Git (add `credentials_json` to `.gitignore` if needed).

3. Run the script (from within the activated venv):
   ```
   python rosa_v_0_3.py
   ```
   - The script may prompt for additional inputs like passwords.
   - If it uses Selenium, ensure you have a compatible web driver (e.g., ChromeDriver) installed and in your PATH.

If the script encounters errors (e.g., missing modules), double-check that all dependencies from `requirements.txt` are installed. Update `requirements.txt` if new packages are needed (e.g., `pip freeze > requirements.txt` after installing).
