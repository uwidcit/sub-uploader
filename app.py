import os
import sys
import json
import csv
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QLabel, QLineEdit, QFileDialog, QVBoxLayout, QTextEdit, QFormLayout)
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request
from tqdm import tqdm

def extract_content_after_file(filename):
    """Remove 'assignsubmission_file_' from Moodle submission filenames."""
    # Pattern: Name_ID_assignsubmission_file_ActualFilename.ext
    # Remove only the 'assignsubmission_file_' part, keep name and ID
    if 'assignsubmission_file_' in filename:
        return filename.replace('assignsubmission_file_', '')
    return filename  # Return original if pattern not found

def load_config(config_file='config.json'):
    """Load configuration from JSON file."""
    try:
        with open(config_file, 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        # Return default configuration if file doesn't exist
        return {
            "google_sheets": {
                "sheet_id": "",
                "sheet_name": "Sheet1",
                "id_column": "A",
                "first_name_column": "",
                "last_name_column": "",
                "link_column": "L",
                "start_row": 3
            },
            "google_drive": {
                "folder_id": ""
            },
            "submissions": {
                "folder_path": ""
            },
            "authentication": {
                "scopes": [
                    "https://www.googleapis.com/auth/drive.file",
                    "https://www.googleapis.com/auth/spreadsheets"
                ],
                "credentials_file": "credentials.json",
                "token_file": "token.json"
            },
            "output": {
                "summary_file": "upload_summary.txt"
            },
            "upload": {
                "mime_type": "application/octet-stream",
                "permissions": {
                    "role": "reader",
                    "type": "anyone"
                }
            }
        }
    except json.JSONDecodeError:
        print(f"Invalid JSON in configuration file '{config_file}'. Using defaults.")
        return load_config()  # Return default config if JSON is invalid

def save_config(config, config_file='config.json'):
    """Save configuration to JSON file."""
    try:
        with open(config_file, 'w') as f:
            json.dump(config, f, indent=2)
        return True
    except Exception as e:
        print(f"Error saving configuration: {e}")
        return False

def load_groups_from_csv(groups_file='groups.csv'):
    """Load student to group mappings from CSV file."""
    groups = {}
    if not os.path.exists(groups_file):
        return groups
    
    try:
        with open(groups_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                # Handle new format: First name, Last name, Group Name
                first_name = row.get('First name', '').strip()
                last_name = row.get('Last name', '').strip()
                group_name = row.get('Group Name', '').strip()
                
                # Handle space in column name
                if not group_name:
                    group_name = row.get(' Group Name', '').strip()
                
                # Try alternative column names if primary ones don't exist
                if not first_name:
                    first_name = row.get('first_name', '').strip()
                if not last_name:
                    last_name = row.get('last_name', '').strip()
                if not group_name:
                    group_name = row.get('group_name', '').strip()
                    if not group_name:
                        group_name = row.get('group', '').strip()
                
                # Handle [] as a valid group name, but filter out truly empty values
                if (first_name or last_name) and group_name and group_name not in ['', ' ']:
                    # Create full name for matching
                    full_name = f"{first_name} {last_name}".strip()
                    groups[full_name.lower()] = group_name
                    
                    # Also store first name only for partial matching
                    if first_name:
                        groups[first_name.lower()] = group_name
                    
                    # Also store last name only for partial matching
                    if last_name:
                        groups[last_name.lower()] = group_name
                    
        return groups
    except Exception as e:
        print(f"Error loading groups file: {e}")
        return {}

def extract_student_name_from_filename(filename):
    """Extract student name from Moodle submission filename."""
    # Pattern: StudentName_ID_assignsubmission_file_...
    if '_' in filename:
        parts = filename.split('_')
        if len(parts) > 2 and 'assignsubmission' in filename:
            # First part should be the student name
            return parts[0].strip()
    return None

def find_group_by_student_name(student_name, groups_data):
    """Find group name for a student using the groups data."""
    if not student_name or not groups_data:
        return None
    
    # Normalize the student name for comparison
    normalized_name = student_name.lower().strip()
    
    # Try exact match first
    if normalized_name in groups_data:
        return groups_data[normalized_name]
    
    # Try partial matches (last name, first name, etc.)
    for group_student, group_name in groups_data.items():
        if normalized_name in group_student or group_student in normalized_name:
            return group_name
    
    return None

def load_matches_from_csv(matches_file='matches.csv'):
    """Load filename to group name matches from CSV file."""
    matches = {}
    if not os.path.exists(matches_file):
        return matches
    
    try:
        with open(matches_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                filename = row.get('filename', '').strip()
                matched_id = row.get('matched_id', '').strip()
                similarity = float(row.get('similarity_score', '0'))
                
                # Only use matches with good similarity (>= 0.7)
                if filename and matched_id and matched_id != 'NO MATCH' and similarity >= 0.7:
                    matches[filename] = matched_id
                    
        return matches
    except Exception as e:
        print(f"Error loading matches file: {e}")
        return {}

SCOPES = ['https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/spreadsheets']

def extract_student_id(filename):
    """
    Extract student ID from the new filename format:
    StudentName_SubmissionID_assignsubmission_file_StudentID_COMP1600_A1.pdf
    Returns the StudentID (816xxxxxx) if found, None otherwise.
    """
    import re
    
    # Pattern to match the new naming convention
    # Look for pattern: assignsubmission_file_[student_id]_
    pattern = r'assignsubmission_file_(\d{9})_'
    match = re.search(pattern, filename)
    
    if match:
        return match.group(1)
    
    # Fallback: try to find any 9-digit number that starts with 816, 320, or 400
    fallback_pattern = r'(816\d{6}|320\d{6}|400\d{6})'
    fallback_match = re.search(fallback_pattern, filename)
    
    if fallback_match:
        return fallback_match.group(1)
    
    return None

def extract_names(filename):
    """
    Extract first and last names from filename.
    Format: FirstName LastName_SubmissionID_assignsubmission_file_StudentID_COMP1600_A1.pdf
    Returns tuple (first_name, last_name) or (None, None) if not found.
    """
    import re
    
    # Pattern to match the beginning of the filename before the first underscore
    # This captures the "FirstName LastName" part
    pattern = r'^([^_]+)_\d+_assignsubmission_file_'
    match = re.search(pattern, filename)
    
    if match:
        full_name = match.group(1).strip()
        # Split by space and assume first word is first name, rest is last name
        name_parts = full_name.split()
        if len(name_parts) >= 2:
            first_name = name_parts[0]
            last_name = ' '.join(name_parts[1:])  # Handle multiple last names
            return (first_name, last_name)
        elif len(name_parts) == 1:
            # Only one name provided
            return (name_parts[0], None)
    
    return (None, None)

def normalize_name(name):
    """
    Normalize a name for comparison by removing extra spaces, 
    converting to lowercase, and handling common variations.
    """
    import re
    
    if not name:
        return ""
    
    # Convert to lowercase and strip whitespace
    normalized = name.lower().strip()
    
    # Remove extra spaces and hyphens for comparison
    normalized = re.sub(r'[-\s]+', ' ', normalized)
    
    return normalized

def find_match_by_names(first_name, last_name, sheet_first_names, sheet_last_names):
    """
    Find a match by comparing first and last names.
    Returns the index if found, None otherwise.
    Handles cases where sheet columns might be empty.
    """
    if not first_name:
        return None
    
    # Handle cases where name columns might not exist or be empty
    if not sheet_first_names:
        return None
        
    norm_first = normalize_name(first_name)
    norm_last = normalize_name(last_name) if last_name else ""
    
    for i, sheet_first in enumerate(sheet_first_names):
        sheet_first_norm = normalize_name(sheet_first) if sheet_first else ""
        
        # Get corresponding last name if available
        sheet_last_norm = ""
        if sheet_last_names and i < len(sheet_last_names):
            sheet_last_norm = normalize_name(sheet_last_names[i]) if sheet_last_names[i] else ""
        
        # Skip empty rows
        if not sheet_first_norm and not sheet_last_norm:
            continue
            
        # Try exact match first
        if sheet_first_norm and norm_first == sheet_first_norm:
            if not norm_last or not sheet_last_norm or norm_last == sheet_last_norm:
                return i
                
        # Try first name + partial last name match
        if sheet_first_norm and norm_first == sheet_first_norm and norm_last and sheet_last_norm:
            if norm_last in sheet_last_norm or sheet_last_norm in norm_last:
                return i
    
    return None

def find_match_by_group_name(filename, sheet_ids):
    """
    Find a match by comparing group names from filename with ID column entries.
    Returns the index if found, None otherwise.
    """
    import os
    import re
    
    if not sheet_ids or not filename:
        return None
    
    # Use stripped filename for better matching
    stripped_filename = extract_content_after_file(filename)
    
    # Extract potential group names from stripped filename
    # Remove file extension and common patterns
    base_name = os.path.splitext(stripped_filename)[0]
    
    # Look for patterns like "Group 1", "Team A", "GroupName", etc.
    group_patterns = [
        r'group[\s_-]*(\w+)',
        r'team[\s_-]*(\w+)', 
        r'(\w*group\w*)',
        r'(\w*team\w*)',
    ]
    
    potential_groups = []
    for pattern in group_patterns:
        matches = re.finditer(pattern, base_name, re.IGNORECASE)
        for match in matches:
            if match.group(1):
                potential_groups.append(match.group(1).strip())
            potential_groups.append(match.group(0).strip())
    
    # Also try the whole filename without extension as a potential group name
    potential_groups.append(base_name)
    
    # Normalize potential group names
    norm_groups = [normalize_name(group) for group in potential_groups if group]
    
    # Try to match against ID column entries
    for i, sheet_id in enumerate(sheet_ids):
        if not sheet_id:
            continue
            
        norm_sheet_id = normalize_name(str(sheet_id))
        
        for norm_group in norm_groups:
            if not norm_group:
                continue
                
            # Try exact match
            if norm_group == norm_sheet_id:
                return i
                
            # Try partial match (both directions)
            if len(norm_group) > 2 and len(norm_sheet_id) > 2:
                if norm_group in norm_sheet_id or norm_sheet_id in norm_group:
                    return i
    
    return None

def resource_path(relative_path):
    """ Get the absolute path to a resource, works for dev and PyInstaller. """
    try:
        # PyInstaller stores the path in sys._MEIPASS when packaged
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

credentials_path = resource_path('credentials.json')

class FileUploaderApp(QWidget):
    def __init__(self):
        super().__init__()
        self.config = load_config()
        self.folder_path = None
        self.creds = None
        
        # Initialize UI first
        self.initUI()
        
        # Load config into UI fields (no auto-save signals)
        self.load_config_to_ui()
        
        # Check authentication status
        self.check_token()

    def initUI(self):
        self.setWindowTitle("Google Drive File Uploader")
        self.setGeometry(100, 100, 600, 600)

        layout = QVBoxLayout()
        form_layout = QFormLayout()

        # Input fields for parameters
        self.sheet_id_input = QLineEdit(self)
        self.sheet_name_input = QLineEdit(self)
        self.id_column_input = QLineEdit(self)
        self.first_name_column_input = QLineEdit(self)
        self.last_name_column_input = QLineEdit(self)
        self.link_column_input = QLineEdit(self)
        self.start_row_input = QLineEdit(self)
        self.folder_id_input = QLineEdit(self)
        self.submissions_folder_input = QLineEdit(self)

        # No auto-save connections - manual save only

        form_layout.addRow(QLabel("Sheet ID:"), self.sheet_id_input)
        form_layout.addRow(QLabel("Sheet Name:"), self.sheet_name_input)
        form_layout.addRow(QLabel("ID Column (Column with student IDs):"), self.id_column_input)
        form_layout.addRow(QLabel("First Name Column:"), self.first_name_column_input)
        form_layout.addRow(QLabel("Last Name Column:"), self.last_name_column_input)
        form_layout.addRow(QLabel("Link Column (Where submission file link will be placed):"), self.link_column_input)
        form_layout.addRow(QLabel("Start Row (Row where data starts):"), self.start_row_input)
        form_layout.addRow(QLabel("Google Drive Folder ID:"), self.folder_id_input)
        
        # Submissions folder section with button between label and text field
        submissions_label = QLabel("Submissions Folder Path:")
        self.select_folder_button = QPushButton('Select Folder', self)
        self.select_folder_button.clicked.connect(self.open_folder_dialog)
        form_layout.addRow(submissions_label, self.select_folder_button)
        form_layout.addRow(QLabel(""), self.submissions_folder_input)

        layout.addLayout(form_layout)

        # Button to authorize app (moved to top)
        self.auth_button = QPushButton('Authorize App', self)
        self.auth_button.clicked.connect(self.authorize_app)
        layout.addWidget(self.auth_button)

        # Token status label
        self.token_status_label = QLabel('Token Status: Not Authorized', self)
        layout.addWidget(self.token_status_label)

        # Save Config button
        self.save_config_button = QPushButton('Save Configuration', self)
        self.save_config_button.clicked.connect(self.save_config_from_ui)
        layout.addWidget(self.save_config_button)

        # Upload button
        self.upload_button = QPushButton('Start Upload', self)
        self.upload_button.clicked.connect(self.start_upload)
        layout.addWidget(self.upload_button)

        # Log output box
        self.log_output = QTextEdit(self)
        self.log_output.setReadOnly(True)
        layout.addWidget(self.log_output)

        self.setLayout(layout)

    def log(self, message):
        self.log_output.append(message)

    def load_config_to_ui(self):
        """Load configuration values into UI fields."""
        sheets_config = self.config.get('google_sheets', {})
        drive_config = self.config.get('google_drive', {})
        submissions_config = self.config.get('submissions', {})
        
        # Load values from config
        sheet_id = sheets_config.get('sheet_id', '')
        sheet_name = sheets_config.get('sheet_name', '')
        id_column = sheets_config.get('id_column', '')
        first_name_column = sheets_config.get('first_name_column', '')
        last_name_column = sheets_config.get('last_name_column', '')
        link_column = sheets_config.get('link_column', '')
        start_row = sheets_config.get('start_row', 3)
        folder_id = drive_config.get('folder_id', '')
        submissions_folder = submissions_config.get('folder_path', '')
        
        # Set the UI fields with the loaded values
        self.sheet_id_input.setText(sheet_id)
        self.sheet_name_input.setText(sheet_name)
        self.id_column_input.setText(id_column)
        self.first_name_column_input.setText(first_name_column)
        self.last_name_column_input.setText(last_name_column)
        self.link_column_input.setText(link_column)
        self.start_row_input.setText(str(start_row))
        self.folder_id_input.setText(folder_id)
        self.submissions_folder_input.setText(submissions_folder)
        
        # Log the loading for debugging
        self.log("✅ Configuration loaded from config.json")
        if sheet_id:
            self.log(f"  Sheet ID: {sheet_id}")
        if sheet_name:
            self.log(f"  Sheet Name: {sheet_name}")
        if folder_id:
            self.log(f"  Folder ID: {folder_id[:20]}..." if len(folder_id) > 20 else f"  Folder ID: {folder_id}")
        if submissions_folder:
            self.log(f"  Submissions Folder: {submissions_folder}")

    def save_config_from_ui(self):
        """Save UI values to configuration file."""
        try:
            # Update configuration with current UI values
            self.config['google_sheets']['sheet_id'] = self.sheet_id_input.text().strip()
            self.config['google_sheets']['sheet_name'] = self.sheet_name_input.text().strip()
            self.config['google_sheets']['id_column'] = self.id_column_input.text().strip()
            self.config['google_sheets']['first_name_column'] = self.first_name_column_input.text().strip()
            self.config['google_sheets']['last_name_column'] = self.last_name_column_input.text().strip()
            self.config['google_sheets']['link_column'] = self.link_column_input.text().strip()
            
            # Handle start_row conversion
            try:
                self.config['google_sheets']['start_row'] = int(self.start_row_input.text().strip()) if self.start_row_input.text().strip() else 3
            except ValueError:
                self.config['google_sheets']['start_row'] = 3
                self.log("⚠ Invalid start row value, defaulting to 3")
                
            self.config['google_drive']['folder_id'] = self.folder_id_input.text().strip()
            
            # Ensure submissions section exists
            if 'submissions' not in self.config:
                self.config['submissions'] = {}
            self.config['submissions']['folder_path'] = self.submissions_folder_input.text().strip()
            
            # Save to file
            if save_config(self.config):
                self.log("✅ Configuration saved to config.json")
            else:
                self.log("❌ Failed to save configuration")
                
        except Exception as e:
            self.log(f"✗ Error saving configuration: {e}")

    def open_folder_dialog(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select Submissions Folder')
        if folder:
            self.folder_path = folder
            
            # Update the submissions folder input field and config
            self.submissions_folder_input.setText(folder)
            
            # Ensure submissions section exists in config
            if 'submissions' not in self.config:
                self.config['submissions'] = {}
            self.config['submissions']['folder_path'] = folder
            
            self.log(f"📁 Submissions folder selected: {folder}")

    def check_token(self):
        """Check if credentials.json and token.json exist and are valid."""
        credentials_file = self.config['authentication'].get('credentials_file', 'credentials.json')
        token_file = self.config['authentication'].get('token_file', 'token.json')
        scopes = self.config['authentication'].get('scopes', SCOPES)
        
        credentials_path = resource_path(credentials_file)
        
        # First check if credentials file exists
        if not os.path.exists(credentials_path):
            self.token_status_label.setText('❌ No credentials.json file found')
            self.token_status_label.setStyleSheet("color: red;")
            self.log("❌ Error: credentials.json file not found. Please add your Google API credentials file.")
            self.auth_button.setEnabled(False)
            return
        
        self.auth_button.setEnabled(True)
        
        if os.path.exists(token_file):
            try:
                self.creds = Credentials.from_authorized_user_file(token_file, scopes)
                if not self.creds or not self.creds.valid:
                    if self.creds and self.creds.expired and self.creds.refresh_token:
                        try:
                            self.creds.refresh(Request())
                            self.token_status_label.setText('✅ Token Status: Valid (Refreshed)')
                            self.token_status_label.setStyleSheet("color: green;")
                            self.log("✅ Token refreshed successfully.")
                        except Exception as e:
                            self.token_status_label.setText('⚠️ Token Status: Expired, Reauthorization Needed')
                            self.token_status_label.setStyleSheet("color: orange;")
                            self.log("⚠️ Token refresh failed: " + str(e))
                    else:
                        self.token_status_label.setText('⚠️ Token Status: Expired or Invalid')
                        self.token_status_label.setStyleSheet("color: orange;")
                        self.log("⚠️ Token is expired or invalid. Please reauthorize the app.")
                else:
                    self.token_status_label.setText('✅ Token Status: Valid')
                    self.token_status_label.setStyleSheet("color: green;")
                    self.log("✅ Token is valid and ready to use.")
            except Exception as e:
                self.token_status_label.setText('❌ Token Status: Error Loading Token')
                self.token_status_label.setStyleSheet("color: red;")
                self.log(f"❌ Error loading token: {str(e)}")
        else:
            self.token_status_label.setText('⚠️ Token Status: No Token Found - Authorization Required')
            self.token_status_label.setStyleSheet("color: orange;")
            self.log("⚠️ No token found. Please authorize the app to access Google APIs.")

    def authorize_app(self):
        """Handles the Google OAuth authorization process."""
        try:
            credentials_file = self.config['authentication'].get('credentials_file', 'credentials.json')
            token_file = self.config['authentication'].get('token_file', 'token.json')
            scopes = self.config['authentication'].get('scopes', SCOPES)
            
            credentials_path = resource_path(credentials_file)
            
            if not os.path.exists(credentials_path):
                self.log("❌ Error: credentials.json file not found. Please add your Google API credentials file.")
                self.token_status_label.setText('❌ No credentials.json file found')
                self.token_status_label.setStyleSheet("color: red;")
                return
            
            self.log("🔄 Starting authorization process...")
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, scopes)
            self.creds = flow.run_local_server(port=0)
            
            # Save the credentials for the next run
            with open(token_file, 'w') as token:
                token.write(self.creds.to_json())
            self.token_status_label.setText('✅ Token Status: Authorized')
            self.token_status_label.setStyleSheet("color: green;")
            self.log("✅ App authorized successfully! Token saved.")
        except FileNotFoundError:
            error_msg = f"❌ Credentials file not found: {credentials_file}"
            self.log(error_msg)
            self.token_status_label.setText('❌ Credentials file missing')
            self.token_status_label.setStyleSheet("color: red;")
        except Exception as e:
            error_msg = f"❌ Authorization failed: {str(e)}"
            self.log(error_msg)
            self.token_status_label.setText('❌ Authorization Failed')
            self.token_status_label.setStyleSheet("color: red;")

    def start_upload(self):
        # Check if we have a folder path - prefer the configured submissions folder
        submissions_folder = self.submissions_folder_input.text().strip()
        upload_folder = submissions_folder if submissions_folder else self.folder_path
        
        if not upload_folder:
            self.log("❌ Please select a submissions folder or configure the submissions folder path before uploading.")
            return
        if not self.creds or not self.creds.valid:
            self.log("❌ App is not authorized. Please click 'Authorize App' button first.")
            return

        self.log("🚀 Starting upload process...")
        
        # Retrieve user input
        sheet_id = self.sheet_id_input.text().strip()
        sheet_name = self.sheet_name_input.text().strip()
        id_column = self.id_column_input.text().strip()
        first_name_column = self.first_name_column_input.text().strip()
        last_name_column = self.last_name_column_input.text().strip()
        link_column = self.link_column_input.text().strip()
        start_row = int(self.start_row_input.text().strip())
        folder_id = self.folder_id_input.text().strip()

        self.log(f"📁 Uploading files from: {upload_folder}")
        self.upload_files(sheet_id, sheet_name, id_column, first_name_column, last_name_column, link_column, start_row, folder_id, upload_folder)

    def upload_files(self, sheet_id, sheet_name, id_column, first_name_column, last_name_column, link_column, start_row, folder_id, upload_folder):
        drive_service = build('drive', 'v3', credentials=self.creds)
        sheets_service = build('sheets', 'v4', credentials=self.creds)

        # Define ranges for Google Sheets
        id_range = f"{sheet_name}!{id_column}{start_row}:{id_column}"
        first_name_range = f"{sheet_name}!{first_name_column}{start_row}:{first_name_column}" if first_name_column else None
        last_name_range = f"{sheet_name}!{last_name_column}{start_row}:{last_name_column}" if last_name_column else None
        link_range = f"{sheet_name}!{link_column}{start_row}:{link_column}"

        # Fetch all IDs and Links from Google Sheets
        result_ids = sheets_service.spreadsheets().values().get(spreadsheetId=sheet_id, range=id_range).execute()
        ids = [item[0] if item else "" for item in result_ids.get('values', [])]

        # Fetch names only if columns are specified
        first_names = []
        last_names = []
        
        if first_name_range:
            try:
                result_first_names = sheets_service.spreadsheets().values().get(spreadsheetId=sheet_id, range=first_name_range).execute()
                first_names = [item[0] if item else "" for item in result_first_names.get('values', [])]
            except Exception as e:
                self.log(f"Warning: Could not fetch first names: {e}")
                first_names = []
        
        if last_name_range:
            try:
                result_last_names = sheets_service.spreadsheets().values().get(spreadsheetId=sheet_id, range=last_name_range).execute()
                last_names = [item[0] if item else "" for item in result_last_names.get('values', [])]
            except Exception as e:
                self.log(f"Warning: Could not fetch last names: {e}")
                last_names = []

        result_links = sheets_service.spreadsheets().values().get(spreadsheetId=sheet_id, range=link_range).execute()
        links = [item[0] for item in result_links.get('values', []) if item]

        # Upload stats and progress tracking
        total_files_in_directory = len(os.listdir(upload_folder))
        files_uploaded_successfully = 0
        files_failed_to_upload = []
        skipped_files = []
        uploaded_ids = []  # Track which IDs/groups were uploaded

        for filename in os.listdir(upload_folder):
            file_path = os.path.join(upload_folder, filename)
            if os.path.isfile(file_path):
                # Extract student ID from the new filename format
                file_id = extract_student_id(filename)
                
                # Extract names from filename
                first_name, last_name = extract_names(filename)
                
                # Safe logging for Unicode filenames
                safe_filename = filename.encode('ascii', 'replace').decode('ascii')
                self.log(f"Processing file: {safe_filename}")
                self.log(f"Extracted student ID: {file_id}")
                self.log(f"Extracted names: {first_name} {last_name}")
                
                row_index = None
                match_method = ""
                
                # Check for group mode
                groups_data = load_groups_from_csv()
                GROUP_MODE = len(groups_data) > 0
                
                # Load pre-computed matches from CSV
                filename_matches = load_matches_from_csv()
                group_name_to_write = None
                
                if GROUP_MODE:
                    # In group mode, use matches.csv to find the matched group name
                    if filename in filename_matches:
                        matched_group = filename_matches[filename]
                        try:
                            # Find the row where the ID column matches the group name
                            row_index = ids.index(matched_group)
                            match_method = f"group mode CSV match: {matched_group}"
                            group_name_to_write = None  # Don't overwrite existing group name
                        except ValueError:
                            # Group name from matches.csv not found in ID column - create new row
                            # Find the first empty row (where ID column is empty)
                            for i, id_entry in enumerate(ids):
                                if not id_entry or id_entry.strip() == "":
                                    row_index = i
                                    match_method = f"group mode new row: {matched_group}"
                                    group_name_to_write = matched_group
                                    # Update the ids list to mark this row as occupied
                                    ids[i] = matched_group
                                    self.log(f"Creating new row for group '{matched_group}' at row {i + start_row}")
                                    break
                            
                            if row_index is None:
                                # No empty rows found, extend the list
                                row_index = len(ids)
                                ids.append(matched_group)  # Add the group name to extend the list
                                match_method = f"group mode extended row: {matched_group}"
                                group_name_to_write = matched_group
                                self.log(f"Extending spreadsheet for group '{matched_group}' at row {row_index + start_row}")
                    else:
                        self.log(f"No match found in matches.csv for: {safe_filename}")
                else:
                    # Standard mode - First check pre-computed matches from CSV
                    if filename in filename_matches:
                        matched_group = filename_matches[filename]
                        try:
                            row_index = ids.index(matched_group)
                            match_method = f"CSV match: {matched_group}"
                        except ValueError:
                            # Group name from CSV not found in current spreadsheet
                            pass
                
                # If no CSV match, try to match by student ID
                if row_index is None and file_id:
                    try:
                        row_index = ids.index(file_id)
                        match_method = f"student ID {file_id}"
                    except ValueError:
                        # ID not found in list
                        pass
                
                # If no student ID match, try name matching (only if name columns are available)
                if row_index is None and first_name and (first_names or last_names):
                    name_match_index = find_match_by_names(first_name, last_name, first_names, last_names)
                    if name_match_index is not None:
                        row_index = name_match_index
                        match_method = f"name match: {first_name} {last_name}"
                
                # If still no match, try group name matching against ID column
                if row_index is None:
                    group_match_index = find_match_by_group_name(filename, ids)
                    if group_match_index is not None:
                        row_index = group_match_index
                        match_method = f"group name match with ID: {ids[group_match_index]}"
                
                if row_index is None:
                    skipped_files.append(filename)
                    if file_id:
                        self.log(f"Student ID {file_id} not found in spreadsheet and no name match found")
                    else:
                        self.log(f"No valid student ID found in filename and no name match: {safe_filename}")
                    continue

                self.log(f"Matched by {match_method}, row {row_index + start_row}")

                # Check if a link already exists for this row
                if row_index < len(links) and links[row_index].strip():
                    self.log(f"Link already exists for row {row_index + start_row}, skipping...")
                    continue  # Skip uploading if link already exists

                try:
                    # Upload to Google Drive
                    media = MediaFileUpload(file_path, resumable=True)
                    mime_type = self.config['upload'].get('mime_type', 'application/octet-stream')
                    file_metadata = {'name': filename, 'mimeType': mime_type, 'parents': [folder_id]}
                    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()

                    # Share and link
                    permissions = self.config['upload'].get('permissions', {'role': 'reader', 'type': 'anyone'})
                    drive_service.permissions().create(fileId=file['id'], body=permissions).execute()
                    link = f"https://drive.google.com/file/d/{file['id']}/view"
                    hyperlink_formula = f'=HYPERLINK("{link}", "Open File")'
                    row_num = row_index + start_row
                    update_range = f"{sheet_name}!{link_column}{row_num}"
                    values = [[hyperlink_formula]]
                    body = {'values': values}
                    sheets_service.spreadsheets().values().update(spreadsheetId=sheet_id, range=update_range, valueInputOption="USER_ENTERED", body=body).execute()
                    
                    # In group mode, write the group name to ID column for new rows
                    if GROUP_MODE and group_name_to_write:
                        id_update_range = f"{sheet_name}!{id_column}{row_num}"
                        id_values = [[group_name_to_write]]
                        id_body = {'values': id_values}
                        sheets_service.spreadsheets().values().update(spreadsheetId=sheet_id, range=id_update_range, valueInputOption="USER_ENTERED", body=id_body).execute()
                        self.log(f"Created new row with group name: {group_name_to_write}")
                    elif GROUP_MODE:
                        self.log(f"Using existing group row: {ids[row_index] if row_index < len(ids) else 'Unknown'}")
                    
                    files_uploaded_successfully += 1
                    
                    # Track the uploaded ID/group name
                    uploaded_id = group_name_to_write if GROUP_MODE and group_name_to_write else (ids[row_index] if row_index < len(ids) else "Unknown")
                    uploaded_ids.append(f"{uploaded_id} (row {row_index + start_row})")
                    
                    safe_filename = filename.encode('ascii', 'replace').decode('ascii')
                    self.log(f"✓ Uploaded: {safe_filename} -> {match_method}")

                except Exception as e:
                    files_failed_to_upload.append((filename, str(e)))
                    safe_filename = filename.encode('ascii', 'replace').decode('ascii')
                    self.log(f"✗ Failed to upload {safe_filename}: {str(e)}")

        # Summary output
        self.log(f"Total files: {total_files_in_directory}")
        self.log(f"Files uploaded successfully: {files_uploaded_successfully}")
        self.log(f"Files skipped: {len(skipped_files)}")
        
        # List uploaded IDs/groups
        if uploaded_ids:
            self.log("\nIDs/Groups Uploaded:")
            self.log("-------------------")
            for uploaded_id in uploaded_ids:
                self.log(f"✓ {uploaded_id}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    uploader = FileUploaderApp()
    uploader.show()
    sys.exit(app.exec_())
