import os
import sys
import json
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QLabel, QLineEdit, QFileDialog, QVBoxLayout, QTextEdit, QFormLayout)
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request
from tqdm import tqdm

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
                "first_name_column": "B",
                "last_name_column": "C",
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
    """
    if not first_name or not sheet_first_names:
        return None
        
    norm_first = normalize_name(first_name)
    norm_last = normalize_name(last_name) if last_name else ""
    
    for i, (sheet_first, sheet_last) in enumerate(zip(sheet_first_names, sheet_last_names)):
        sheet_first_norm = normalize_name(sheet_first) if sheet_first else ""
        sheet_last_norm = normalize_name(sheet_last) if sheet_last else ""
        
        # Try exact match first
        if norm_first == sheet_first_norm and norm_last == sheet_last_norm:
            return i
            
        # Try first name + partial last name match
        if norm_first == sheet_first_norm and norm_last and sheet_last_norm:
            if norm_last in sheet_last_norm or sheet_last_norm in norm_last:
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
        first_name_column = sheets_config.get('first_name_column', 'B')
        last_name_column = sheets_config.get('last_name_column', 'C')
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
        first_name_range = f"{sheet_name}!{first_name_column}{start_row}:{first_name_column}"
        last_name_range = f"{sheet_name}!{last_name_column}{start_row}:{last_name_column}"
        link_range = f"{sheet_name}!{link_column}{start_row}:{link_column}"

        # Fetch all IDs, Names, and Links from Google Sheets
        result_ids = sheets_service.spreadsheets().values().get(spreadsheetId=sheet_id, range=id_range).execute()
        ids = [item[0] for item in result_ids.get('values', []) if item]

        result_first_names = sheets_service.spreadsheets().values().get(spreadsheetId=sheet_id, range=first_name_range).execute()
        first_names = [item[0] if item else "" for item in result_first_names.get('values', [])]
        
        result_last_names = sheets_service.spreadsheets().values().get(spreadsheetId=sheet_id, range=last_name_range).execute()
        last_names = [item[0] if item else "" for item in result_last_names.get('values', [])]

        result_links = sheets_service.spreadsheets().values().get(spreadsheetId=sheet_id, range=link_range).execute()
        links = [item[0] for item in result_links.get('values', []) if item]

        # Upload stats and progress tracking
        total_files_in_directory = len(os.listdir(upload_folder))
        files_uploaded_successfully = 0
        files_failed_to_upload = []
        skipped_files = []

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
                
                # First try to match by student ID
                if file_id and file_id in ids:
                    row_index = ids.index(file_id)
                    match_method = f"student ID {file_id}"
                
                # If no student ID match, try name matching
                elif first_name:
                    name_match_index = find_match_by_names(first_name, last_name, first_names, last_names)
                    if name_match_index is not None:
                        row_index = name_match_index
                        match_method = f"name match: {first_name} {last_name}"
                
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
                    files_uploaded_successfully += 1
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


if __name__ == '__main__':
    app = QApplication(sys.argv)
    uploader = FileUploaderApp()
    uploader.show()
    sys.exit(app.exec_())
