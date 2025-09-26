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
                "id_column": "C",
                "link_column": "L",
                "start_row": 3
            },
            "google_drive": {
                "folder_id": ""
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
        self.initUI()
        self.load_config_to_ui()
        self.folder_path = None
        self.creds = None
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
        self.link_column_input = QLineEdit(self)
        self.start_row_input = QLineEdit(self)
        self.folder_id_input = QLineEdit(self)

        # Connect input fields to auto-save function
        self.sheet_id_input.textChanged.connect(self.save_config_from_ui)
        self.sheet_name_input.textChanged.connect(self.save_config_from_ui)
        self.id_column_input.textChanged.connect(self.save_config_from_ui)
        self.link_column_input.textChanged.connect(self.save_config_from_ui)
        self.start_row_input.textChanged.connect(self.save_config_from_ui)
        self.folder_id_input.textChanged.connect(self.save_config_from_ui)

        form_layout.addRow(QLabel("Sheet ID:"), self.sheet_id_input)
        form_layout.addRow(QLabel("Sheet Name:"), self.sheet_name_input)
        form_layout.addRow(QLabel("ID Column (Column with file IDs):"), self.id_column_input)
        form_layout.addRow(QLabel("Link Column (Where submission file link will be placed):"), self.link_column_input)
        form_layout.addRow(QLabel("Start Row (Row where data starts):"), self.start_row_input)
        form_layout.addRow(QLabel("Google Drive Folder ID:"), self.folder_id_input)

        layout.addLayout(form_layout)

        # Token status label
        self.token_status_label = QLabel('Token Status: Not Authorized', self)
        layout.addWidget(self.token_status_label)

        # Button to authorize app
        self.auth_button = QPushButton('Authorize App', self)
        self.auth_button.clicked.connect(self.authorize_app)
        layout.addWidget(self.auth_button)

        # Folder selection button and label
        self.select_folder_button = QPushButton('Select Folder', self)
        self.select_folder_button.clicked.connect(self.open_folder_dialog)
        layout.addWidget(self.select_folder_button)

        self.folder_label = QLabel('Folder: Not selected', self)
        layout.addWidget(self.folder_label)

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
        
        self.sheet_id_input.setText(sheets_config.get('sheet_id', ''))
        self.sheet_name_input.setText(sheets_config.get('sheet_name', 'Sheet1'))
        self.id_column_input.setText(sheets_config.get('id_column', 'C'))
        self.link_column_input.setText(sheets_config.get('link_column', 'L'))
        self.start_row_input.setText(str(sheets_config.get('start_row', 3)))
        self.folder_id_input.setText(drive_config.get('folder_id', ''))

    def save_config_from_ui(self):
        """Save UI values to configuration file."""
        try:
            # Update configuration with current UI values
            self.config['google_sheets']['sheet_id'] = self.sheet_id_input.text().strip()
            self.config['google_sheets']['sheet_name'] = self.sheet_name_input.text().strip()
            self.config['google_sheets']['id_column'] = self.id_column_input.text().strip()
            self.config['google_sheets']['link_column'] = self.link_column_input.text().strip()
            
            # Handle start_row conversion
            try:
                self.config['google_sheets']['start_row'] = int(self.start_row_input.text().strip()) if self.start_row_input.text().strip() else 3
            except ValueError:
                self.config['google_sheets']['start_row'] = 3
                
            self.config['google_drive']['folder_id'] = self.folder_id_input.text().strip()
            
            # Save to file
            save_config(self.config)
        except Exception as e:
            self.log(f"Error saving configuration: {e}")

    def open_folder_dialog(self):
        folder = QFileDialog.getExistingDirectory(self, 'Select Folder')
        if folder:
            self.folder_path = folder
            self.folder_label.setText(f"Folder: {folder}")
        else:
            self.folder_label.setText('Folder: Not selected')

    def check_token(self):
        """Check if token.json exists and is valid."""
        token_file = self.config['authentication'].get('token_file', 'token.json')
        scopes = self.config['authentication'].get('scopes', SCOPES)
        
        if os.path.exists(token_file):
            self.creds = Credentials.from_authorized_user_file(token_file, scopes)
            if not self.creds or not self.creds.valid:
                if self.creds and self.creds.expired and self.creds.refresh_token:
                    try:
                        self.creds.refresh(Request())
                        self.token_status_label.setText('Token Status: Valid (Refreshed)')
                        self.log("Token refreshed successfully.")
                    except Exception as e:
                        self.token_status_label.setText('Token Status: Expired, Reauthorization Needed')
                        self.log("Token refresh failed: " + str(e))
                else:
                    self.token_status_label.setText('Token Status: Expired or Invalid')
            else:
                self.token_status_label.setText('Token Status: Valid')
                self.log("Token is valid.")
        else:
            self.token_status_label.setText('Token Status: No Token Found')
            self.log("No token found, please authorize the app.")

    def authorize_app(self):
        """Handles the Google OAuth authorization process."""
        try:
            credentials_file = self.config['authentication'].get('credentials_file', 'credentials.json')
            token_file = self.config['authentication'].get('token_file', 'token.json')
            scopes = self.config['authentication'].get('scopes', SCOPES)
            
            credentials_path = resource_path(credentials_file)
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, scopes)
            self.creds = flow.run_local_server(port=0)
            
            # Save the credentials for the next run
            with open(token_file, 'w') as token:
                token.write(self.creds.to_json())
            self.token_status_label.setText('Token Status: Authorized')
            self.log("App authorized successfully.")
        except Exception as e:
            self.log(f"Authorization failed: {str(e)}")
            self.token_status_label.setText('Token Status: Authorization Failed')

    def start_upload(self):
        if not self.folder_path:
            self.log("Please select a folder before uploading.")
            return
        if not self.creds or not self.creds.valid:
            self.log("Please authorize the app before uploading.")
            return

        # Retrieve user input
        sheet_id = self.sheet_id_input.text().strip()
        sheet_name = self.sheet_name_input.text().strip()
        id_column = self.id_column_input.text().strip()
        link_column = self.link_column_input.text().strip()
        start_row = int(self.start_row_input.text().strip())
        folder_id = self.folder_id_input.text().strip()

        self.log(f"Uploading files from: {self.folder_path}")
        self.upload_files(sheet_id, sheet_name, id_column, link_column, start_row, folder_id)

    def upload_files(self, sheet_id, sheet_name, id_column, link_column, start_row, folder_id):
        drive_service = build('drive', 'v3', credentials=self.creds)
        sheets_service = build('sheets', 'v4', credentials=self.creds)

        # Define ranges for Google Sheets
        id_range = f"{sheet_name}!{id_column}{start_row}:{id_column}"
        link_range = f"{sheet_name}!{link_column}{start_row}:{link_column}"

        # Fetch all IDs and Links from Google Sheets
        result_ids = sheets_service.spreadsheets().values().get(spreadsheetId=sheet_id, range=id_range).execute()
        ids = [item[0] for item in result_ids.get('values', []) if item]

        result_links = sheets_service.spreadsheets().values().get(spreadsheetId=sheet_id, range=link_range).execute()
        links = [item[0] for item in result_links.get('values', []) if item]

        # Upload stats and progress tracking
        total_files_in_directory = len(os.listdir(self.folder_path))
        files_uploaded_successfully = 0
        files_failed_to_upload = []
        skipped_files = []

        for filename in os.listdir(self.folder_path):
            file_path = os.path.join(self.folder_path, filename)
            if os.path.isfile(file_path):
                # Extract student ID from the new filename format
                file_id = extract_student_id(filename)
                
                # Safe logging for Unicode filenames
                safe_filename = filename.encode('ascii', 'replace').decode('ascii')
                self.log(f"Processing file: {safe_filename}")
                self.log(f"Extracted student ID: {file_id}")
                
                if not file_id:
                    skipped_files.append(filename)
                    self.log(f"No valid student ID found in filename: {safe_filename}")
                    continue
                
                # Check if this student ID exists in the spreadsheet
                if file_id not in ids:
                    skipped_files.append(filename)
                    self.log(f"Student ID {file_id} not found in spreadsheet")
                    continue

                # Get row index for this ID
                row_index = ids.index(file_id)

                # Check if a link already exists for this ID
                if row_index < len(links) and links[row_index]:
                    self.log(f"Link already exists for student ID {file_id}, skipping...")
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
                    self.log(f"✓ Uploaded: {safe_filename} -> Student ID: {file_id}")

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
