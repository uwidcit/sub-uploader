import os
import sys
import json
import io
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request
from tqdm import tqdm
import datetime

# Set console encoding to UTF-8 to handle Unicode characters
if sys.platform == 'win32':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

def load_config(config_file='config.json'):
    """Load configuration from JSON file."""
    try:
        with open(config_file, 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"Configuration file '{config_file}' not found.")
        sys.exit(1)
    except json.JSONDecodeError:
        print(f"Invalid JSON in configuration file '{config_file}'.")
        sys.exit(1)

# Load configuration
config = load_config()

# Extract configuration values
FOLDER_PATH = sys.argv[1] if len(sys.argv) > 1 else None
SHEET_ID = config['google_sheets']['sheet_id']
SHEET_NAME = config['google_sheets']['sheet_name']
ID_COLUMN = config['google_sheets']['id_column']
LINK_COLUMN = config['google_sheets']['link_column']
START_ROW = config['google_sheets']['start_row']
FOLDER_ID = config['google_drive']['folder_id']
ID_RANGE = f"{SHEET_NAME}!{ID_COLUMN}{START_ROW}:{ID_COLUMN}"
LINK_RANGE = f"{SHEET_NAME}!{LINK_COLUMN}{START_ROW}:{LINK_COLUMN}"
SCOPES = config['authentication']['scopes']
CREDENTIALS_FILE = config['authentication']['credentials_file']
TOKEN_FILE = config['authentication']['token_file']
SUMMARY_FILE = config['output']['summary_file']
MIME_TYPE = config['upload']['mime_type']
PERMISSIONS = config['upload']['permissions']

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

def main():
    # Check for folder path
    if not FOLDER_PATH:
        print("Please provide the folder path as an argument.")
        sys.exit(1)

    # Authentication and service setup
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    drive_service = build('drive', 'v3', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)

    # Fetch all IDs and Links from the Google Sheet
    result_ids = sheets_service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=ID_RANGE).execute()
    ids = [item[0] for item in result_ids.get('values', []) if item]
    
    result_links = sheets_service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=LINK_RANGE).execute()
    links = [item[0] for item in result_links.get('values', []) if item]

    # Upload stats and progress tracking
    total_files_in_directory = len(os.listdir(FOLDER_PATH))
    files_uploaded_successfully = 0
    files_failed_to_upload = []
    skipped_files = []

    for filename in tqdm(os.listdir(FOLDER_PATH), desc="Uploading Files", ncols=100):
        file_path = os.path.join(FOLDER_PATH, filename)
        if os.path.isfile(file_path):
            # Extract student ID from the new filename format
            file_id = extract_student_id(filename)
            
            # Safe printing for Unicode filenames
            safe_filename = filename.encode('ascii', 'replace').decode('ascii')
            print(f"Extracted student ID from '{safe_filename}': {file_id}")
            
            if not file_id:
                skipped_files.append(filename)
                print(f"No valid student ID found in filename: {safe_filename}")
                continue
            
            # Check if this student ID exists in the spreadsheet
            if file_id not in ids:
                skipped_files.append(filename)
                print(f"Student ID {file_id} not found in spreadsheet")
                continue
            
            # Get row index for this ID
            row_index = ids.index(file_id)

            # Check if a link already exists for this ID
            if row_index < len(links) and links[row_index]:
                print(f"Link already exists for student ID {file_id}, skipping...")
                continue  # Skip uploading if link already exists

            try:
                # Upload to Google Drive
                media = MediaFileUpload(file_path, resumable=True)
                file_metadata = {'name': filename, 'mimeType': MIME_TYPE, 'parents': [FOLDER_ID]}
                file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()

                # Share and link
                drive_service.permissions().create(fileId=file['id'], body=PERMISSIONS).execute()
                link = f"https://drive.google.com/file/d/{file['id']}/view"
                hyperlink_formula = f'=HYPERLINK("{link}", "Open File")'
                row_num = row_index + START_ROW
                update_range = f"{SHEET_NAME}!{LINK_COLUMN}{row_num}"
                values = [[hyperlink_formula]]
                body = {'values': values}
                sheets_service.spreadsheets().values().update(spreadsheetId=SHEET_ID, range=update_range, valueInputOption="USER_ENTERED", body=body).execute()
                files_uploaded_successfully += 1

            except Exception as e:
                files_failed_to_upload.append((filename, str(e)))

    # Write summary
    with open(SUMMARY_FILE, 'w', encoding='utf-8') as report:
        # add date to report
        report.write("Upload Summary\n")
        report.write("--------------\n\n")
        for file in skipped_files:
            # Clean filename for safe writing
            clean_filename = file.encode('ascii', 'replace').decode('ascii')
            report.write(f"Skipped: {clean_filename}\n")
        report.write(f"Folder path: {FOLDER_PATH}\n")
        report.write(f"Sheet ID: {SHEET_ID}\n")
        report.write(f"Sheet Name: {SHEET_NAME}\n")
        report.write(f"ID Column: {ID_COLUMN}\n")
        report.write(f"Link Column: {LINK_COLUMN}\n")
        report.write(f"Start Row: {START_ROW}\n")
        report.write(f"Folder ID: {FOLDER_ID}\n")
        report.write(f"ID Range: {ID_RANGE}\n")
        report.write(f"Link Range: {LINK_RANGE}\n\n")
        report.write(f"Date: {datetime.datetime.now()}\n\n")
        report.write(f"Total files in directory: {total_files_in_directory}\n")
        report.write(f"Files uploaded successfully: {files_uploaded_successfully}\n")
        report.write(f"Files failed to upload: {len(files_failed_to_upload)}\n\n")
        for file, error in files_failed_to_upload:
            # Clean filename and error message for safe writing
            clean_filename = file.encode('ascii', 'replace').decode('ascii')
            clean_error = str(error).encode('ascii', 'replace').decode('ascii')
            report.write(f"Failed: {clean_filename} - Error: {clean_error}\n")

if __name__ == '__main__':
    main()
