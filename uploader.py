import os
import sys
import json
import io
import re
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
FIRST_NAME_COLUMN = config['google_sheets'].get('first_name_column', 'B')
LAST_NAME_COLUMN = config['google_sheets'].get('last_name_column', 'C')
LINK_COLUMN = config['google_sheets']['link_column']
START_ROW = config['google_sheets']['start_row']
FOLDER_ID = config['google_drive']['folder_id']
ID_RANGE = f"{SHEET_NAME}!{ID_COLUMN}{START_ROW}:{ID_COLUMN}"
FIRST_NAME_RANGE = f"{SHEET_NAME}!{FIRST_NAME_COLUMN}{START_ROW}:{FIRST_NAME_COLUMN}"
LAST_NAME_RANGE = f"{SHEET_NAME}!{LAST_NAME_COLUMN}{START_ROW}:{LAST_NAME_COLUMN}"
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

    # Fetch all IDs, Names, and Links from the Google Sheet
    result_ids = sheets_service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=ID_RANGE).execute()
    ids = [item[0] for item in result_ids.get('values', []) if item]
    
    result_first_names = sheets_service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=FIRST_NAME_RANGE).execute()
    first_names = [item[0] if item else "" for item in result_first_names.get('values', [])]
    
    result_last_names = sheets_service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=LAST_NAME_RANGE).execute()
    last_names = [item[0] if item else "" for item in result_last_names.get('values', [])]
    
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
            
            # Extract names from filename
            first_name, last_name = extract_names(filename)
            
            # Safe printing for Unicode filenames
            safe_filename = filename.encode('ascii', 'replace').decode('ascii')
            print(f"Processing file: {safe_filename}")
            print(f"Extracted student ID: {file_id}")
            print(f"Extracted names: {first_name} {last_name}")
            
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
                    print(f"Student ID {file_id} not found in spreadsheet and no name match found")
                else:
                    print(f"No valid student ID found in filename and no name match: {safe_filename}")
                continue

            print(f"Matched by {match_method}, row {row_index + START_ROW}")

            # Check if a link already exists for this row
            if row_index < len(links) and links[row_index].strip():
                print(f"Link already exists for row {row_index + START_ROW}, skipping...")
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
