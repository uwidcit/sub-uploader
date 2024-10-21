import os
import sys
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request
from tqdm import tqdm
import datetime

# Constants and parameters
FOLDER_PATH = sys.argv[1] if len(sys.argv) > 1 else None
SHEET_ID = "1ABgftOkfGjoxX_V-d8wKqsP0L960jv5IhRZpZaJiWKE"
SHEET_NAME = "A1"
ID_COLUMN = "D"
LINK_COLUMN = "M"
START_ROW = 3
FOLDER_ID = '1vZzzF81HrNYtMesnM7vnSSy8zF0Q6TpG'
ID_RANGE = f"{SHEET_NAME}!{ID_COLUMN}{START_ROW}:{ID_COLUMN}"
LINK_RANGE = f"{SHEET_NAME}!{LINK_COLUMN}{START_ROW}:{LINK_COLUMN}"
SCOPES = ['https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/spreadsheets']

def main():
    # Check for folder path
    if not FOLDER_PATH:
        print("Please provide the folder path as an argument.")
        sys.exit(1)

    # Authentication and service setup
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json')
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
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
            file_id = next((id for id in ids if id in filename), None)
            print("file_id: ", file_id)
            if not file_id:
                skipped_files.append(filename)
                continue
            
            # Get row index for this ID
            row_index = ids.index(file_id)

            # Check if a link already exists for this ID
            if row_index < len(links) and links[row_index]:
                continue  # Skip uploading if link already exists

            try:
                # Upload to Google Drive
                media = MediaFileUpload(file_path, resumable=True)
                file_metadata = {'name': filename, 'mimeType': 'application/octet-stream', 'parents': [FOLDER_ID]}
                file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()

                # Share and link
                permissions = {'role': 'reader', 'type': 'anyone'}
                drive_service.permissions().create(fileId=file['id'], body=permissions).execute()
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
    with open('upload_summary.txt', 'w') as report:
        # add date to report
        report.write("Upload Summary\n")
        report.write("--------------\n\n")
        for file in skipped_files:
            report.write(f"Skipped: {file}\n")
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
            report.write(f"Failed: {file} - Error: {error}\n")

if __name__ == '__main__':
    main()
