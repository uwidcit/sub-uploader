import os
import sys
import json
import io
import re
import csv
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request
from tqdm import tqdm
import datetime

def extract_content_after_file(filename):
    """Remove 'assignsubmission_file_' from Moodle submission filenames."""
    # Pattern: Name_ID_assignsubmission_file_ActualFilename.ext
    # Remove only the 'assignsubmission_file_' part, keep name and ID
    if 'assignsubmission_file_' in filename:
        return filename.replace('assignsubmission_file_', '')
    return filename  # Return original if pattern not found

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

def load_groups_from_csv(groups_file='groups.csv', mapping=None):
    """Load student to group mappings from CSV file.

    mapping: dict with keys 'member_first_name_column', 'member_last_name_column', 'group_name_column'
    """
    groups = {}
    if not os.path.exists(groups_file):
        return groups

    try:
        with open(groups_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)

            # Determine the header names to use (case-insensitive match)
            headers = reader.fieldnames or []

            def find_header(expected):
                if not expected:
                    return None
                for h in headers:
                    if h and h.lower() == expected.lower():
                        return h
                return None

            first_col = None
            last_col = None
            group_col = None
            if mapping:
                first_col = find_header(mapping.get('member_first_name_column', 'First Name'))
                last_col = find_header(mapping.get('member_last_name_column', 'Last Name'))
                group_col = find_header(mapping.get('group_name_column', 'Group Name'))

            # Fallbacks if headers not found
            if not first_col:
                first_col = find_header('First Name') or find_header('First name') or find_header('first_name')
            if not last_col:
                last_col = find_header('Last Name') or find_header('Last name') or find_header('last_name')
            if not group_col:
                group_col = find_header('Group Name') or find_header('Group') or find_header('group_name')

            for row in reader:
                first_name = (row.get(first_col, '') if first_col else '').strip()
                last_name = (row.get(last_col, '') if last_col else '').strip()
                group_name = (row.get(group_col, '') if group_col else '').strip()

                # Filter out empty rows
                if not group_name:
                    continue

                # Create full name for matching
                full_name = f"{first_name} {last_name}".strip()
                if full_name:
                    groups[full_name.lower()] = group_name

                if first_name:
                    groups[first_name.lower()] = group_name
                if last_name:
                    groups[last_name.lower()] = group_name

        print(f"Loaded {len(set(groups.values()))} unique groups for {len(groups)} name variations from {groups_file}")
        return groups
    except Exception as e:
        print(f"Error loading groups file: {e}")
        return {}

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
                    
        print(f"Loaded {len(matches)} filename matches from {matches_file}")
        return matches
    except Exception as e:
        print(f"Error loading matches file: {e}")
        return {}

def parse_args():
    """Parse command-line args for dry-run and optional folder path."""
    dry_run = False
    folder_arg = None
    args = sys.argv[1:]
    for a in args:
        if a in ('--dry-run', '-n'):
            dry_run = True
        elif not a.startswith('-') and not folder_arg:
            folder_arg = a
    return dry_run, folder_arg

def perform_dry_run(folder_path, groups_data, filename_matches):
    """Perform a mapping-only dry run without Google API calls.

    Writes a summary file similar to the upload summary but does not upload or touch Google Sheets/Drive.
    """
    total_files_in_directory = len([f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))])
    mapped = []
    skipped = []

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if not os.path.isfile(file_path):
            continue

        first_name, last_name = extract_names(filename)
        file_id = extract_student_id(filename)

        matched_group = None
        match_method = None

        # Try precomputed filename matches first
        if filename in filename_matches:
            matched_group = filename_matches[filename]
            match_method = 'matches.csv'

        # Then try groups mapping if enabled
        if not matched_group and groups_data:
            if first_name:
                full = f"{first_name} {last_name or ''}".strip()
                matched_group = find_group_by_student_name(full, groups_data)
                if matched_group:
                    match_method = 'group_mappings (full name)'

            if not matched_group:
                simple = extract_student_name_from_filename(filename)
                if simple:
                    matched_group = find_group_by_student_name(simple, groups_data)
                    if matched_group:
                        match_method = 'group_mappings (simple name)'

        if matched_group:
            mapped.append((filename, matched_group, match_method, file_id))
        else:
            skipped.append((filename, file_id))

    # Write dry-run summary
    with open(SUMMARY_FILE, 'w', encoding='utf-8') as report:
        report.write("Upload Dry Run Summary (no Google API calls)\n")
        report.write("-----------------------------------------\n\n")
        report.write(f"Folder path: {folder_path}\n")
        report.write(f"Date: {datetime.datetime.now()}\n\n")
        report.write(f"Total files in directory: {total_files_in_directory}\n")
        report.write(f"Files that would be mapped: {len(mapped)}\n")
        report.write(f"Files with no mapping: {len(skipped)}\n\n")

        if mapped:
            report.write("Files -> Group Mapping:\n")
            report.write("-----------------------\n")
            for fn, grp, method, fid in mapped:
                safe_fn = fn.encode('ascii', 'replace').decode('ascii')
                report.write(f"{safe_fn} => {grp} (method: {method}, id: {fid})\n")
            report.write("\n")

        if skipped:
            report.write("Files skipped (no mapping):\n")
            report.write("-------------------------\n")
            for fn, fid in skipped:
                safe_fn = fn.encode('ascii', 'replace').decode('ascii')
                report.write(f"{safe_fn} (id: {fid})\n")

    print(f"Dry run complete. Summary written to {SUMMARY_FILE}")

# Load configuration
config = load_config()

# Parse CLI args (supports '--dry-run' or '-n' and optional folder path)
DRY_RUN, arg_folder = parse_args()

# Extract folder path (CLI arg beats config)
FOLDER_ARG = arg_folder

# Load group mappings according to config
group_mapping_cfg = config.get('group_mappings', {})
groups_file = group_mapping_cfg.get('file', 'groups.csv')
groups_data = load_groups_from_csv(groups_file, mapping=group_mapping_cfg)
GROUP_MODE = len(groups_data) > 0

# Load pre-computed matches
filename_matches = load_matches_from_csv()

print(f"Group mode: {'ENABLED' if GROUP_MODE else 'DISABLED'}")

# Extract configuration values
# Folder path may be provided as CLI arg (positional) or via config; CLI arg parsed into FOLDER_ARG
FOLDER_PATH = FOLDER_ARG if FOLDER_ARG else config['submissions']['folder_path']
SHEET_ID = config['google_sheets']['sheet_id']
SHEET_NAME = config['google_sheets']['sheet_name']
ID_COLUMN = config['google_sheets']['id_column']
FIRST_NAME_COLUMN = config['google_sheets'].get('first_name_column', '')
LAST_NAME_COLUMN = config['google_sheets'].get('last_name_column', '')
LINK_COLUMN = config['google_sheets']['link_column']
START_ROW = config['google_sheets']['start_row']
FOLDER_ID = config['google_drive']['folder_id']
ID_RANGE = f"{SHEET_NAME}!{ID_COLUMN}{START_ROW}:{ID_COLUMN}"
FIRST_NAME_RANGE = f"{SHEET_NAME}!{FIRST_NAME_COLUMN}{START_ROW}:{FIRST_NAME_COLUMN}" if FIRST_NAME_COLUMN else None
LAST_NAME_RANGE = f"{SHEET_NAME}!{LAST_NAME_COLUMN}{START_ROW}:{LAST_NAME_COLUMN}" if LAST_NAME_COLUMN else None
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
    if not sheet_ids or not filename:
        return None
    
    # Use stripped filename for better matching
    stripped_filename = extract_content_after_file(filename)
    
    # Extract potential group names from stripped filename
    # Remove file extension and common patterns
    base_name = os.path.splitext(stripped_filename)[0]
    
    # Try to extract group-like patterns from filename
    import re
    
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

def main():
    # Check for folder path
    if not FOLDER_PATH:
        print("Please provide the folder path as an argument or set it in config.json.")
        sys.exit(1)
    
    # Verify folder exists
    if not os.path.exists(FOLDER_PATH):
        print(f"Error: Folder path '{FOLDER_PATH}' does not exist.")
        print("Please provide a valid folder path as an argument or update config.json.")
        sys.exit(1)

    # If dry-run, perform mapping only and exit (no Google API calls)
    if DRY_RUN:
        print("Running in dry-run mode: no Google API calls will be made.")
        perform_dry_run(FOLDER_PATH, groups_data, filename_matches)
        return

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
    ids = [item[0] if item else "" for item in result_ids.get('values', [])]
    
    # Fetch names only if columns are specified
    first_names = []
    last_names = []
    
    if FIRST_NAME_RANGE:
        try:
            result_first_names = sheets_service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=FIRST_NAME_RANGE).execute()
            first_names = [item[0] if item else "" for item in result_first_names.get('values', [])]
        except Exception as e:
            print(f"Warning: Could not fetch first names: {e}")
            first_names = []
    
    if LAST_NAME_RANGE:
        try:
            result_last_names = sheets_service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=LAST_NAME_RANGE).execute()
            last_names = [item[0] if item else "" for item in result_last_names.get('values', [])]
        except Exception as e:
            print(f"Warning: Could not fetch last names: {e}")
            last_names = []
    
    result_links = sheets_service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=LINK_RANGE).execute()
    links = [item[0] for item in result_links.get('values', []) if item]

    # Upload stats and progress tracking
    total_files_in_directory = len(os.listdir(FOLDER_PATH))
    files_uploaded_successfully = 0
    files_failed_to_upload = []
    skipped_files = []
    uploaded_ids = []  # Track which IDs/groups were uploaded

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
            group_name_to_write = None
            
            if GROUP_MODE:
                # In group mode, prefer precomputed filename matches, otherwise try member->group mapping
                matched_group = None

                if filename in filename_matches:
                    matched_group = filename_matches[filename]
                    match_source = 'CSV'
                else:
                    # Try to extract member name from filename and map to group using groups_data
                    fn_first, fn_last = extract_names(filename)
                    if fn_first:
                        full = f"{fn_first} {fn_last or ''}".strip()
                        matched_group = find_group_by_student_name(full, groups_data)

                    # Fallback: attempt a simpler student name extraction
                    if not matched_group:
                        simple_name = extract_student_name_from_filename(filename)
                        if simple_name:
                            matched_group = find_group_by_student_name(simple_name, groups_data)

                    match_source = 'group_mappings' if matched_group else None

                if matched_group:
                    try:
                        # Find the row where the ID column matches the group name
                        row_index = ids.index(matched_group)
                        match_method = f"group mode {match_source} match: {matched_group}"
                        group_name_to_write = None  # Don't overwrite existing group name
                    except ValueError:
                        # Group name not found in ID column - create new row
                        for i, id_entry in enumerate(ids):
                            if not id_entry or id_entry.strip() == "":
                                row_index = i
                                match_method = f"group mode new row: {matched_group}"
                                group_name_to_write = matched_group
                                ids[i] = matched_group
                                print(f"Creating new row for group '{matched_group}' at row {i + START_ROW}")
                                break

                        if row_index is None:
                            # No empty rows found, extend the list
                            row_index = len(ids)
                            ids.append(matched_group)
                            match_method = f"group mode extended row: {matched_group}"
                            group_name_to_write = matched_group
                            print(f"Extending spreadsheet for group '{matched_group}' at row {row_index + START_ROW}")
                else:
                    print(f"No match found in matches.csv or group mappings for: {safe_filename}")
            else:
                # Standard mode matching
                # First check pre-computed matches from CSV
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
                
                # Update link column
                update_range = f"{SHEET_NAME}!{LINK_COLUMN}{row_num}"
                values = [[hyperlink_formula]]
                body = {'values': values}
                sheets_service.spreadsheets().values().update(spreadsheetId=SHEET_ID, range=update_range, valueInputOption="USER_ENTERED", body=body).execute()
                
                # In group mode, write the group name to ID column for new rows
                if GROUP_MODE and group_name_to_write:
                    id_update_range = f"{SHEET_NAME}!{ID_COLUMN}{row_num}"
                    id_values = [[group_name_to_write]]
                    id_body = {'values': id_values}
                    sheets_service.spreadsheets().values().update(spreadsheetId=SHEET_ID, range=id_update_range, valueInputOption="USER_ENTERED", body=id_body).execute()
                    print(f"Created new row with group name: {group_name_to_write}")
                elif GROUP_MODE:
                    print(f"Using existing group row: {ids[row_index] if row_index < len(ids) else 'Unknown'}")
                
                files_uploaded_successfully += 1
                
                # Track the uploaded ID/group name
                uploaded_id = group_name_to_write if GROUP_MODE and group_name_to_write else (ids[row_index] if row_index < len(ids) else "Unknown")
                uploaded_ids.append(f"{uploaded_id} (row {row_index + START_ROW})")

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
        
        # List uploaded IDs/groups
        if uploaded_ids:
            report.write("IDs/Groups Uploaded:\n")
            report.write("-------------------\n")
            for uploaded_id in uploaded_ids:
                report.write(f"✓ {uploaded_id}\n")
            report.write("\n")
        for file, error in files_failed_to_upload:
            # Clean filename and error message for safe writing
            clean_filename = file.encode('ascii', 'replace').decode('ascii')
            clean_error = str(error).encode('ascii', 'replace').decode('ascii')
            report.write(f"Failed: {clean_filename} - Error: {clean_error}\n")

if __name__ == '__main__':
    main()
