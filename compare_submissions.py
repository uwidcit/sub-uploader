import os
import json
import csv
import re
from difflib import SequenceMatcher
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import datetime

def load_config(config_file='config.json'):
    """Load configuration from JSON file."""
    try:
        with open(config_file, 'r') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"Error loading config: {e}")
        return None

def normalize_text(text):
    """Normalize text by removing/standardizing punctuation and spaces."""
    if not text:
        return ""
    
    # Convert to lowercase
    text = text.lower()
    
    # Remove common file extensions
    text = re.sub(r'\.(pdf|docx?|txt|rtf)$', '', text, flags=re.IGNORECASE)
    
    # Remove common assignment indicators
    text = re.sub(r'[_\-\s]*(a\d+|assignment\s*\d*|project|submission)[_\-\s]*', '', text, flags=re.IGNORECASE)
    
    # Standardize separators (underscores, hyphens, spaces to single space)
    text = re.sub(r'[_\-\s]+', ' ', text)
    
    # Remove extra punctuation but keep letters and numbers
    text = re.sub(r'[^\w\s]', '', text)
    
    # Remove extra whitespace
    text = ' '.join(text.split())
    
    return text.strip()

def extract_group_names_from_filename(filename):
    """Extract potential group names from filename."""
    # Remove file extension
    base_name = os.path.splitext(filename)[0]
    
    # Split on common delimiters and extract potential group names
    potential_names = []
    
    # Try different splitting patterns
    patterns = [
        r'([a-zA-Z]+(?:[_\-\s]*[a-zA-Z]+)*)',  # Word groups
        r'([a-zA-Z]+\d*)',  # Words with optional numbers
        r'(team[_\-\s]*[a-zA-Z]+)',  # Team names
        r'(group[_\-\s]*[a-zA-Z]+)',  # Group names
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, base_name, re.IGNORECASE)
        for match in matches:
            normalized = normalize_text(match)
            if len(normalized) >= 3:  # Only consider names with 3+ characters
                potential_names.append(normalized)
    
    # Also add the full normalized filename
    potential_names.append(normalize_text(base_name))
    
    # Remove duplicates while preserving order
    seen = set()
    unique_names = []
    for name in potential_names:
        if name and name not in seen:
            seen.add(name)
            unique_names.append(name)
    
    return unique_names

def similarity_score(a, b):
    """Calculate similarity score between two strings."""
    if not a or not b:
        return 0.0
    
    # Exact match gets highest score
    if a == b:
        return 1.0
    
    # Use SequenceMatcher for fuzzy matching
    return SequenceMatcher(None, a, b).ratio()

def find_best_match(filename, id_entries):
    """Find the best matching ID entry for a filename."""
    # Strip Moodle prefix for matching, but keep original filename
    stripped_filename = extract_content_after_file(filename)
    potential_names = extract_group_names_from_filename(stripped_filename)
    
    best_match = {
        'matched_id': None,
        'matched_name': None,
        'similarity': 0.0,
        'suggested_name': None
    }
    
    # Normalize all ID entries
    normalized_ids = [(normalize_text(id_entry), id_entry) for id_entry in id_entries if id_entry]
    
    # Try each potential name from filename against each ID entry
    for potential_name in potential_names:
        for normalized_id, original_id in normalized_ids:
            score = similarity_score(potential_name, normalized_id)
            
            if score > best_match['similarity']:
                best_match['matched_id'] = original_id
                best_match['matched_name'] = potential_name
                best_match['similarity'] = score
                
                # Suggest the original ID as the correct spelling if similarity is good but not perfect
                if 0.7 <= score < 1.0:
                    best_match['suggested_name'] = original_id
    
    return best_match

def authenticate_google_services():
    """Authenticate with Google Services."""
    config = load_config()
    if not config:
        return None, None
    
    SCOPES = config['authentication']['scopes']
    TOKEN_FILE = config['authentication']['token_file']
    CREDENTIALS_FILE = config['authentication']['credentials_file']
    
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                print(f"Credentials file '{CREDENTIALS_FILE}' not found.")
                return None, None
            
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    
    return build('sheets', 'v4', credentials=creds), config

def get_id_entries():
    """Get ID entries from Google Sheets."""
    service, config = authenticate_google_services()
    if not service or not config:
        return []
    
    SHEET_ID = config['google_sheets']['sheet_id']
    SHEET_NAME = config['google_sheets']['sheet_name']
    ID_COLUMN = config['google_sheets']['id_column']
    START_ROW = config['google_sheets']['start_row']
    
    ID_RANGE = f"{SHEET_NAME}!{ID_COLUMN}{START_ROW}:{ID_COLUMN}"
    
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=ID_RANGE
        ).execute()
        
        values = result.get('values', [])
        # Flatten the list and filter out empty values
        id_entries = [row[0] for row in values if row and row[0].strip()]
        return id_entries
        
    except Exception as e:
        print(f"Error fetching ID entries: {e}")
        return []

def extract_content_after_file(filename):
    """Remove 'assignsubmission_file_' from Moodle submission filenames."""
    # Pattern: Name_ID_assignsubmission_file_ActualFilename.ext
    # Remove only the 'assignsubmission_file_' part, keep name and ID
    if 'assignsubmission_file_' in filename:
        return filename.replace('assignsubmission_file_', '')
    return filename  # Return original if pattern not found



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

def main():
    """Main function to compare submissions and generate report."""
    config = load_config()
    if not config:
        print("Failed to load configuration.")
        return
    
    # Get submissions folder path
    submissions_folder = config['submissions']['folder_path']
    if not os.path.exists(submissions_folder):
        print(f"Submissions folder not found: {submissions_folder}")
        return
    
    # Check for group mode
    groups_data = load_groups_from_csv()
    GROUP_MODE = len(groups_data) > 0
    
    print(f"Group mode: {'ENABLED' if GROUP_MODE else 'DISABLED'}")
    if GROUP_MODE:
        print(f"Loaded {len(groups_data)} student-group mappings")
    
    # Get ID entries from spreadsheet (or create dummy entries for group mode)
    if GROUP_MODE:
        print("Group mode: Using group names from groups.csv...")
        # Create a list of unique group names
        id_entries = list(set(groups_data.values()))
        print(f"Found {len(id_entries)} unique groups")
    else:
        print("Fetching ID entries from spreadsheet...")
        id_entries = get_id_entries()
        if not id_entries:
            print("No ID entries found or failed to fetch from spreadsheet.")
            return
        print(f"Found {len(id_entries)} ID entries in spreadsheet.")
    
    # Get all files in submissions folder
    files = [f for f in os.listdir(submissions_folder) 
             if os.path.isfile(os.path.join(submissions_folder, f)) 
             and not f.startswith('.')]
    
    print(f"Found {len(files)} files in submissions folder.")
    
    # Generate report
    report_data = []
    for filename in files:
        print(f"Processing: {filename}")
        stripped_filename = extract_content_after_file(filename)
        
        if GROUP_MODE:
            # In group mode, extract student name and find their group
            student_name = extract_student_name_from_filename(filename)
            if student_name:
                group_name = find_group_by_student_name(student_name, groups_data)
                if group_name:
                    match_result = {
                        'matched_id': group_name,
                        'matched_name': student_name,
                        'similarity': 1.0,
                        'suggested_name': None
                    }
                else:
                    match_result = {
                        'matched_id': 'NO MATCH',
                        'matched_name': student_name,
                        'similarity': 0.0,
                        'suggested_name': None
                    }
            else:
                match_result = {
                    'matched_id': 'NO MATCH',
                    'matched_name': 'N/A',
                    'similarity': 0.0,
                    'suggested_name': None
                }
        else:
            # Standard mode - use existing matching logic
            match_result = find_best_match(filename, id_entries)
        
        report_data.append({
            'filename': filename,
            'stripped_filename': stripped_filename if stripped_filename != filename else 'N/A',
            'matched_id': match_result['matched_id'] or 'NO MATCH',
            'extracted_name': match_result['matched_name'] or 'N/A',
            'similarity_score': f"{match_result['similarity']:.2f}",
            'suggested_spelling': match_result['suggested_name'] or 'N/A',
            'match_quality': get_match_quality(match_result['similarity']),
            'mode': 'GROUP' if GROUP_MODE else 'STANDARD'
        })
    
    # Write CSV report
    report_filename = "matches.csv"
    
    with open(report_filename, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['filename', 'stripped_filename', 'matched_id', 'extracted_name', 'similarity_score', 
                     'suggested_spelling', 'match_quality', 'mode']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        writer.writerows(report_data)
    
    print(f"\nReport generated: {report_filename}")
    
    # Print summary
    exact_matches = sum(1 for row in report_data if float(row['similarity_score']) == 1.0)
    good_matches = sum(1 for row in report_data if 0.7 <= float(row['similarity_score']) < 1.0)
    poor_matches = sum(1 for row in report_data if 0.3 <= float(row['similarity_score']) < 0.7)
    no_matches = sum(1 for row in report_data if float(row['similarity_score']) < 0.3)
    
    print(f"\nSummary:")
    print(f"Exact matches: {exact_matches}")
    print(f"Good matches (>= 70%): {good_matches}")
    print(f"Poor matches (30-69%): {poor_matches}")
    print(f"No matches (< 30%): {no_matches}")

def get_match_quality(similarity):
    """Determine match quality based on similarity score."""
    if similarity >= 1.0:
        return "EXACT"
    elif similarity >= 0.8:
        return "EXCELLENT"
    elif similarity >= 0.7:
        return "GOOD"
    elif similarity >= 0.5:
        return "FAIR"
    elif similarity >= 0.3:
        return "POOR"
    else:
        return "NO MATCH"

if __name__ == "__main__":
    main()