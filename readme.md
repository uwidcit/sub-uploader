# Submission Uploader

A desktop application for uploading student submissions to Google Drive and automatically linking them in your Google Sheets gradebook. This tool streamlines the process of managing assignment submissions by organizing files in Google Drive and updating spreadsheet entries with direct links.

## Features

- **Intelligent File Matching**: Match submission files by student ID or first/last name from Moodle exports
- **Batch Upload**: Upload multiple student submission files to Google Drive
- **Automatic Linking**: Automatically update Google Sheets with clickable hyperlinks to uploaded files
- **Link Conflict Prevention**: Skip files that would overwrite existing submission links
- **Dual Interface**: User-friendly PyQt5 GUI and powerful command-line interface
- **Smart Configuration**: Persistent configuration with auto-loading and manual save functionality
- **Credentials Detection**: Automatic detection of Google API credentials with helpful error messages
- **Enhanced Authentication**: Visual status indicators and improved OAuth flow
- **Unicode Support**: Handles international characters in filenames safely
- **Progress Tracking**: Real-time upload progress with detailed logging and emoji indicators

## Prerequisites

- Python 3.6 or higher
- Google account with access to Google Drive and Google Sheets
- Google Cloud Project with Drive and Sheets API enabled

## Quick Start

### GUI Mode (Recommended)
1. **Install dependencies**: `pip install -r requirements.txt`
2. **Set up Google API credentials** (see [Google API Setup](#google-api-setup-and-authentication))
3. **Run the application**: `python app.py`
4. **Follow the improved workflow**:
   - ✅ Configuration auto-loads from `config.json` (if exists)
   - 🔐 Click "Authorize App" (now at the top) - credentials detection included
   - ⚙️ Fill/modify configuration fields as needed
   - 📁 Use "Select Folder" button next to submissions path field
   - 💾 Click "Save Configuration" to persist changes
   - 🚀 Click "Start Upload" to begin processing

### CLI Mode
1. Copy configuration: `copy config.sample.json config.json`
2. Edit `config.json` with your settings
3. Set up authentication: `python cli_auth.py setup`
4. Upload files: `python uploader.py /path/to/submissions`

## Installing Dependencies

```bash
pip install -r requirements.txt
```

## Google API Setup and Authentication

### 1. Create Google Cloud Project and Enable APIs

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the following APIs:
   - Google Drive API
   - Google Sheets API

### 2. Create OAuth 2.0 Credentials

1. In the Google Cloud Console, navigate to **APIs & Services** > **Credentials**
2. Click **+ CREATE CREDENTIALS** > **OAuth client ID**
3. If prompted, configure the OAuth consent screen:
   - Choose **External** user type (unless you're in a Google Workspace organization)
   - Fill in the required information:
     - App name: "Submission Uploader"
     - User support email: Your email
     - Developer contact information: Your email
   - Add your email to test users during development
4. For the OAuth client ID:
   - Application type: **Desktop application**
   - Name: "Submission Uploader Desktop Client"
5. Download the credentials JSON file
6. Rename the downloaded file to `credentials.json` and place it in the project root directory

### 3. Generate Authentication Token

The application will generate a `token.json` file automatically when you first authorize it:

1. Run the application
2. Click the **"Authorize App"** button
3. Your web browser will open to Google's authorization page
4. Sign in to your Google account
5. Grant the requested permissions:
   - See, edit, create, and delete your Google Drive files
   - See, edit, create, and delete your spreadsheets in Google Sheets
6. The authorization will complete and `token.json` will be created automatically

**Note**: The `token.json` file contains your authorization token and should be kept secure. Don't share it or commit it to version control.

## Configuration

The application now uses a `config.json` file that loads into the GUI on startup with manual save functionality.

### Configuration Workflow

- **GUI (app.py)**: Automatically loads settings from `config.json` on startup
- **Manual Save**: Click "Save Configuration" button to save changes to `config.json`
- **CLI (uploader.py)**: Reads configuration from `config.json` at runtime
- **Unified Settings**: Saved changes from GUI are immediately available for CLI usage

### Initial Setup

1. **Copy the sample configuration**:
   ```bash
   copy config.sample.json config.json
   ```
   
2. **Edit your configuration**:
   Open `config.json` and update the following fields:
   - `google_sheets.sheet_id`: Your Google Sheet ID
   - `google_sheets.sheet_name`: Sheet tab name (e.g., "A2", "Submissions")
   - `google_sheets.id_column`: Column containing student IDs (e.g., "A")
   - `google_sheets.first_name_column`: Column containing first names (e.g., "B")
   - `google_sheets.last_name_column`: Column containing last names (e.g., "C")
   - `google_sheets.link_column`: Column for file links (e.g., "N")
   - `google_sheets.start_row`: Data start row (usually 2 or 3)
   - `google_drive.folder_id`: Target Google Drive folder ID
   - `submissions.folder_path`: Default path to submissions folder (optional)

### Configuration Structure

```json
{
  "google_sheets": {
    "sheet_id": "YOUR_SHEET_ID_HERE",
    "sheet_name": "Sheet1",
    "id_column": "A",
    "first_name_column": "B",
    "last_name_column": "C",
    "link_column": "N",
    "start_row": 3
  },
  "google_drive": {
    "folder_id": "YOUR_FOLDER_ID_HERE"
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
```

### Group mappings (Grouped mode)

To support course assignments where submissions are uploaded by one group member and should be recorded under a group name, the uploader can load a local CSV that maps members to group names. Add a `group_mappings` block to `config.json` pointing to your CSV and the column names to use:

```json
  "group_mappings": {
    "file": "C:\\path\\to\\groups.csv",
    "member_first_name_column": "First Name",
    "member_last_name_column": "Last Name",
    "group_name_column": "Group Name"
  }
```

Notes:
- The CSV must be readable from the path in `file` (absolute or relative to the project root).
- The uploader looks for headers case-insensitively but you should provide the exact header names when possible. The expected headers are `First Name`, `Last Name`, and `Group Name`.
- When grouped mode is enabled (the CSV loads successfully), the uploader will, for each filename, attempt the following in order:
  1. Check `matches.csv` (pre-computed filename -> group matches)
  2. Extract a student name from the filename and map the student to their group using the CSV
  3. Fallback name extraction strategies (simple name parts)

If a mapped group name is not present in the sheet ID column, the uploader will create a new row and write the group name into the ID column before writing the link.

## Usage

### GUI Application

```bash
python app.py
```

**Enhanced Features**:
- **Auto-load Configuration**: Settings automatically populate from `config.json` on startup
- **Integrated Folder Selection**: Select submissions folder with button positioned next to the path field
- **Smart Authentication Flow**: 
  - "Authorize App" button moved to top for better workflow
  - Automatic credentials file detection
  - Visual status indicators with colors and emojis
  - Error messages when credentials are missing
- **Manual Configuration Save**: Click "Save Configuration" to persist changes
- **Real-time Feedback**: Visual confirmation of save operations and authentication status

### Command Line Interface

#### Step 1: Authentication Setup

Use the CLI authentication tool to generate your token:

```bash
# Check if authentication is already set up
python cli_auth.py check

# Set up new authentication (opens browser)
python cli_auth.py setup

# Get help
python cli_auth.py help
```

#### Step 2: Upload Files

```bash
python uploader.py /path/to/submissions/folder
```

Dry-run (no Google API calls)
-----------------------------

If you want to verify how files will be mapped before performing any uploads, use the dry-run flag. This maps filenames to groups/IDs using `matches.csv` and the optional `group_mappings` CSV, and writes a summary to `upload_summary.txt` without contacting Google APIs.

```bash
# Run a mapping-only dry-run and write results to upload_summary.txt
python uploader.py --dry-run

# Or provide a folder path and dry-run in one command
python uploader.py --dry-run "C:\\path\\to\\submissions"
```

The dry-run summary includes:
- total files scanned
- files that would be mapped (filename -> group/ID and match method)
- files with no mapping (skipped)

Check `upload_summary.txt` in the project root after the dry-run completes.


**Prerequisites for CLI**:
1. `config.json` must exist with your settings
2. `credentials.json` must be downloaded from Google Cloud Console
3. Authentication token must be generated (use `cli_auth.py setup`)

### Setting Up Your Upload

1. **Sheet ID**: The ID from your Google Sheets URL (the long string between `/d/` and `/edit`)
2. **Sheet Name**: The name of the specific sheet tab (e.g., "Sheet1", "Gradebook")
3. **ID Column**: Column letter containing student IDs (e.g., "A", "B", "C")
4. **Link Column**: Column letter where file links will be placed (e.g., "D", "E", "F")
5. **Start Row**: Row number where student data begins (usually 2 if row 1 has headers)
6. **Google Drive Folder ID**: The ID of the destination folder (from the folder's URL)

### File Organization and Matching

**Supported File Formats**: PDF, DOCX, and other common document formats

**File Matching Strategies**:
1. **Primary**: Student ID matching from Moodle export format
2. **Secondary**: First and last name matching when student ID fails

**Moodle Export Naming Format**:
```
FirstName LastName_SubmissionID_assignsubmission_file_StudentID_COMP1600_A1.pdf
```

**Examples**:
- `Aadam Seenath_1835025_assignsubmission_file_816050357_COMP1600_A1.pdf`
- `Aaron Charran_1835097_assignsubmission_file_816049096_ COMP1600_A1.pdf`
- `Billy Dee Williams_1834952_assignsubmission_file_816044707_COMP1600_A1.pdf`

**Legacy Format Support**:
The system also supports older naming conventions like:
- `816040296_A2.pdf`
- `320053318_A2.pdf`
- `400014890.A2.pdf`

**Smart Matching Algorithm**:
1. **Primary**: Extract student ID from filename using regex patterns
2. **Secondary**: When ID matching fails, extract and match names:
   - Exact first name + last name match
   - First name + partial last name match
   - Normalized comparison (handles hyphens, spaces, case differences)
3. **Conflict Prevention**: Skip files where target cells already contain links

**Spreadsheet Requirements**:
- **ID Column**: Contains student IDs (e.g., 816050357, 320053318)
- **First Name Column**: Contains first names for name-based matching
- **Last Name Column**: Contains last names for name-based matching
- **Link Column**: Target column for submission file links (skipped if already populated)

## Building Executable

### Prerequisites for Building

1. **Install PyInstaller**:
   ```bash
   pip install pyinstaller
   ```

2. **Ensure all dependencies are installed**:
   ```bash
   pip install -r requirements.txt
   ```

### Building the GUI Application

**Windows**:
```bash
pyinstaller --onefile --windowed --add-data "credentials.json;." --name "SubmissionUploader" app.py
```

**Mac/Linux**:
```bash
pyinstaller --onefile --windowed --add-data "credentials.json:." --name "SubmissionUploader" app.py
```

### Building the CLI Application

**Windows**:
```bash
pyinstaller --onefile --add-data "credentials.json;." --name "SubmissionUploaderCLI" uploader.py
```

**Mac/Linux**:
```bash
pyinstaller --onefile --add-data "credentials.json:." --name "SubmissionUploaderCLI" uploader.py
```

### Build Options Explained

- `--onefile`: Creates a single executable file
- `--windowed`: Suppresses console window (GUI only)
- `--add-data`: Includes credentials.json in the build
- `--name`: Sets the executable name

### Advanced Build Configuration

Create a `build.spec` file for more control:

```python
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[('credentials.json', '.')],
    hiddenimports=[
        'google.oauth2.credentials',
        'google_auth_oauthlib.flow',
        'googleapiclient.discovery',
        'googleapiclient.http'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SubmissionUploader',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico'  # Optional: add your own icon
)
```

Then build with:
```bash
pyinstaller build.spec
```

### Distribution Package

The built executable will be in the `dist/` folder. For distribution:

1. **Include required files**:
   - The executable
   - `config.json` (template or pre-configured)
   - Installation instructions

2. **Create distribution folder**:
   ```
   SubmissionUploader_v1.0/
   ├── SubmissionUploader.exe (or SubmissionUploader on Mac/Linux)
   ├── config.json
   ├── README.txt
   └── credentials_setup_guide.txt
   ```

### Build Troubleshooting

**Common Issues**:
- **Missing modules**: Add to `hiddenimports` in spec file
- **File not found errors**: Ensure all data files are included with `--add-data`
- **Large executable size**: Consider using `--exclude-module` for unused packages
- **Slow startup**: Use `--onedir` instead of `--onefile` for faster startup

**Testing the Build**:
1. Test the executable on a clean system without Python
2. Verify all features work correctly
3. Check that credentials.json is properly bundled
4. Test both GUI and authentication flows

## Troubleshooting

### Configuration Issues

- **"Configuration file not found"**: Copy `config.sample.json` to `config.json`
- **UI not loading previous settings**: Check that `config.json` is valid JSON
- **Settings not saving**: Click "Save Configuration" button and ensure write permissions in the project directory

### Authentication Issues

- **❌ "No credentials.json file found"**: 
  - Download credentials from Google Cloud Console
  - Rename to `credentials.json` and place in project root
  - "Authorize App" button will be disabled until credentials are found

- **⚠️ "Token Status: No Token Found"**: 
  - GUI: Click "Authorize App" button (now at the top)
  - CLI: Run `python cli_auth.py setup`

- **⚠️ "Token Status: Expired, Reauthorization Needed"**: 
  - App will attempt automatic refresh
  - If refresh fails, click "Authorize App" again

- **❌ "Authorization Failed"**: 
  - Verify `credentials.json` is valid JSON
  - Check Google Cloud Console project settings
  - Ensure OAuth consent screen is configured

- **GUI shows red error messages**: 
  - Red text indicates missing or invalid credentials
  - Green text indicates successful authentication
  - Orange text indicates warnings or expired tokens

### Upload Issues

- **File matching problems**: 
  - Check that filenames follow Moodle export format or contain student IDs
  - Verify first/last name columns are correctly configured for name-based matching
  - Review upload log for specific matching failures

- **"Link already exists, skipping..."**: 
  - This is normal behavior - prevents overwriting existing submissions
  - Files are automatically skipped if the target spreadsheet cell already has content

- **Permission errors**: 
  - Verify Google account has edit access to the target Drive folder and Sheet
  - Check that OAuth scopes include both Drive and Sheets permissions

- **API quota exceeded**: 
  - Google APIs have usage limits; wait and retry if needed
  - Consider smaller batch sizes for large numbers of files

### UI and Configuration Issues

- **Configuration not loading**: 
  - Ensure `config.json` exists and is valid JSON
  - Check file permissions in the project directory

- **"Select Folder" button not working**: 
  - Button is now positioned between the label and text field
  - Selected folder automatically updates the submissions path configuration

- **Settings not persisting**: 
  - Click "Save Configuration" button after making changes
  - Manual save prevents accidental overwrites during form editing

### CLI-Specific Issues

- **"Invalid JSON in configuration file"**: Validate your `config.json` syntax
- **Token generation fails**: Ensure `credentials.json` is in the correct location
- **Upload script can't find config**: Run from the same directory as `config.json`

### Google Drive Folder ID

To find your Google Drive folder ID:
1. Open Google Drive in your browser
2. Navigate to the desired folder
3. Copy the folder ID from the URL: `https://drive.google.com/drive/folders/FOLDER_ID_HERE`

## File Structure

```
Sub_Uploader/
├── app.py                 # Main GUI application with enhanced UI
├── uploader.py           # CLI upload script with name matching
├── cli_auth.py           # CLI authentication helper tool
├── config.json           # Configuration file (auto-loads in GUI)
├── credentials.json      # Google OAuth credentials (required)
├── token.json           # Generated auth token (auto-created)
├── requirements.txt     # Python dependencies
├── readme.md           # Documentation (this file)
├── upload_summary.txt  # Upload report (generated after each run)
├── app.spec            # PyInstaller spec file (optional)
├── submissions/        # Example submissions folder with Moodle exports
│   ├── Aadam Seenath_1835025_assignsubmission_file_816050357_COMP1600_A1.pdf
│   ├── Aaron Charran_1835097_assignsubmission_file_816049096_ COMP1600_A1.pdf
│   └── ...
├── build/              # PyInstaller build artifacts
└── dist/               # Built executables
    ├── SubmissionUploader.exe (Windows)
    └── SubmissionUploaderCLI.exe
```

### Key Files Explained

- **app.py**: Enhanced GUI with improved authentication flow, smart folder selection, and visual feedback
- **uploader.py**: CLI version with intelligent file matching (student ID + name fallback)
- **config.json**: Unified configuration file that auto-loads in GUI and supports manual save
- **credentials.json**: OAuth credentials from Google Cloud Console (must be added by user)
- **token.json**: Auto-generated authentication token (keep secure)
- **submissions/**: Contains student submission files from Moodle exports

## Security Notes

- Keep `credentials.json` secure and don't share it publicly
- The `token.json` file contains access tokens - treat it as sensitive
- Consider using environment variables for sensitive data in production
- Regularly review and rotate OAuth credentials if needed