# Submission Uploader

A desktop application for uploading student submissions to Google Drive and automatically linking them in your Google Sheets gradebook. This tool streamlines the process of managing assignment submissions by organizing files in Google Drive and updating spreadsheet entries with direct links.

## Features

- **Batch Upload**: Upload multiple student submission files to Google Drive
- **Automatic Linking**: Automatically update Google Sheets with links to uploaded files
- **File Matching**: Match submission files to student IDs in your gradebook
- **GUI Interface**: User-friendly PyQt5 interface for easy operation
- **OAuth Authentication**: Secure Google API authentication
- **Progress Tracking**: Real-time upload progress and logging

## Prerequisites

- Python 3.6 or higher
- Google account with access to Google Drive and Google Sheets
- Google Cloud Project with Drive and Sheets API enabled

## Quick Start

### GUI Mode (Recommended for first-time setup)
1. Install dependencies: `pip install -r requirements.txt`
2. Set up Google API credentials (see [Google API Setup](#google-api-setup-and-authentication))
3. Run: `python app.py`
4. Fill in your configuration in the GUI - it saves automatically!
5. Click "Authorize App" to generate your token
6. Select folder and start uploading

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

The application now uses a unified `config.json` file that automatically synchronizes between the GUI and command-line interfaces.

### Automatic Configuration Sync

- **GUI (app.py)**: Automatically loads settings from `config.json` on startup and saves changes in real-time
- **CLI (uploader.py)**: Reads configuration from `config.json` at runtime
- **Unified Settings**: Changes made in the GUI are immediately available for CLI usage and vice versa

### Initial Setup

1. **Copy the sample configuration**:
   ```bash
   copy config.sample.json config.json
   ```
   
2. **Edit your configuration**:
   Open `config.json` and update the following fields:
   - `google_sheets.sheet_id`: Your Google Sheet ID
   - `google_sheets.sheet_name`: Sheet tab name (e.g., "A2", "Submissions")
   - `google_sheets.id_column`: Column containing student IDs (e.g., "C")
   - `google_sheets.link_column`: Column for file links (e.g., "L")
   - `google_sheets.start_row`: Data start row (usually 2 or 3)
   - `google_drive.folder_id`: Target Google Drive folder ID

### Configuration Structure

```json
{
  "google_sheets": {
    "sheet_id": "YOUR_SHEET_ID_HERE",
    "sheet_name": "Sheet1",
    "id_column": "C",
    "link_column": "L",
    "start_row": 3
  },
  "google_drive": {
    "folder_id": "YOUR_FOLDER_ID_HERE"
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

## Usage

### GUI Application

```bash
python app.py
```

**Features**:
- **Auto-load**: Configuration values are automatically loaded into form fields
- **Real-time save**: Changes are saved to `config.json` as you type
- **Visual feedback**: See configuration sync in real-time

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

### File Organization

- Place all submission files in a local folder
- Files should be named with student IDs that match your spreadsheet
- Supported formats: PDF, DOCX, and other common document formats
- Example naming: `816040296_A2.pdf`, `320053318_A2.pdf`

## Building Executable

To create a standalone executable:

```bash
pyinstaller --onefile --windowed --add-data "credentials.json;." app.py
```

The executable will be created in the `dist/` folder and can be distributed without requiring Python installation.

## Troubleshooting

### Configuration Issues

- **"Configuration file not found"**: Copy `config.sample.json` to `config.json`
- **UI not loading previous settings**: Check that `config.json` is valid JSON
- **Settings not saving**: Ensure write permissions in the project directory

### Authentication Issues

- **"No Token Found"**: 
  - GUI: Click "Authorize App" to generate a new token
  - CLI: Run `python cli_auth.py setup`
- **"Token Expired"**: The app will attempt to refresh automatically, or re-authorize if needed
- **"Authorization Failed"**: Check that `credentials.json` is present and valid
- **CLI auth not working**: Ensure `config.json` exists before running `cli_auth.py`

### Upload Issues

- **File matching problems**: Ensure submission filenames contain student IDs that match your spreadsheet
- **Permission errors**: Verify that your Google account has access to the target Drive folder and Sheet
- **API quota exceeded**: Google APIs have usage limits; wait and retry if needed

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
├── app.py              # Main GUI application (auto-syncs with config.json)
├── uploader.py         # Command-line upload script (uses config.json)
├── cli_auth.py         # CLI authentication helper tool
├── config.json         # Main configuration file (auto-created/synced)
├── config.sample.json  # Configuration template
├── credentials.json    # Google OAuth credentials (required)
├── token.json         # Generated auth token (auto-created)
├── requirements.txt    # Python dependencies
├── readme.md          # This file
├── upload_summary.txt  # Upload report (generated)
├── submissions/       # Example submissions folder
└── build/            # Build artifacts
```

## Security Notes

- Keep `credentials.json` secure and don't share it publicly
- The `token.json` file contains access tokens - treat it as sensitive
- Consider using environment variables for sensitive data in production
- Regularly review and rotate OAuth credentials if needed