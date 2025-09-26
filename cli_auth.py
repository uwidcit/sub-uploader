#!/usr/bin/env python3
"""
CLI tool for setting up Google API authentication for the Submission Uploader.
This script helps generate the required authentication token for command-line usage.
"""

import os
import sys
import json
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

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
        print("Please create a config.json file or copy from config.sample.json")
        return None
    except json.JSONDecodeError:
        print(f"Invalid JSON in configuration file '{config_file}'.")
        return None

def check_credentials_file(credentials_file):
    """Check if credentials file exists."""
    if not os.path.exists(credentials_file):
        print(f"Credentials file '{credentials_file}' not found.")
        print("Please follow the setup instructions in the README to create credentials.json")
        return False
    return True

def generate_token(config):
    """Generate authentication token."""
    credentials_file = config['authentication'].get('credentials_file', 'credentials.json')
    token_file = config['authentication'].get('token_file', 'token.json')
    scopes = config['authentication'].get('scopes', [
        'https://www.googleapis.com/auth/drive.file',
        'https://www.googleapis.com/auth/spreadsheets'
    ])

    if not check_credentials_file(credentials_file):
        return False

    try:
        print("Starting OAuth authorization flow...")
        flow = InstalledAppFlow.from_client_secrets_file(credentials_file, scopes)
        creds = flow.run_local_server(port=0)
        
        # Save the credentials for future use
        with open(token_file, 'w') as token:
            token.write(creds.to_json())
        
        print(f"✓ Authentication successful!")
        print(f"✓ Token saved to '{token_file}'")
        return True
        
    except Exception as e:
        print(f"✗ Authorization failed: {str(e)}")
        return False

def check_token_status(config):
    """Check the status of existing token."""
    token_file = config['authentication'].get('token_file', 'token.json')
    scopes = config['authentication'].get('scopes', [
        'https://www.googleapis.com/auth/drive.file',
        'https://www.googleapis.com/auth/spreadsheets'
    ])

    if not os.path.exists(token_file):
        print(f"✗ Token file '{token_file}' not found.")
        return False

    try:
        creds = Credentials.from_authorized_user_file(token_file, scopes)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                    # Save refreshed token
                    with open(token_file, 'w') as token:
                        token.write(creds.to_json())
                    print("✓ Token refreshed successfully.")
                    return True
                except Exception as e:
                    print(f"✗ Token refresh failed: {e}")
                    return False
            else:
                print("✗ Token expired or invalid. Re-authorization needed.")
                return False
        else:
            print("✓ Token is valid and ready to use.")
            return True
    except Exception as e:
        print(f"✗ Error checking token: {e}")
        return False

def print_usage():
    """Print usage instructions."""
    print("Submission Uploader CLI Authentication Tool")
    print("==========================================")
    print()
    print("Usage:")
    print("  python cli_auth.py [command]")
    print()
    print("Commands:")
    print("  setup     - Generate new authentication token")
    print("  check     - Check status of existing token")
    print("  help      - Show this help message")
    print()
    print("Prerequisites:")
    print("  1. config.json file must exist (copy from config.sample.json)")
    print("  2. credentials.json file must exist (download from Google Cloud Console)")
    print()
    print("For detailed setup instructions, see README.md")

def main():
    if len(sys.argv) < 2:
        print_usage()
        return

    command = sys.argv[1].lower()

    if command == 'help':
        print_usage()
        return

    # Load configuration
    config = load_config()
    if not config:
        return

    if command == 'setup':
        print("Setting up Google API authentication...")
        print("This will open a browser window for authorization.")
        print()
        if generate_token(config):
            print()
            print("Setup complete! You can now use the command-line uploader:")
            print("python uploader.py /path/to/submissions/folder")
        else:
            print()
            print("Setup failed. Please check the error messages above.")
            sys.exit(1)

    elif command == 'check':
        print("Checking authentication status...")
        if check_token_status(config):
            print("Authentication is ready for command-line usage.")
        else:
            print("Authentication setup required. Run: python cli_auth.py setup")
            sys.exit(1)

    else:
        print(f"Unknown command: {command}")
        print_usage()
        sys.exit(1)

if __name__ == '__main__':
    main()