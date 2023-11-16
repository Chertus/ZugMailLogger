import os
import subprocess
import sys
import json
from googleapiclient import discovery
from google.oauth2 import service_account
import re
import math
import pkg_resources

# List of required packages
required_packages = [
    'google-api-python-client', 
    'google-auth',
    'google-auth-httplib2',
    'google-auth-oauthlib'
]

# Function to check and install packages
def check_and_install_packages(packages):
    for package in packages:
        try:
            pkg_resources.get_distribution(package)
            print(f"{package} is already installed.")
        except pkg_resources.DistributionNotFound:
            print(f"Installing {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Google API setup
SCOPES = ['https://www.googleapis.com/auth/documents.readonly',
          'https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'D:\\Projects\\ZugMailRaffleUpdater\\zugmaildonations-4e0d971c18af.json'

credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Initialize the Google Docs and Sheets API clients
docs_service = discovery.build('docs', 'v1', credentials=credentials)
sheets_service = discovery.build('sheets', 'v4', credentials=credentials)

# Define the sheet_id variable
sheet_id = '1W_GKIr2yKxt5fF4KsyY4DaENZJSRYI062R6l7S07sOE'

# Load character name mapping from file
with open('character_mapping.json', 'r') as file:
    character_name_mapping = json.load(file)

# Load item stacks from file
with open('item_stacks.json', 'r') as file:
    item_stacks = json.load(file)

# Compiled regular expression for efficiency
parse_pattern = re.compile(r'\[(\d{4}-\d{2}-\d{2})\s\d{2}:\d{2}:\d{2}\]\s+(.*?)\s+received item\(s\) from\s+(\w+).*?Receive Item\(s\):\s+(.*?)\n', re.DOTALL)

# Function to parse the text
def parse_text(text):
    parsed_data = []
    matches = parse_pattern.finditer(text)

    for match in matches:
        date, _, sender, items = match.groups()
        month_day = '-'.join(date.split('-')[1:])
        sender = character_name_mapping.get(sender, sender)
        item_lines = items.strip().split('\n')
        for item_line in item_lines:
            item_parts = re.search(r'\[([\W_]*)([A-Za-z].*?)\]\s+\((\d+)\)', item_line)
            if item_parts:
                leading_symbols, item_name, quantity = item_parts.groups()
                quantity = int(quantity)
                stack_size = item_stacks.get(item_name, 1)
                full_stacks = math.floor(quantity / stack_size)
                parsed_data.append([month_day, sender, leading_symbols + item_name, full_stacks])

    return parsed_data

# Function to read the document
def read_google_doc(doc_id):
    document = docs_service.documents().get(documentId=doc_id).execute()
    text_content = ''
    for element in document.get('body').get('content'):
        if 'paragraph' in element:
            for para_element in element['paragraph']['elements']:
                if 'textRun' in para_element:
                    text_content += para_element['textRun']['content']
    return text_content

# Function to update Google Sheets
def update_google_sheets(sheet_id, data):
    # Determine the next empty row
    sheet = sheets_service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id, range='Donation Logs Active Week!A:D').execute()
    values = result.get('values', [])
    next_row = len(values) + 1

    # Specify the range for the new data
    range_ = f'Donation Logs Active Week!A{next_row}:D{next_row+len(data)-1}'

    # Prepare the values to be inserted
    body = {'values': data}

    # Use the Sheets API to update the sheet
    result = sheet.values().update(
        spreadsheetId=sheet_id, range=range_,
        valueInputOption='USER_ENTERED', body=body).execute()

    print(f"{result.get('updatedCells')} cells updated.")

# Main execution
check_and_install_packages(required_packages)
doc_id = '1n_ZwKP_wA4FGNQNpVwyixcrX_NqrjMWyxcjpFlVBM0Y'

raw_text = read_google_doc(doc_id)
parsed_data = parse_text(raw_text)
update_google_sheets(sheet_id, parsed_data)
