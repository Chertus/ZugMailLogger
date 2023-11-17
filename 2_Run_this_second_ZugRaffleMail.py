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

# Define the blacklist of player names
blacklisted_players = {"Alinorin", "Izlich", "Joulesburne", "Cowching", "Horde"}

# Compiled regular expression for efficiency
parse_pattern = re.compile(r'\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]\s+(.*?)\s+received item\(s\) from\s+(\w+).*?Receive Item\(s\):\s+(.*?)\n', re.DOTALL)

# Function to normalize data rows
def normalize_row(row):
    return [str(item).strip() for item in row]

# Function to parse the text
def parse_text(text):
    parsed_data = []
    processed_entries = set()
    matches = parse_pattern.finditer(text)

    for match in matches:
        full_date, _, sender, items = match.groups()

        # Skip entries from blacklisted players
        if sender in blacklisted_players:
            continue

        # Check for duplicates
        entry_id = f"{full_date}_{sender}"
        if entry_id in processed_entries:
            continue
        processed_entries.add(entry_id)

        sender = character_name_mapping.get(sender, sender)
        item_lines = items.strip().split('\n')
        for item_line in item_lines:
            item_parts = re.search(r'\[([\W_]*)([A-Za-z].*?)\]\s+\((\d+)\)', item_line)
            if item_parts:
                leading_symbols, item_name, quantity = item_parts.groups()
                quantity = int(quantity)
                stack_size = item_stacks.get(item_name, 1)
                full_stacks = math.floor(quantity / stack_size)
                parsed_data.append([full_date, sender, leading_symbols + item_name, full_stacks])

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
    # Fetch existing data in columns A-D
    sheet = sheets_service.spreadsheets()
    range_to_check = 'Donation Logs Active Week!A:D'
    result = sheet.values().get(spreadsheetId=sheet_id, range=range_to_check).execute()
    existing_values = result.get('values', [])

    # Normalize existing data for comparison
    existing_normalized = set(tuple(normalize_row(row)) for row in existing_values)

    # Filter out duplicate entries and count them
    new_entries = []
    skipped_duplicates = 0
    for row in data:
        normalized_row = normalize_row(row)
        if tuple(normalized_row) in existing_normalized:
            skipped_duplicates += 1
            continue
        new_entries.append(row)

    print(f"Skipped {skipped_duplicates} duplicate entries.")
    print(f"Adding {len(new_entries)} new entries.")

    # Check if there are new entries to update
    if not new_entries:
        print("No new entries to update.")
        return

    # Determine the next empty row for the new data
    next_row = len(existing_values) + 1
    range_for_new_data = f'Donation Logs Active Week!A{next_row}:D{next_row+len(new_entries)-1}'

    # Prepare the values to be inserted
    body = {'values': new_entries}

    # Use the Sheets API to append the new data
    result = sheet.values().update(
        spreadsheetId=sheet_id, range=range_for_new_data,
        valueInputOption='USER_ENTERED', body=body).execute()

    print(f"{result.get('updatedCells')} cells updated.")

# Function to remove duplicates from Google Sheets
def remove_duplicates_from_sheet(sheet_id, range_to_check):
    # Fetch the updated data in the specified range
    result = sheets_service.spreadsheets().values().get(spreadsheetId=sheet_id, range=range_to_check).execute()
    updated_values = result.get('values', [])

    # Normalize and identify duplicates
    seen = set()
    unique_values = []
    for row in updated_values:
        normalized_row = tuple(normalize_row(row))
        if normalized_row not in seen:
            seen.add(normalized_row)
            unique_values.append(row)

    # Check if duplicates were found
    if len(unique_values) != len(updated_values):
        print(f"Removing {len(updated_values) - len(unique_values)} duplicate entries.")

        # Clear the existing range
        sheets_service.spreadsheets().values().clear(spreadsheetId=sheet_id, range=range_to_check, body={}).execute()

        # Write back the unique values
        body = {'values': unique_values}
        sheets_service.spreadsheets().values().update(
            spreadsheetId=sheet_id, range=range_to_check,
            valueInputOption='USER_ENTERED', body=body).execute()
    else:
        print("No additional duplicates found.")

# Main execution
check_and_install_packages(required_packages)
doc_id = '1n_ZwKP_wA4FGNQNpVwyixcrX_NqrjMWyxcjpFlVBM0Y'

raw_text = read_google_doc(doc_id)
parsed_data = parse_text(raw_text)
update_google_sheets(sheet_id, parsed_data)
final_range_to_check = 'Donation Logs Active Week!A:D'
remove_duplicates_from_sheet(sheet_id, final_range_to_check)
