import re
import math
import hashlib
import httplib2
import json
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

# Required packages
required_packages = ['google-api-python-client', 'oauth2client']

# Load JSON data
with open('character_mapping.json') as file:
    character_name_mapping = json.load(file)

with open('player_blacklist.json') as file:
    blacklisted_players = json.load(file)

with open('item_stacks.json') as file:
    item_stacks = json.load(file)

# Regular expression pattern for parsing the text
parse_pattern = re.compile(r'\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]\n\s+(.*?) received item\(s\) from (.*?)\n\s+Receive Item\(s\):\s*\n((?:\s*\[\d+\] \[.*?\] \(\d+\)\n)+)')

# Function to check and install required packages
def check_and_install_packages(packages):
    import subprocess
    import sys

    for package in packages:
        try:
            __import__(package)
        except ImportError:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Function to read data from Google Doc
def read_google_doc(doc_id):
    # Authentication and building the service
    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', ['https://www.googleapis.com/auth/documents.readonly'])
    http = credentials.authorize(httplib2.Http())
    service = build('docs', 'v1', http=http)

    # Request to read the document
    document = service.documents().get(documentId=doc_id).execute()
    raw_text = []
    for element in document.get('body').get('content', []):
        if 'paragraph' in element:
            for para_element in element['paragraph']['elements']:
                if 'textRun' in para_element:
                    raw_text.append(para_element['textRun']['content'])
    return ''.join(raw_text)

# Function to parse the text
def parse_text(text):
    parsed_data = []
    matches = parse_pattern.finditer(text)

    for match in matches:
        full_date, _, sender, items = match.groups()

        # Skip entries from blacklisted players
        if sender in blacklisted_players and blacklisted_players[sender]:
            continue

        sender = character_name_mapping.get(sender, sender)
        item_lines = items.strip().split('\n')
        for item_line in item_lines:
            item_parts = re.search(r'\[([\W_]*)([A-Za-z].*?)\]\s+\((\d+)\)', item_line)
            if item_parts:
                leading_symbols, item_name, quantity = item_parts.groups()
                quantity = int(quantity)
                stack_size = item_stacks.get(item_name, 1)
                full_stacks = math.floor(quantity / stack_size)

                # Generate hash for duplicate detection
                entry_hash = hashlib.md5(f"{full_date}_{sender}_{item_name}_{full_stacks}".encode()).hexdigest()

                parsed_data.append([full_date, sender, leading_symbols + item_name, full_stacks, '', '', entry_hash])

    return parsed_data

# Function to update Google Sheets
def update_google_sheets(sheet_id, data):
    # Authentication and building the service
    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', ['https://www.googleapis.com/auth/spreadsheets'])
    http = credentials.authorize(httplib2.Http())
    service = build('sheets', 'v4', http=http)

    # Prepare the data for Sheets API
    body = {'values': data}
    range_name = 'Donation Logs Active Week!A:D'  # Update only columns A to D

    # Update the sheet
    service.spreadsheets().values().append(
        spreadsheetId=sheet_id, range=range_name,
        valueInputOption='USER_ENTERED', body=body).execute()

# Function to remove duplicates and blacklisted entries from the sheet
def remove_duplicates_and_blacklisted_from_sheet(sheet_id, range_to_check):
    # Authentication and building the service
    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', ['https://www.googleapis.com/auth/spreadsheets'])
    http = credentials.authorize(httplib2.Http())
    sheets_service = build('sheets', 'v4', http=http)

    # Read the current data from the sheet
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=sheet_id, range=range_to_check).execute()
    values = result.get('values', [])

    # Process for duplicates and blacklisted entries
    unique_values = []
    seen_hashes = set()
    for row in values:
        if len(row) < 7:
            continue
        date, sender, item, stacks, _, _, entry_hash = row
        if entry_hash not in seen_hashes and sender not in blacklisted_players:
            seen_hashes.add(entry_hash)
            unique_values.append([date, sender, item, stacks])  # Exclude columns E, F, and G

    # Write back the unique values
    if unique_values:
        body = {'values': unique_values}
        sheets_service.spreadsheets().values().update(
            spreadsheetId=sheet_id, range='Donation Logs Active Week!A:D',  # Update only columns A to D
            valueInputOption='USER_ENTERED', body=body).execute()
    else:
        print("No additional duplicates or blacklisted entries found.")

# Main execution
check_and_install_packages(required_packages)
doc_id = '1n_ZwKP_wA4FGNQNpVwyixcrX_NqrjMWyxcjpFlVBM0Y'

raw_text = read_google_doc(doc_id)
parsed_data = parse_text(raw_text)
update_google_sheets('1W_GKIr2yKxt5fF4KsyY4DaENZJSRYI062R6l7S07sOE', parsed_data)
remove_duplicates_and_blacklisted_from_sheet('1W_GKIr2yKxt5fF4KsyY4DaENZJSRYI062R6l7S07sOE', 'Donation Logs Active Week!A:D')
