import re
import math
import httplib2
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

# Required packages
required_packages = ['google-api-python-client', 'oauth2client']

# Regular expression pattern for parsing the text
parse_pattern = re.compile(r'\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]\n\s+(.*?) received item\(s\) from (.*?)\n\s+Receive Item\(s\):\s*\n((?:\s*\[\d+\] \[.*?\] \(\d+\)\n)+)')

# Character name mapping
character_name_mapping = {
    "Fitchbister": "Idolizeme",
    "Shiftables": "Idolizeme",
    # Add other mappings here
}

# Blacklisted players
blacklisted_players = {
    "Raztrad": True,
    "Grizzlegom": True,
    "Alinorin": True,
    "Izlich": True
}

# Item stack sizes
item_stacks = {
    "Adder's Tongue": 20,
    "Ametrine": 1,
    # Add other items here
}

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
    raw_text = document.get('body').get('content', '')
    return raw_text

# Function to parse the text
def parse_text(text):
    parsed_data = []
    processed_entries = set()
    matches = parse_pattern.finditer(text)

    for match in matches:
        full_date, _, sender, items = match.groups()

        # Skip entries from blacklisted players
        if sender in blacklisted_players and blacklisted_players[sender]:
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

# Function to update Google Sheets
def update_google_sheets(sheet_id, data):
    # Authentication and building the service
    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', ['https://www.googleapis.com/auth/spreadsheets'])
    http = credentials.authorize(httplib2.Http())
    service = build('sheets', 'v4', http=http)

    # Prepare the data for Sheets API
    body = {'values': data}
    range_name = 'Donation Logs Active Week!A:D'

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
    seen_entries = set()
    for row in values:
        if len(row) < 4:
            continue
        date, sender, item, _ = row
        entry_id = f"{date}_{sender}_{item}"
        if entry_id not in seen_entries and sender not in blacklisted_players:
            seen_entries.add(entry_id)
            unique_values.append(row)

    # Write back the unique values
    if unique_values:
        body = {'values': unique_values}
        sheets_service.spreadsheets().values().update(
            spreadsheetId=sheet_id, range=range_to_check,
            valueInputOption='USER_ENTERED', body=body).execute()
    else:
        print("No additional duplicates or blacklisted entries found.")

# Main execution
check_and_install_packages(required_packages)
doc_id = '1n_ZwKP_wA4FGNQNpVwyixcrX_NqrjMWyxcjpFlVBM0Y'

raw_text = read_google_doc(doc_id)
parsed_data = parse_text(raw_text)
update_google_sheets('1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms', parsed_data)
remove_duplicates_and_blacklisted_from_sheet('1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms', 'Donation Logs Active Week!A:D')
