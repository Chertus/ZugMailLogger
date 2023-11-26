import httplib2
import os
import csv
from googleapiclient.discovery import build
from google.oauth2 import service_account
from collections import defaultdict

# Setup the Sheets API
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
SERVICE_ACCOUNT_FILE = 'credentials.json'

creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# The ID and range of the spreadsheet.
SAMPLE_SPREADSHEET_ID = '1W_GKIr2yKxt5fF4KsyY4DaENZJSRYI062R6l7S07sOE'
SAMPLE_RANGE_NAME = 'Donation Logs Active Week!A:D'

service = build('sheets', 'v4', credentials=creds)

# Function to read data from the sheet and write to a temp file
def read_data_to_temp_file(sheet_id, range_name, temp_file_path):
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
    values = result.get('values', [])

    # Write headers and rows to the temp file
    with open(temp_file_path, 'w', newline='') as file:
        writer = csv.writer(file)
        # Assuming the first row in the sheet contains headers
        headers = values[0] if values else ['Date', 'Name', 'Item', 'Quantity']
        writer.writerow(headers)
        writer.writerows(values[1:])  # Skip header row

# Function to consolidate entries from a temp file
def consolidate_entries_from_temp_file(temp_file_path):
    data = defaultdict(int)
    with open(temp_file_path, 'r', newline='') as file:
        reader = csv.reader(file)
        next(reader)  # Skip header row
        for row in reader:
            key = tuple(row[:3])  # Date, Name, Item
            quantity = int(row[3]) if len(row) > 3 else 0
            data[key] += quantity

    return [list(key) + [str(data[key])] for key in data]

# Function to update the sheet from a temp file
def update_sheet_from_temp_file(sheet_id, range_name, temp_file_path):
    # Consolidate entries
    consolidated_data = consolidate_entries_from_temp_file(temp_file_path)

    # Clear the existing data
    service.spreadsheets().values().clear(
        spreadsheetId=sheet_id, range=range_name).execute()

    # Write consolidated data back to the sheet
    body = {'values': consolidated_data}
    service.spreadsheets().values().update(
        spreadsheetId=sheet_id, range=range_name,
        valueInputOption='USER_ENTERED', body=body).execute()

# Main execution
def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    temp_file_path = os.path.join(script_dir, 'temp_data.csv')
    
    # Read data from the sheet and write to temp file
    read_data_to_temp_file(SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME, temp_file_path)

    # Update the sheet from the temp file
    update_sheet_from_temp_file(SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME, temp_file_path)

if __name__ == '__main__':
    main()
