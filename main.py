import os
import re
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

# Get input data from environment variables
text = os.getenv("INPUT_DATA", "")
date_text = os.getenv("DATE_DATA", "")
google_sheet_name = os.getenv("GOOGLE_SHEET_NAME")  # Name of your Google Sheet
service_account_file = os.getenv("SERVICE_ACCOUNT_FILE")  # Path to Google service account JSON file

print(text)
print(date_text)

# Initialize variables for processing
output = {
    'product_names': [],
    'artist_names': [],
    'quantities': [],
    'prices': [],
    'categories': [],
    'date_ddmmyyyy': [],
    'time_hhmmss': [],
    'month_year': []
}


# Parse date
date_text = date_text.strip()
date_parsed = datetime.strptime(date_text, '%Y-%m-%dT%H:%M:%S.%fZ')

# Process email content
idx = text.find("Product Quantity Price\\n")
idx += len("Product Quantity Price\\n")
text = text[idx:]

pattern = r'(.+?)\s+by\s+(.*?)\s*\(#\d+\)\s+(\d+)\s+\$(\d+\.\d{2})'
for match in re.findall(pattern, text):
    output["product_names"].append(match[0].strip())
    output["artist_names"].append(match[1].strip())
    output["quantities"].append(match[2])
    output["prices"].append(match[3])
    output["categories"].append('=IFERROR(VLOOKUP(INDIRECT("A"&ROW()), Categories!A:B, 2, FALSE), "Other")')
    output["date_ddmmyyyy"].append(date_parsed.strftime('%d/%m/%Y'))
    output["time_hhmmss"].append(date_parsed.strftime('%H:%M:%S'))
    output["month_year"].append(date_parsed.strftime('%m/%y'))

# Google Sheets API Setup
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

credentials = Credentials.from_service_account_file(service_account_file, scopes=SCOPES)
gc = gspread.authorize(credentials)

spreadsheet = gc.open(google_sheet_name)
worksheet = spreadsheet.worksheet("Data")

# Prepare data to insert into Google Sheets
data_rows = [
    [output['product_names'][i], output['categories'][i], output['quantities'][i], output['prices'][i], 
     output['date_ddmmyyyy'][i], output['time_hhmmss'][i], output['month_year'][i]]
    for i in range(len(output['product_names']))
]

# Find the next available row
next_row = len(worksheet.get_all_values()) + 1

# Batch update to avoid single quotes in Google Sheets
cell_range = f"A{next_row}:G{next_row + len(data_rows) - 1}"
worksheet.batch_update([{
    "range": cell_range,
    "values": data_rows
}], value_input_option="USER_ENTERED")

print("Data inserted into Google Sheets successfully!")
