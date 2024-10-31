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

# Initialize variables for processing
product_names = []
artist_names = []
quantities = []
prices = []
categories = []

# Parse date
date_text = date_text.strip()
date_parsed = datetime.strptime(date_text, '%d %B %Y %I:%M %p')

# Process email content
idx = text.find("ProductQuantityPrice")
if idx >= 0:
    idx += len("ProductQuantityPrice")
    text = text[idx:]

pattern = r'(.+?)\s+by\s+(.*?)\s*\(#\d+\)(\d+)\$(\d+\.\d{2})'
matches = re.findall(pattern, text)

for match in matches:
    product_name = match[0].strip()
    artist_name = match[1].strip()
    quantity = int(match[2])
    price = float(match[3])
    category_formula = '=IFERROR(VLOOKUP(INDIRECT("A"&ROW()), Categories!A:B, 2, FALSE), "Other")'

    product_names.append(product_name)
    artist_names.append(artist_name)
    quantities.append(quantity)
    prices.append(price)
    categories.append(category_formula)

# Prepare output data
output = {
    'product_names': product_names,
    'artist_names': artist_names,
    'quantities': quantities,
    'prices': prices,
    'categories': categories,
    'date_ddmmyyyy': date_parsed.strftime('%d/%m/%Y'),
    'month_year': date_parsed.strftime('%m/%y'),
    'time_hhmmss': date_parsed.strftime('%H:%M:%S')
}

# Google Sheets API Setup
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

credentials = Credentials.from_service_account_file(service_account_file, scopes=SCOPES)
gc = gspread.authorize(credentials)

spreadsheet = gc.open(google_sheet_name)
worksheet = spreadsheet.worksheet("Data")

# Prepare data to insert into Google Sheets
data_rows = [
    [output['product_names'][i], output['categories'][i], output['quantities'][i], output['prices'][i], 
     output['date_ddmmyyyy'], output['time_hhmmss'], output['month_year']]
    for i in range(len(output['product_names']))
]

# Find the next available row
next_row = len(worksheet.get_all_values()) + 1

# Batch update to avoid single quotes in Google Sheets
cell_range = f"A{next_row}:F{next_row + len(data_rows) - 1}"
worksheet.batch_update([{
    "range": cell_range,
    "values": data_rows
}], value_input_option="USER_ENTERED")

print("Data inserted into Google Sheets successfully!")
