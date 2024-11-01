import os
import re
import json
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

# Setup Google Sheets access
service_account_file = os.getenv("SERVICE_ACCOUNT_FILE")
google_sheet_name = os.getenv("GOOGLE_SHEET_NAME")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
credentials = Credentials.from_service_account_file(service_account_file, scopes=SCOPES)
gc = gspread.authorize(credentials)
spreadsheet = gc.open(google_sheet_name)
worksheet = spreadsheet.worksheet("Data")

# Get input data from environment variables
order_text = os.getenv("ORDER_DATA", "")
date_text = os.getenv("DATE_DATA", "")
input_text = json.loads(os.getenv("INPUT_DATA", "")) 

# Debug
print(order_text)
print(date_text)
print(input_text)

# Initialize output dictionary
output = {
    "order_num": [],
    "product_names": [],
    "artist_names": [],
    "quantities": [],
    "prices": [],
    "categories": [],
    "date_ddmmyyyy": [],
    "time_hhmmss": [],
    "month_year": []
}

# Parse the date in the given format
date_parsed = datetime.strptime(date_text.strip(), '%Y-%m-%dT%H:%M:%S.%fZ')

for sale_details in input_text:
    product_artist = re.search(r"^(.*?)\s+by\s+([^\(]+)", sale_details['product'])

    # Append extracted data to output lists
    output["order_num"].append(order_text)
    output["product_names"].append(product_artist.group(1).strip())
    output["artist_names"].append(product_artist.group(2).strip())
    output["quantities"].append(sale_details['quantity'])
    output["prices"].append(re.sub(r'[^\d.]', '', sale_details["price"]))
    output["categories"].append('=IFERROR(VLOOKUP(INDIRECT("B"&ROW()), Categories!A:B, 2, FALSE), "Other")')
    output["date_ddmmyyyy"].append(date_parsed.strftime('%d/%m/%Y'))
    output["time_hhmmss"].append(date_parsed.strftime('%H:%M:%S'))
    output["month_year"].append(date_parsed.strftime('%m/%Y'))

# Prepare data for batch update
# Remove artist name; to be added back in later if needed
data_rows = [
    [output['order_num'][i], output['product_names'][i], output['quantities'][i], output['prices'][i], 
     output['categories'][i], output['date_ddmmyyyy'][i], output['time_hhmmss'][i], output['month_year'][i]]
    for i in range(len(output['product_names']))
]

# Find the next available row
next_row = len(worksheet.get_all_values()) + 1

# Batch update to Google Sheets
cell_range = f"A{next_row}:H{next_row + len(data_rows) - 1}"
worksheet.batch_update([{
    "range": cell_range,
    "values": data_rows
}], value_input_option="USER_ENTERED")

print("Data inserted into Google Sheets successfully!")
