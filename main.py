import os
import re
import json
import time
import random
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.errors import HttpError

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
attachments_version = os.getenv("ATTACHMENTS", "false").lower() == "true"

# Debug
print("Order Text:", order_text)
print("Date Text:", date_text)
print("Input Text:", input_text)

# Initialize output dictionary
output = {
    "order_num": [],
    "product_names": [],
    "artist_names": [],
    "quantities": [],
    "prices": [],
    "categories": [],
    "date_time": [],
    "month_year": []
}

# Parse the date in the given format
date_parsed = datetime.strptime(date_text.strip(), '%a, %d %b %Y %H:%M:%S %z') \
    if attachments_version else datetime.strptime(date_text.strip(), '%Y-%m-%dT%H:%M:%S.%fZ')

for sale_details in input_text:
    product_artist = re.search(r"^(.*?)\s+by\s+([^\(]+)", sale_details['product'])

    # Append extracted data to output lists
    output["order_num"].append(order_text)
    output["product_names"].append(product_artist.group(1).strip())
    output["artist_names"].append(product_artist.group(2).strip())
    output["quantities"].append(sale_details['quantity'])
    output["prices"].append(re.sub(r'[^\d.]', '', sale_details["price"]))
    output["categories"].append('=IFERROR(VLOOKUP(INDIRECT("B"&ROW()), Categories!A:B, 2, FALSE), "Other")')
    output["date_time"].append(date_parsed.strftime('%d/%m/%Y %H:%M:%S'))
    output["month_year"].append(date_parsed.strftime('%m/%Y'))

# Prepare data for batch update
# Remove artist name; to be added back in later if needed
data_rows = [
    [output['order_num'][i], output['product_names'][i], output['categories'][i], output['quantities'][i], output['prices'][i], 
     output['date_time'][i], output['month_year'][i]]
    for i in range(len(output['product_names']))
]

def read_sheet_with_retry(worksheet, range_name, retries=5):
    for attempt in range(retries):
        try:
            data = worksheet.get(range_name)
            return data
        except HttpError as e:
            if e.resp.status == 429:
                print("Quota exceeded. Retrying in a few seconds...")
                time.sleep(random.randint(30, 90))  # Wait before retrying
            else:
                raise  # Re-raise other exceptions
    raise Exception("Failed to read from Google Sheets after multiple retries. Order num: {data_rows['order_num'][0]}")

def batch_update_with_retry(worksheet, cell_range, data_rows, retries=5):
    for attempt in range(retries):
        try:
            worksheet.batch_update([{
                "range": cell_range,
                "values": data_rows
            }], value_input_option="USER_ENTERED")
            print("Data inserted into Google Sheets successfully!")
            return
        except HttpError as e:
            if e.resp.status == 429:
                print("Quota exceeded. Retrying in a few seconds...")
                time.sleep(random.randint(30, 90))  # Wait before retrying
            else:
                raise  # Re-raise other exceptions
    raise Exception(f"Failed to write to Google Sheets after multiple retries. Order num: {data_rows['order_num'][0]}")

# Find the next available row
next_row = len(read_sheet_with_retry(worksheet, "A:A")) + 1  # Adjusted for retry mechanism

# Define the cell range for batch update
cell_range = f"A{next_row}:G{next_row + len(data_rows) - 1}"

# Perform the batch update with retry logic
batch_update_with_retry(worksheet, cell_range, data_rows)
