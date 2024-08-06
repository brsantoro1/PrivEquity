"""
The purpose of this script is to find updated information about people from AdamsStreet
"""

import openpyxl
from datetime import datetime
import win32com.client as win32
import traceback
import time
from Generate_Header_Dictionary import get_column_headers
from scrape import search_webpage
import time
from concurrent.futures import ThreadPoolExecutor, as_completed


try:
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = True
    print("Excel Application Initialized Successfully")
except Exception as e:
    print(f"Error initializing Excel Application: {e}")


#load the workbook and sheet
workbook_path = r"C:\\Users\\Bay Street - Larry B\\Documents\\Brielle\\Programming\\Projects\\Private Equity\\AdamsStreet.xlsm"
output_workbook_path = r"C:\\Users\\Bay Street - Larry B\\Documents\\Brielle\\Programming\\Projects\\Private Equity\\Outputs.xlsm"
sheet_name = "Master"
table_name = "Master"
wb = openpyxl.load_workbook(workbook_path, data_only=True)
sheet = wb[sheet_name]


def process_row(row, column_headers):
    name = str(row[column_headers['Name']])
    name_url = name.replace(' ', '-') + '/'
    #Check for specified conditions
    output = search_webpage(name_url)
    time_of_check = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    split = output.split(", ")
    title = split[0]
    if split.__len__ == 2:
        strategy = split[1]
    else:
        strategy = "Error"

    # Create a dictionary for the row
    return {
        # "Map Firm": row[column_headers['Firm']],
        "Name": name,
        "Title": title,
        "Strategy / Team": strategy,
        "Time Checked": time_of_check
    }


# Get column headers
column_headers = get_column_headers(workbook_path, sheet_name, table_name)

# Initialize an empty list to store the results
results = []


# Process rows using ThreadPoolExecutor
with ThreadPoolExecutor(max_workers=8) as executor:
    futures = [executor.submit(process_row, row, column_headers) for row in sheet.iter_rows(min_row=2, max_row=10, values_only=True)]
    for future in as_completed(futures):
        result = future.result()
        if result:
            results.append(result)
print(results)

#Create a new workbook for output
output_wb = openpyxl.Workbook()
output_ws = output_wb.active

# Write headers
headers = ["Name", "Title", "Strategy / Team", "Time Checked"]
output_ws.append(headers)

# Write data rows
for result in results:
    row = [result[header] for header in headers]
    output_ws.append(row)
    print(row)

# Save the output workbook
output_wb.save(output_workbook_path)
output_wb.close()
wb.close()


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'intern2@baystreetadvisorsllc.com'
mail.Cc = 'ojones@baystreetadvisorsllc.com'
mail.Subject = 'Adams Street'
mail.Body = str(results)
# mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

mail.Send()