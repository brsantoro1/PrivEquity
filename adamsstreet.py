"""
The purpose of this script is to find updated information about people from AdamsStreet
"""

import openpyxl
import win32com.client as win32
import traceback
from Generate_Header_Dictionary import get_column_headers
from scrape import search_webpage
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

try:
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = True
    print("Excel Application Initialized Successfully")
except Exception as e:
    print(f"Error initializing Excel Application: {e}")


#load the workbook and sheet
workbook_path = r"C:\\Users\\Bay Street - Larry B\\Documents\\Brielle\\Programming\\Projects\\Private Equity\\AdamsStreet.xlsm"
sheet_name = "Master"
table_name = "Master"
wb = openpyxl.load_workbook(workbook_path, data_only=True)
sheet = wb[sheet_name]


def process_row(row, column_headers):
    name = str(row[column_headers['Name']])
    name_url = name.lower().replace(' ', '-') + '/'
    #Check for specified conditions
    output = search_webpage(name_url)
    split = output.split(", ")
    title = split[0]
    if len(split) == 2:
        strategy = split[1]
    elif title == "PERSON NOT FOUND":
        strategy = "PERSON NOT FOUND"
    else:
        strategy = "Error"

    # Create a dictionary for the row
    return {
        # "Map Firm": row[column_headers['Firm']],
        "NAME": name,
        "TITLE": title,
        "STRATEGY / TEAM": strategy,
    }


# Get column headers
column_headers = get_column_headers(workbook_path, sheet_name, table_name)

# Initialize an empty list to store the results
results = []


# Process rows using ThreadPoolExecutor
with ThreadPoolExecutor(max_workers=8) as executor:
    futures = [executor.submit(process_row, row, column_headers) for row in sheet.iter_rows(min_row=2, values_only=True)]
    for future in as_completed(futures):
        result = future.result()
        if result:
            results.append(result)
print(results)

# Convert list of dictionaries to dataframe using pandas
df = pd.DataFrame(results)

# Save dataframe as pdf
rows_per_page = 15
pdf_path = r'C:\Users\Bay Street - Larry B\Documents\Brielle\Programming\Projects\Private Equity/dataframe.pdf'
with PdfPages(pdf_path) as pdf:
    for start_row in range(0, len(df), rows_per_page):
        end_row = min(start_row + rows_per_page, len(df))
        df_chunk = df.iloc[start_row:end_row]
    
        fig, ax = plt.subplots(figsize=(8.5, 2))  # Set size for the PDF page
        ax.axis('tight')
        ax.axis('off')
        table = ax.table(cellText=df_chunk.values, colLabels=df_chunk.columns, cellLoc='center', loc='center')
        
        table.auto_set_font_size(False)
        table.set_fontsize(6)
        
        # Bold the first row
        for (i, j), cell in table.get_celld().items():
            if i == 0:  # First row
                cell.set_text_props(weight='bold')

        pdf.savefig(fig, bbox_inches='tight')
        plt.close(fig)


# Send EMAILLLLLL
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'intern2@baystreetadvisorsllc.com'
# mail.Cc = 'ojones@baystreetadvisorsllc.com'
mail.Subject = 'Adams Street'
mail.Body = 'View attached PDF for dataframe'
attachment = pdf_path
mail.Attachments.Add(attachment)
# mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

mail.Send()