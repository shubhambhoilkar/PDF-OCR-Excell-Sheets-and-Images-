import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Config
pdf_path = "Payslip.pdf"
output_excel = "Payslip_all_tables.xlsx"

# Step 1: Extract all tables
all_tables = []

with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        page_tables = page.extract_tables()
        if page_tables:
            all_tables.extend(page_tables)

# Step 2: Create Excel Workbook
wb = Workbook()
ws = wb.active
ws.title = "PayslipRaw"

current_row = 1

for table in all_tables:
    # Transform tables to Dataframes
    df = pd.DataFrame(table)

    # Add two blank rows for seperations
    for row in dataframe_to_rows(df, index= False, header= False):
        ws.append(row)
    
    # Add two blank rows for seperations
    current_row +=len(df) + 2
    ws.append([])
    ws.append([])

# Step 3: Save Excell
wb.save(output_excel)


print(f"Excel created successfully: {output_excel}")
