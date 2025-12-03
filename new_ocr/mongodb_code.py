import pandas as pd
from pymongo import MongoClient

try:
    # 1. Load Excel
    file_path = "C:\\Users\\Developer\\Shubham_files\\ocr\\output.xlsx"
    df= pd.read_excel(file_path)
    print(df.to_string())
    #Additional FIX
    df.columns = df.columns.map(str)

    # 2. Identify column groups
    payslips_cols = [c for c in df.columns if c.startswith("PAYSLIP_")]
    timesheet_cols = [c for c in df.columns if c.startswith("TIMESHEET_")]

    identify_cols = ["EMPLOYEE_NAME", "PASSPORT_NUMBER"]
    identify_cols = [c for c in identify_cols if c in df.columns]

    # 3. Slip into the names for MongoDB
    df_payslip = df[identify_cols + payslips_cols].copy()
    df_timesheet = df[identify_cols + timesheet_cols].copy()

    # 4. Clean column names for MongoDB
    df_payslip.columns = [c.lower() for c in df_payslip.columns]
    df_timesheet.columns = [c.lower() for c in df_timesheet.columns]

    df_payslip = df_payslip.where(pd.notnull(df_payslip), None)
    df_timesheet = df_timesheet.where(pd.notnull(df_timesheet), None)

    # 5. MongoDB Nan -> None for MongoDB
    client = MongoClient("mongodb://localhost:27017")
    db = client["OCR_Database"]

    payslip_collection = db["Payslip_Data"]
    timesheet_collection = db["Timesheet_Data"]

    # 6 . Insert into MongoDB:
    payslip_data = df_payslip.to_dict(orient="records")
    timesheet_data = df_timesheet.to_dict(orient="records")

except Exception as e:
    print("Error at getting the sheet data. ", e)

try:
    if payslip_data:
        payslip_collection.insert_many(payslip_data)
        print("Data inserted into Payslip_Data.")

    if timesheet_data:
        timesheet_collection.insert_many(timesheet_data)
        print("Data inserted into Timesheet_Data.")

except Exception as e:
    print("Error at: ", e)
