from pymongo import MongoClient
from datetime import datetime

class MongoDB:
    def __init__(self, uri= "mongodb://localhost:27021/",db_name = "PDF_OCR"):
        self.client = MongoClient(uri)
        self.db = self.client[db_name]
        self.collection = self.db["employee_records"]

    def insert_record(self, timesheet_data, payslip_data, ts_file, ps_file):
        employee_id =  timesheet_data.get("IDENTIFIER") or payslip_data.get("PAYSLIP_BADGE")

        record = {
            "employee_id" : employee_id,
            "employee_name" : timesheet_data.get("NAME") or payslip_data.get("PAYSLIP_NAME"),

            "timesheet": timesheet_data,
            "payslip" : payslip_data,

            "source_files" :{
                "timesheet_pdf" : ts_file,
                "payslip_pdf" : ps_file
            },

            "created_at" :datetime.now()
            }
        return self.collection.insert_one(record)