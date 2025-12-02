import re
import os
import pdfplumber
import pandas as pd

class PaySlipExtractor:
    def __init__(self):
        self.tables = None
        self.master = {}

    def extract_tables(self, path):
        try:
            tables = []
            with pdfplumber.open(path) as pdf:
                first_page = pdf.pages[0]
                extracted = first_page.extract_tables()
                for table in extracted:
                    tables.append(pd.DataFrame(table))
            return tables
        except Exception as e:
            print("Error extracting tables:", e)
            return None

    def get_adjacent_value(self, df, label_pattern):
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                cell = df.iat[r, c]
                if isinstance(cell, str) and re.search(label_pattern, cell, flags=re.I):
                    if c + 1 < df.shape[1]:
                        return df.iat[r, c + 1]
                    else:
                        return None
        return None

    def parse_ot(self, df):
        columns = df.iloc[0]
        df.columns = columns
        df = df[1:].reset_index(drop=True)
        df_cleaned = df.replace("", pd.NA).dropna(how="all")
        normal_ot = self.get_adjacent_value(df, "NORMAL OT")
        weekend_ot = self.get_adjacent_value(df, "WEEKEND OT")
        holiday_ot = self.get_adjacent_value(df, "HOLIDAY OT")
        total_ot = self.get_adjacent_value(df, "TOTAL DIRHAM")
        self.master["PAYSLIP_NORMAL_OT_THIS_MONTH"] = normal_ot
        self.master["PAYSLIP_WEEKEND_OT_THIS_MONTH"] = weekend_ot
        self.master["PAYSLIP_HOLIDAY_OT_THIS_MONTH"] = holiday_ot
        self.master["PAYSLIP_CURRENT_OT_THIS_MONTH"] = total_ot

    def parse_addition(self, df2):
        columns = df2.iloc[0]
        df2.columns = columns
        df2_cleaned = df2[1:].reset_index(drop=True)
        df2_cleaned = df2_cleaned.replace("", pd.NA).dropna(how="all").reset_index(drop=True)
        end_index = df2_cleaned[df2_cleaned["ADDITION"].str.contains("TOTAL DIRHAM", case = False, na = False)].index[0]
        additions = df2_cleaned.iloc[:end_index]["ADDITION"].to_list()
        addition_amounts = df2_cleaned.iloc[:end_index]["AMOUNT"].to_list()
        total_addition = df2_cleaned.loc[df2_cleaned["ADDITION"] == "TOTAL DIRHAM", "AMOUNT"].iloc[0]
        self.master["PAYSLIP_ADDITIONS"] = additions
        self.master["PAYSLIP_ADDITION_AMOUNT"] = addition_amounts
        self.master["PAYSLIP_ADDITION_TOTAL"] = total_addition

    def parse_deductions(self, df):
        columns = df.iloc[0]
        columns  = [i.replace('*', '').strip() for i in columns]
        df.columns = columns
        df_cleaned = df[1:].reset_index(drop=True)
        df_cleaned = df_cleaned.replace("", pd.NA).dropna(how="all").reset_index(drop=True)
        end_index = df_cleaned[df_cleaned["DEDUCTION"].apply(lambda x: re.sub(r"\s+", " ", re.sub(r"[^A-Z ]", " ", str(x).upper())).strip())
    .str.contains("TOTAL DIRHAM", case=False, na=False)].index[0]
        deductions = df_cleaned.iloc[:end_index]["DEDUCTION"].to_list()
        deduction_amounts = df_cleaned.iloc[:end_index]["AMOUNT"].to_list()
        total_deductions = df_cleaned.loc[df_cleaned["DEDUCTION"].str.contains("TOTAL", case=False, na =False)]["AMOUNT"].iloc[0]
        self.master["PAYSLIP_DEDUCTIONS"] = deductions
        self.master["PAYSLIP_DEDUCTION_AMOUNT"] = deduction_amounts
        self.master["PAYSLIP_DEDUCTION_TOTAL"] = total_deductions

    def parse_timesheet(self, df1):
        doj = self.get_adjacent_value(df1, r"Joining\s*Date")
        badge = self.get_adjacent_value(df1, r"Badge/ID")
        name = self.get_adjacent_value(df1, r"Name")
        passport = self.get_adjacent_value(df1, r"passport")
        position = self.get_adjacent_value(df1, r"position")
        project = self.get_adjacent_value(df1, r"project")
        bank_ac = self.get_adjacent_value(df1, r"bank")
        other_leaves = self.get_adjacent_value(df1, r"Other Leaves")
        non_paid_sl = self.get_adjacent_value(df1, r"Non Paid Sick")
        full_paid_sl = self.get_adjacent_value(df1, r"Fully Paid Sick")
        half_paid_sl = self.get_adjacent_value(df1, r"Half Paid Sick")
        authorised_absent = self.get_adjacent_value(df1, r"authorised absents")

        self.master["PAYSLIP_JOINING_DATE"] = doj
        self.master["PAYSLIP_PASSPORT"] = passport
        self.master["PAYSLIP_POSITION"] = position
        self.master["PAYSLIP_BADGE"] = badge
        self.master["PAYSLIP_NAME"] = name
        self.master["PAYSLIP_PROJECT"] = project
        self.master["PAYSLIP_BANK_ACCOUNT_NUMBER"] = bank_ac
        self.master["PAYSLIP_OTHER_LEAVES"] = other_leaves
        self.master["PAYSLIP_NON_PAID_SICK_LEAVES"] = non_paid_sl
        self.master["PAYSLIP_FULLY_PAID_SICK_LEAVES"] = full_paid_sl
        self.master["PAYSLIP_HALF_PAID_SICK_LEAVES"] = half_paid_sl
        self.master["PAYSLIP_AUTHORISED_ABSENTS"] = authorised_absent

        ts_start=df1[df1[0].str.contains(r"date - days", case=False, na=False)].index[0]
        ts_df = df1.iloc[ts_start:].reset_index(drop=True)
        columns = ts_df.iloc[0]
        ts_df.columns = columns
        ts_df = ts_df[1:].reset_index(drop=True)
        day_list = ts_df.loc[:,"Date - Days"].to_list()
        day_list = [i for i in day_list if  re.match(r"^\d{1,2} - [A-Za-z]{3}$", i)]

        for i,day in enumerate(day_list, start=1):
            value = ts_df.loc[ts_df["Date - Days"] == day, "Total Hrs."].values[0]
            self.master[f"PAYSLIP_DAY_{i}"] = value

    def parse_salary(self, df):
        gross_salary = self.get_adjacent_value(df, "GROSS SALARY")
        calculated_salary = self.get_adjacent_value(df, 'CALCULATED')
        fixed_ot = self.get_adjacent_value(df, 'FIXED OT')
        normal_ot = self.get_adjacent_value(df, 'NORMAL OT')
        weekend_ot = self.get_adjacent_value(df, 'WEEKEND OT')
        holiday_ot = self.get_adjacent_value(df, 'HOLIDAY OT')
        fully_paid_sl = self.get_adjacent_value(df, 'FULLY PAID SICK')
        half_paid_sl = self.get_adjacent_value(df, 'HALF PAID SICK')
        addition = self.get_adjacent_value(df, 'ADDITION')
        quarantine = self.get_adjacent_value(df, 'QUARANTINE')
        idle = self.get_adjacent_value(df, 'IDLE')
        deduction = self.get_adjacent_value(df, 'DEDUCTION')
        absent_deduction = self.get_adjacent_value(df, 'ABSENT')
        net_pay = self.get_adjacent_value(df, 'NET SALARY')

        self.master["PAYSLIP_GROSS_SALARY"] = gross_salary
        self.master["PAYSLIP_CALCULATED_SALARY"] = calculated_salary
        self.master["PAYSLIP_FIXED_OT_OF_THIS_MONTH"] = fixed_ot
        self.master["PAYSLIP_NORMAL_OT_OF_PREVIOUS_MONTH"] = normal_ot
        self.master["PAYSLIP_WEEKEND_OT_OF_PREVIOUS_MONTH"] = weekend_ot
        self.master["PAYSLIP_HOLIDAY_OT_OF_PREVIOUS_MONTH"] = holiday_ot
        self.master["PAYSLIP_FULLY_PAID_SICK_LEAVES"] = fully_paid_sl
        self.master["PAYSLIP_HALF_PAID_SICK_LEAVES"] = half_paid_sl
        self.master["PAYSLIP_ADDITION"] = addition
        self.master["PAYSLIP_QUARANTINE_SALARY"] = quarantine
        self.master["PAYSLIP_IDLE / STAND_BY_SALARY"] = idle
        self.master["PAYSLIP_DEDUCTION"] = deduction
        self.master["PAYSLIP_ABSENT_DEDUCTION"] = absent_deduction
        self.master["PAYSLIP_NET_SALARY_PAY(DIRHAM)"] = net_pay

    def extract(self, path):
        if not os.path.exists(path):
            raise FileNotFoundError(f"PDF file not found: {path}")
        tables = self.extract_tables(path)
        self.parse_timesheet(tables[0])
        self.parse_salary(tables[1])
        self.parse_ot(tables[2])
        self.parse_addition(tables[3])
        self.parse_deductions(tables[4])
        return self.master