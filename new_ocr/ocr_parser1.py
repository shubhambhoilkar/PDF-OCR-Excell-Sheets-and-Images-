import pandas as pd
from pathlib import Path
import re

class OCRParser:
    def __init__(self, name, identifier):
        self.identifier = identifier
        self.name = name
    # -------------------------------------------
    # CLASSIFY TABLE
    # -------------------------------------------
    def classify_timesheet_table(self, table):
        text = " ".join(str(v).lower() for row in table for v in row.values())

        # Table 0 usually contains personal info
        if "badge" in text and "employee" in text:
            return "employee"

        # Table 1: company info
        if "weekly working hours" in text or "hiring company" in text:
            return "company"

        # Actual Timesheet Table
        header_text = " ".join(str(v).lower() for v in table[1].values())
        if "job no" in header_text and "descr" in header_text or "time reporting" in header_text or "wbs" in header_text:
            return "timesheet"

        return "unknown"

    # -------------------------------------------
    # EMPLOYEE TABLE EXTRACTION
    # -------------------------------------------
    def extract_employee(self, table):
        out = {}
        for row in table:
            if 1 in row and 2 in row:
                key = str(row[1]).strip(" :").lower()
                val = row[2]
                if key.startswith("badge"):
                    out["TIMESHEET_BADGE_NUMBER"] = val
                if key.startswith("position"):
                    out["TIMESHEET_POSITION"] = val
                if key.startswith("department"):
                    out["TIMESHEET_DEPARTMENT"] = val
        return out

    # -------------------------------------------
    # COMPANY TABLE EXTRACTION
    # -------------------------------------------
    def extract_company(self, table):
        out = {}
        for row in table:
            if 1 in row and 2 in row:
                key = row[1].strip(" :").lower()
                val = row[2]
                if "weekly working hours" in key:
                    out["TIMESHEET_WEEKLY_WORKING_HOURS"] = val
                if "hiring company" in key:
                    out["TIMESHEET_HIRING_COMPANY"] = val
                if "cost center" in key:
                    out["TIMESHEET_COST_CENTER"] = val
                if "work location" in key:
                    out["TIMESHEET_WORK_LOCATION"] = val
        return out

    # -------------------------------------------
    # MAIN TIMESHEET EXTRACTION
    # -------------------------------------------
    def extract_timesheet(self, table):
        parsed = table
        WEEKDAYS = {"MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"}

        # 1. CLEAN INITIAL ROWS
        cleaned_rows = []
        for row in parsed:
            values = [str(v).strip().upper() for v in row.values()]
            if any(v in WEEKDAYS for v in values):
                continue
            if all(v == "" or v == "NONE" for v in values):
                continue
            cleaned_rows.append(row)

        df = pd.DataFrame(cleaned_rows)

        # 2. DETECT HEADER + END
        def is_header_row(row):
            header_regex = r"(?:JOB|DESCR|JOB\s*NO|TIME\s*REPORTING|REPORTING\s*CODE)"
            return row.astype(str).str.upper().str.strip().str.contains(header_regex, na=False).any()

        def is_end_row(row):
            end_regex = r"(?:PRESENCE|PRESENC)"
            return row.astype(str).str.upper().str.strip().str.contains(end_regex, na=False).any()

        header_idx = df.apply(is_header_row, axis=1).idxmax()
        end_idx = df.apply(is_end_row, axis=1).idxmax()

        # 3. CREATE JOB TABLE DATAFRAME
        timesheet_df = df.iloc[header_idx + 1 : end_idx].copy()
        timesheet_df.columns = df.iloc[header_idx]

        # remove all-dash empty rows
        EMPTY_TOKENS = {"", "-", "–", "—"}
        def is_empty_or_dash_row(row):
            return all(str(v).strip() in EMPTY_TOKENS for v in row)

        timesheet_df = timesheet_df[~timesheet_df.apply(is_empty_or_dash_row, axis=1)].reset_index(drop=True)

        summary_df = df.iloc[end_idx:, :2]

        # ----------------------
        # Extract Day-wise Hours
        # ----------------------
        job_no_col = timesheet_df.columns[0]
        job_desc_col = timesheet_df.columns[1]

        def parse_hours(val):
            val = str(val).strip()
            if val in ["", "-", "–", "—", "None", "nan"]:
                return None
            try:
                return float(val)
            except:
                return val

        day_cols = [col for col in timesheet_df.columns if re.fullmatch(r"\d{2}", str(col))]
        day_cols = sorted(day_cols, key=lambda x: int(x))

        result = {}

        for i, day_col in enumerate(day_cols, start=1):
            key = f"TIMESHEET_DAY_{i}"
            day_entries = []
            for _, row in timesheet_df.iterrows():
                hrs = parse_hours(row[day_col])
                if hrs is None:
                    continue
                day_entries.append({
                    "JOB_NO": str(row[job_no_col]).strip(),
                    "JOB_DESC": str(row[job_desc_col]).strip(),
                    "HOURS": hrs
                })
            result[key] = day_entries

        # ----------------------
        # Extract Summary Metrics
        # ----------------------
        for row in summary_df.itertuples(index=False):
            label = str(row[0]).strip().lower()
            value = row[1]

            mapping = {
                "presence": "TIMESHEET_PRESENCE_DAYS(P)",
                "sick": "TIMESHEET_SICK_LEAVE(S)",
                "vacation": "TIMESHEET_VACATION(V)",
                "mission days": "TIMESHEET_MISSION_DAYS(M)",
                "europe": "TIMESHEET_MISSION_IN_EUROPE(ME)",
                "offshore": "TIMESHEET_MISSION_OFFSHORE(MO)",
                "unpaid": "TIMESHEET_UNPAID_DAYS(U)",
                "travel": "TIMESHEET_TRAVEL_DAYS(T)",
                "working": "TIMESHEET_TOTAL_WORKING_HOURS",
                "normal": "TIMESHEET_TOTAL_NORMAL_HOURS",
                "weekdays": "TIMESHEET_TOTAL_OVERTIME_WEEKDAYS",
                "weekend": "TIMESHEET_TOTAL_OVERTIME_WEEKEND",
            }

            for key, out_key in mapping.items():
                if key in label:
                    result[out_key] = value

        return result

    # -------------------------------------------
    # PARSE ALL TIMESHEET TABLES
    # -------------------------------------------
    def extract(self, data):
        final = {
            "EMPLOYEE_NAME": self.name,
            "PASSPORT_NUMBER": self.identifier
        }

        for table in data:
            type_ = self.classify_timesheet_table(table)
            if type_ == "employee":
                final.update(self.extract_employee(table))
            elif type_ == "company":
                final.update(self.extract_company(table))
            elif type_ == "timesheet":
                final.update(self.extract_timesheet(table))


        return final
