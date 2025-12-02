class OCRParser:
    def __init__(self, name, identifier):
        self.identifier = identifier
        self.name = name
        self.master = {}

    def extract_timesheet(self, df):

        # -----------------------------------
        # Utility: find a row containing prefix
        # -----------------------------------
        def find_row_by_prefix(df, prefix):
            prefix = prefix.lower()
            for idx, row in df.iterrows():
                for cell in row:
                    text = str(cell).strip().lower()
                    if text.startswith(prefix):
                        return idx, str(cell)
            return None, None

        # -----------------------------------
        # Utility: find Job Code row
        # -----------------------------------
        def find_job_code_row(df):
            for idx, row in df.iterrows():
                if any(str(c).strip().lower().startswith("job code") for c in row):
                    cleaned = [str(v).strip() for v in row]
                    numeric_count = sum(v.isdigit() for v in cleaned[1:])
                    if numeric_count >= 5:      # valid job-code row
                        return idx, cleaned
            return None, None

        # -----------------------------------
        # Utility: extract aligned normal hours
        # -----------------------------------
        def extract_normal_hours(df, num_days):
            idx, row = find_row_by_prefix(df, "normal hours")
            if idx is None:
                return []

            # Clean entire row, skip label
            cleaned = [str(v).strip() for v in df.loc[idx].tolist()][1:]

            # Remove total hours if > 100
            if cleaned and cleaned[-1].isdigit() and int(cleaned[-1]) > 100:
                cleaned = cleaned[:-1]

            return cleaned[:num_days]



        # -----------------------------------
        # STEP 1: Extract Name + Designation
        # -----------------------------------

        # FIX: Ensure employee_df is a DataFrame (not Series)
        employee_df = df.iloc[1:3, [0]].reset_index(drop=True)

        name_idx, name_cell = find_row_by_prefix(employee_df, "name")
        designation_idx, designation_cell = find_row_by_prefix(employee_df, "designation")

        if name_cell:
            parts = name_cell.replace("Name:", "").strip().split("\n")
            name = " ".join(p.strip() for p in parts)
            self.master["TIMESHEET_EMPLOYEE_NAME"] = name

        if designation_cell:
            designation = designation_cell.split(":", 1)[1].strip()
            self.master["TIMESHEET_EMPLOYEE_DESIGNATION"] = designation

        # -----------------------------------
        # STEP 2: Find job code alignment row
        # -----------------------------------
        job_idx, job_row = find_job_code_row(df)
        if job_idx is None:
            return

        # Extract sequential days from job code row
        day_numbers = []
        for v in job_row[1:]:
            if v.isdigit():
                day_numbers.append(int(v))
            else:
                break

        num_days = len(day_numbers)

        # -----------------------------------
        # STEP 3: Extract Normal Hours using alignment
        # -----------------------------------
        normal_values = extract_normal_hours(df, num_days)

        # Save aligned values
        for day_index, value in enumerate(normal_values, start=1):
            self.master[f"TIMESHEET_DAY_{day_index}"] = value


    # -----------------------------------
    # Main Extract Wrapper
    # -----------------------------------
    def extract(self, dfs):
        self.master["EMPLOYEE_NAME"] = self.name
        self.master["PASSPORT_NUMBER"] = self.identifier

        # Timesheet → always first df
        self.extract_timesheet(dfs[0])

        return self.master
