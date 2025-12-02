class SelectableParser:
    def __init__(self, name, identifier):
        self.identifier = identifier
        self.name = name
        self.master = {}

    def find_corresonding_value(self, df, keyword, stop_tokens=None):
        keyword = keyword.lower().strip()
        stop_tokens = [s.lower() for s in (stop_tokens or [])]

        for _, row in df.iterrows():
            row_list = [str(c).strip() for c in row]
            lower_list = [c.lower() for c in row_list]

            for i, cell in enumerate(lower_list):
                if keyword in cell:
                    for right_cell in row_list[i + 1:]:

                        rc = right_cell.strip()
                        rc_lower = rc.lower()

                        if any(st in rc_lower for st in stop_tokens):
                            return None

                        if rc != "" and rc_lower != "none":
                            return rc

                    return None

        return None




    # -----------------------------------------------------------------
    # Utility: search full row across ALL columns
    # -----------------------------------------------------------------
    def find_row_by_prefix_anywhere(self, df, prefix):
        prefix = prefix.lower()
        for idx, row in df.iterrows():
            for cell in row:
                txt = str(cell).strip().lower()
                if txt.startswith(prefix):
                    return idx, str(cell)
        return None, None

    # -----------------------------------------------------------------
    # Utility: job code row (search across all columns)
    # -----------------------------------------------------------------
    def find_job_code_row(self, df):
        for idx, row in df.iterrows():
            cells = [str(c).strip().lower() for c in row]

            if not any(c.startswith("job code") for c in cells):
                continue

            cleaned = [str(v).strip() for v in row]

            # numeric day values (1..31)
            numeric_count = sum(v.isdigit() for v in cleaned)

            if numeric_count >= 5:   # valid row
                return idx, cleaned

        return None, None

    # -----------------------------------------------------------------
    # Utility: extract normal hours row correctly
    # -----------------------------------------------------------------
    def extract_normal_hours(self, df, num_days):
        idx, _row = self.find_row_by_prefix_anywhere(df, "normal hours")
        if idx is None:
            return []

        cleaned = [str(v).strip() for v in df.loc[idx].tolist()]

        # Remove label (“Normal Hours”)
        cleaned = cleaned[1:] if cleaned else cleaned

        # Remove total if numeric > 100
        if cleaned and cleaned[-1].isdigit() and int(cleaned[-1]) > 100:
            cleaned = cleaned[:-1]

        return cleaned[:num_days]

    # -----------------------------------------------------------------
    # MAIN TIMESHEET EXTRACTION
    # -----------------------------------------------------------------
    def extract_timesheet(self, df):

        # self.extract_name_and_designation(df)
        name = self.find_corresonding_value(df, "name")
        designation = self.find_corresonding_value(df, "designation")
        badge= self.find_corresonding_value(df, "B. No.")
        project = self.find_corresonding_value(df, "project")
        location = self.find_corresonding_value(df, "location")

        self.master["TIMESHEET_NAME"] = name
        self.master["TIMESHEET_DESIGNATION"] = designation
        self.master["TIMESHEET_BADGE_NUMBER"] = badge
        self.master["TIMESHEET_PROJECT"] = project
        self.master["TIMESHEET_LOCATION"] = location
        # --- Job Code Alignment Row ---
        job_idx, job_row = self.find_job_code_row(df)
        if job_idx is None:
            return

        # Parse day numbers
        day_numbers = []
        for v in job_row:
            if v.isdigit():
                day_numbers.append(int(v))
        num_days = len(day_numbers)

        # --- Normal Hours ---
        normal_values = self.extract_normal_hours(df, num_days)

        for day_index, value in enumerate(normal_values, start=1):
            self.master[f"TIMESHEET_DAY_{day_index}"] = value


    # -----------------------------------------------------------------
    # MAIN ENTRY POINT
    # -----------------------------------------------------------------
    def extract(self, dfs):
        self.master["EMPLOYEE_NAME"] = self.name
        self.master["PASSPORT_NUMBER"] = self.identifier

        # timesheet always first table
        self.extract_timesheet(dfs[0])

        return self.master
