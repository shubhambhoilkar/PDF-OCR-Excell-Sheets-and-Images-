import re
import json
from PyPDF2 import PdfReader


# ---------- Utility: Extract text ----------
def extract_text(pdf_path: str) -> str:
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    return text


# ---------- PAYSLIP PARSER ----------
def parse_payslip(text: str) -> dict:
    payslip = {}

    fields = {
        "Name": r"Name\s+([A-Za-z\s]+)",
        "Passport": r"Passport\s*#\s*([A-Z0-9]+)",
        "Position": r"Position\s+([A-Za-z\s\-0-9]+)",
        "Project": r"Project\s+([A-Za-z\s\-0-9]+)",
        "JoiningDate": r"Joining Date\s*([0-9\-]+)",
        "GrossSalary": r"GROSS SALARY\s*([0-9,.]+)",
        "CalculatedSalary": r"CALCULATED SALARY\s*([0-9,.]+)",
        "NetSalary": r"NET SALARY PAY.*?\s*([0-9,.]+)"
    }

    for key, pattern in fields.items():
        m = re.search(pattern, text)
        payslip[key] = m.group(1).strip() if m else None

    period = re.search(r"From\s*([0-9\-]+)\s*To\s*([0-9\-]+)", text)
    if period:
        payslip["PeriodFrom"], payslip["PeriodTo"] = period.group(1), period.group(2)

    additions = re.search(r"ADDITION\s*\(\+\)\s*([0-9,.]+)", text)
    deductions = re.search(r"DEDUCTION\s*\(-\)\s*([0-9,.]+)", text)

    payslip["Additions"] = additions.group(1) if additions else "0.00"
    payslip["Deductions"] = deductions.group(1) if deductions else "0.00"

    # Attendance rows
    rows = []
    row_pattern = r"(\d{1,2})\s*-\s*([A-Za-z]{3})\s*([\dWALH]*)\s*([\d.]*)"
    for m in re.finditer(row_pattern, text):
        date, day, att, hrs = m.groups()
        rows.append({
            "Date": date,
            "Day": day,
            "Attendance": att or "",
            "Hours": hrs or ""
        })

    payslip["AttendanceTable"] = rows
    return payslip


# ---------- TIMESHEET PARSER ----------
def parse_timesheet(text: str) -> dict:
    result = {}

    emp_fields = {
        "Name": r"Name:\s*([A-Za-z\s]+)",
        "Designation": r"Designation:\s*([A-Za-z\s]+)",
        "Passport": r"PPT\.\s*No\.\:\s*([A-Z0-9]+)",
        "ProjectNo": r"Project No:\s*([A-Za-z0-9]+)",
        "PeriodFrom": r"(\d{1,2}-[A-Za-z]{3}-\d{2})\s*To:",
        "PeriodTo": r"To:\s*([0-9A-Za-z\-]+)"
    }

    for key, pattern in emp_fields.items():
        m = re.search(pattern, text)
        result[key] = m.group(1).strip() if m else None

    # Extract daily symbols: 9, W, AL, H, RR, etc.
    result["DailyStatus"] = re.findall(r"\b(\d+|W|H|AL|RR|SL|CL|UL)\b", text)

    total = re.search(r"Total\s*\(Hrs\)\s*([0-9]+)", text)
    result["TotalHours"] = total.group(1) if total else None

    return result


# ---------- AUTO DETECTION ----------
def detect_document_type(text: str) -> str:
    if "GROSS SALARY" in text or "NET SALARY" in text:
        return "Payslip"
    if "Manual Monthly Time Sheet" in text or "W W" in text or "AL" in text:
        return "Timesheet"
    return "Unknown"


# ---------- MAIN FUNCTION ----------
def extract_pdf_to_json(pdf_path: str, output_json: str):
    text = extract_text(pdf_path)
    doc_type = detect_document_type(text)

    if doc_type == "Payslip":
        parsed = parse_payslip(text)
    elif doc_type == "Timesheet":
        parsed = parse_timesheet(text)
    else:
        parsed = {"error": "Unknown PDF format"}

    final_json = {
        "DocumentType": doc_type,
        "Data": parsed
    }

    with open(output_json, "w") as f:
        json.dump(final_json, f, indent=4)

    return output_json


# ---------- Example RUN ----------
if __name__ == "__main__":
    output = extract_pdf_to_json(
        "C:\\Users\\Developer\\Shubham_files\\PDF-OCR\\Timesheets_AARTI.pdf",
        "extracted_data.json"
        )
    print("JSON created:", output)
