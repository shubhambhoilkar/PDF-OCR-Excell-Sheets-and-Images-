import os 
import re
import fitz
import pandas as pd
from ocr_extractor1 import OCRExtractor1
from payslip_extractor import PaySlipExtractor

def is_image(pdf_path, text_threshold=50, word_threshold=20):
    doc = fitz.open(pdf_path)

    total_text = ""
    image_found = False

    for page in doc:
        text = page.get_text().strip()
        if text:
            total_text += " " + text

        if page.get_images():
            image_found = True

    text_len = len(total_text.strip())
    word_count = len(total_text.strip().split())
    return image_found and (text_len < text_threshold or word_count < word_threshold)

def extract_first_page(input_pdf, output_pdf):
    doc = fitz.open(input_pdf)
    new_doc = fitz.open()

    new_doc.insert_pdf(doc, from_page=0, to_page=0)
    new_doc.save(output_pdf)
    new_doc.close()
    doc.close()
    return output_pdf

def list_image_pdfs(folder_path):
    image_pdf_paths = []

    for file in os.listdir(folder_path):
        if file.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder_path, file)

            if is_image(pdf_path):
                image_pdf_paths.append(pdf_path)
    return image_pdf_paths

def find_corresponding_payslip(timesheet_file: str, payslip_folder: str):
    def extract_name_identifier(filename: str):
        pattern = r"(?:Timesheets_|Payslip_)(.*?)\s+([A-Za-z0-9]+)\s+(?:Timesheets|PAYSLIP)"
        match = re.search(pattern, filename, flags=re.IGNORECASE)
        if not match:
            return None, None
        return match.group(1).strip(), match.group(2).strip()

    ts_filename = os.path.basename(timesheet_file)
    name, identifier = extract_name_identifier(ts_filename)

    if not name or not identifier:
        return None
    for file in os.listdir(payslip_folder):
        if not file.lower().endswith(".pdf"):
            continue
        ps_name, ps_id = extract_name_identifier(file)
        if ps_id == identifier:
            return os.path.join(payslip_folder, file)
    return None

def save_to_excel(merged_dict, excel_path="merged_output.xlsx"):
    new_df = pd.DataFrame([merged_dict])

    if os.path.exists(excel_path):

        old_df = pd.read_excel(excel_path)
        all_columns = sorted(set(old_df.columns).union(set(new_df.columns)))

        old_df = old_df.reindex(columns=all_columns)
        new_df = new_df.reindex(columns=all_columns)
        
        final_df = pd.concat([old_df, new_df], ignore_index=True)

    else:
        # No existing file → create with sorted columns
        all_columns = sorted(new_df.columns)
        final_df = new_df.reindex(columns=all_columns)

    final_df.to_excel(excel_path, index=False)

    return excel_path

dir_path = r"output"
timesheets = r"timesheets"
payslips = r"payslip"

paths = list_image_pdfs(timesheets)

timesheet_extractor = OCRExtractor1()
payslip_extractor = PaySlipExtractor()

for path in paths:
    ps_path = find_corresponding_payslip(path, payslips)
    result1= timesheet_extractor.extract(path)
    result2 = payslip_extractor.extract(ps_path)
    combined = {**result1, **result2}
    # Save to Excel (existing)
    save_to_excel(combined, f"{dir_path}/generated_report.xlsx")

    # Save to MongoDB (new)
    db.insert_record(result1, result2, path, ps_path)
    print(f"Data Stored in MongoDB for: {path}")
