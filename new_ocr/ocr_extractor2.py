from img2table.ocr import DocTR
from img2table.document import PDF

from main import result2
from ocr_parser2 import OCRParser
from ocr_parser3 import SelectableParser
import os
import re
from PyPDF2 import PdfReader, PdfWriter
import tempfile
import fitz
import pdfplumber
import pandas as pd

class OCRExtractor2:
    def __init__(self):
        self.ocr = DocTR()

    def is_image(self, pdf_path, text_threshold=30, word_threshold=20):
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

    def extract(self, file_path):
        first_page_pdf = self._get_first_page_pdf(file_path)
        name, identifier = self.extract_name_identifier(file_path)
        if self.is_image(first_page_pdf):
            tables = self.extract_tables_ocr(first_page_pdf)
            extractor = OCRParser(name, identifier)
            result = extractor.extract(tables)
        else:
            tables = self.extract_tables_selectable(first_page_pdf)
            extractor = SelectableParser(name, identifier)
            result = extractor.extract(tables)
        os.remove(first_page_pdf)
        return result

    def _get_first_page_pdf(self, file_path):
        reader = PdfReader(file_path)
        writer = PdfWriter()

        writer.add_page(reader.pages[0])
        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        with open(tmp_file.name, "wb") as f:
            writer.write(f)

        return tmp_file.name

    def extract_tables_ocr(self, file_path):
        pdf = PDF(file_path)
        tables = pdf.extract_tables(ocr=self.ocr)
        dfs = [t.df for t in tables[0]]
        return dfs

    def extract_tables_selectable(self, file_path):
        try:
            tables = []
            with pdfplumber.open(file_path) as pdf:
                first_page = pdf.pages[0]
                extracted = first_page.extract_tables()
                for table in extracted:
                    tables.append(pd.DataFrame(table))
            return tables
        except Exception as e:
            print("Error extracting tables:", e)
            return None


    def extract_name_identifier(self, path: str):
        filename = os.path.basename(path)
        pattern = r"Timesheets_(.*?)\s+([A-Za-z0-9]+)\s+Timesheets"
        match = re.search(pattern, filename)
        if not match:
            return None, None
        return match.group(1).strip(), match.group(2).strip()
