import boto3
import os
import re
import tempfile

from ocr_parser1 import OCRParser
from PyPDF2 import PdfReader, PdfWriter

class OCRExtractor1:
    def __init__(self, region="ap-south-1"):
        self.textract = boto3.client("textract", region_name=region)

    def extract(self, file_path):
        first_page_pdf = self._get_first_page_pdf(file_path)

        tables = self.extract_tables(first_page_pdf)
        name, identifier = self.extract_name_identifier(file_path)

        extractor = OCRParser(name, identifier)
        result = extractor.extract(tables)
        os.remove(first_page_pdf)

        return result

    def _get_first_page_pdf(self, file_path):
        """Create a temporary PDF that contains only page 1."""
        reader = PdfReader(file_path)
        writer = PdfWriter()

        writer.add_page(reader.pages[0])
        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        with open(tmp_file.name, "wb") as f:
            writer.write(f)

        return tmp_file.name

    def extract_tables(self, file_path):
        with open(file_path, "rb") as f:
            content = f.read()

        response = self.textract.analyze_document(
            Document={"Bytes": content},
            FeatureTypes=["TABLES"]
        )
        return self._parse_tables(response)

    def extract_name_identifier(self, path: str):
        filename = os.path.basename(path)
        pattern = r"Timesheets_(.*?)\s+([A-Za-z0-9]+)\s+Timesheets"

        match = re.search(pattern, filename)
        if not match:
            return None, None

        return match.group(1).strip(), match.group(2).strip()

    def _parse_tables(self, response):
        blocks = response["Blocks"]
        block_map = {b["Id"]: b for b in blocks}
        tables = []
        table_blocks = [b for b in blocks if b["BlockType"] == "TABLE"]

        for table in table_blocks:
            rows = {}
            for rel in table.get("Relationships", []):
                if rel["Type"] != "CHILD":
                    continue

                for cell_id in rel["Ids"]:
                    cell = block_map[cell_id]
                    if cell["BlockType"] != "CELL":
                        continue

                    row = cell["RowIndex"]
                    col = cell["ColumnIndex"]

                    text = ""
                    if "Relationships" in cell:
                        for rrel in cell["Relationships"]:
                            if rrel["Type"] == "CHILD":
                                for wid in rrel["Ids"]:
                                    word_block = block_map[wid]
                                    if word_block["BlockType"] == "WORD":
                                        text += word_block.get("Text", "") + " "

                    rows.setdefault(row, {})
                    rows[row][col] = text.strip()

            sorted_rows = [rows[i] for i in sorted(rows.keys())]
            tables.append(sorted_rows)

        return tables
