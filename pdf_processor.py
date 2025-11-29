# Input: PDF

# Output:
# Detects text / image
# Detects table vs plain text inside image
# Extracts table
# Converts to key–value pairs
# 👉 Works on Python 3.12.

import pdfplumber
from pdf2image import convert_from_path
from paddleocr import PaddleOCR
import layoutparser as lp
from PIL import Image
import json

# ---------- CONFIG ----------
ocr_text = PaddleOCR(show_log=False)
ocr_table = PaddleOCR(show_log=False, structure=True)

table_detector = lp.Detectron2LayoutModel(
    "lp://TableBank/faster_rcnn_R_50_FPN",
    extra_config={"MODEL.ROI_HEADS.SCORE_THRESH_TEST": 0.5},
    label_map={0: "Table"}
)

# ---------- HELPERS ----------
def extract_key_value_pairs(table_data):
    """
    Convert PaddleOCR PP-Structure result into key-value pairs.
    """
    final_tables = []

    for entry in table_data:
        if "html" not in entry:
            continue

        # Extract from HTML table — rows and cells
        html_content = entry["html"]

        # simple HTML table extraction using BeautifulSoup (optional)
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html_content, "html.parser")

        rows = []
        for tr in soup.find_all("tr"):
            cols = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
            if cols:
                rows.append(cols)

        if not rows:
            continue

        header = rows[0]
        kv_list = []
        for row in rows[1:]:
            kv_list.append(dict(zip(header, row)))

        final_tables.append(kv_list)

    return final_tables


# ---------- MAIN PDF PROCESSOR ----------
def process_pdf(pdf_path):
    output = []
    pdf = pdfplumber.open(pdf_path)

    for page_index, page in enumerate(pdf.pages):
        print(f"Processing page {page_index+1}/{len(pdf.pages)}")

        text = page.extract_text()

        # ---------------------------
        # CASE 1: Page contains real text
        # ---------------------------
        if text and text.strip():
            output.append({
                "page": page_index + 1,
                "type": "text",
                "data": text
            })
            continue

        # ---------------------------
        # CASE 2: Page contains images (scanned page)
        # ---------------------------
        # Convert page → image
        pil_img = convert_from_path(
            pdf_path,
            first_page=page_index+1,
            last_page=page_index+1
        )[0]

        # Detect tables in image
        layout = table_detector.detect(pil_img)
        tables = [b for b in layout if b.type == "Table"]

        # ---------------------------
        # If table image
        # ---------------------------
        if tables:
            print(" → Detected table")
            result = ocr_table.ocr(pil_img, cls=True)
            kv = extract_key_value_pairs(result)
            output.append({
                "page": page_index + 1,
                "type": "table",
                "raw_ocr": result,
                "key_value_pairs": kv
            })
        else:
            # ---------------------------
            # If just text image
            # ---------------------------
            print(" → Detected text image")
            result = ocr_text.ocr(pil_img)
            text_data = "\n".join([line[1][0] for line in result])
            output.append({
                "page": page_index + 1,
                "type": "text_image",
                "data": text_data
            })

    return output


# ---------- RUN ----------
if __name__ == "__main__":
    pdf_path = "input.pdf"
    final_output = process_pdf(pdf_path)

    with open("output.json", "w", encoding="utf-8") as f:
        json.dump(final_output, f, indent=4, ensure_ascii=False)

    print("\nDONE. Saved to output.json")
