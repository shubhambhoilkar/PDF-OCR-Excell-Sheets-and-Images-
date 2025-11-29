from paddleocr import PaddleOCR
import layoutparser as lp
from PIL import Image
import json

# Load OCR models
ocr_table = PaddleOCR(show_log=False, structure=True)

table_detector = lp.Detectron2LayoutModel(
    "lp://TableBank/faster_rcnn_R_50_FPN",
    extra_config={"MODEL.ROI_HEADS.SCORE_THRESH_TEST": 0.5},
    label_map={0: "Table"}
)

def extract_key_value_pairs(table_data):
    """
    Convert PaddleOCR PP-Structure result into key-value pairs from image.
    """
    final_tables = []

    for entry in table_data:
        if "html" not in entry:
            continue

        from bs4 import BeautifulSoup
        soup = BeautifulSoup(entry["html"], "html.parser")

        rows = []
        for tr in soup.find_all("tr"):
            cols = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
            if cols:
                rows.append(cols)

        if not rows:
            continue

        header = rows[0]
        kv_list = [
            dict(zip(header, row))
            for row in rows[1:]
        ]
        final_tables.append(kv_list)

    return final_tables


def extract_table_from_image(img_path):
    img = Image.open(img_path).convert("RGB")

    # Detect table region
    layout = table_detector.detect(img)
    tables = [b for b in layout if b.type == "Table"]

    if not tables:
        return {"type": "no_table", "message": "No table detected in image"}

    # Extract full structured table
    result = ocr_table.ocr(img, cls=True)

    kv = extract_key_value_pairs(result)

    return {
        "type": "table",
        "raw_ocr": result,
        "key_value_pairs": kv
    }


if __name__ == "__main__":
    img_path = "table_image.jpg"

    output = extract_table_from_image(img_path)

    with open("image_table_output.json", "w", encoding="utf-8") as f:
        json.dump(output, f, indent=4, ensure_ascii=False)

    print("DONE. Saved to image_table_output.json")
