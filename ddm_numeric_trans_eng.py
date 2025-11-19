import pdfplumber
import pandas as pd
import unicodedata
import os
input_pdf = r"C:\Users\Ibrahim\Desktop\pdf\L.pdf"
output_xlsx = os.path.splitext(input_pdf)[0] + "_cleaned.xlsx"
DIGIT_MAP = str.maketrans("০১২৩৪৫৬৭৮৯", "0123456789")

def normalize_text(txt):
    """Normalize Unicode + convert Bangla digits to English."""
    if not isinstance(txt, str):
        return txt
    txt = unicodedata.normalize("NFC", txt)
    txt = txt.translate(DIGIT_MAP)
    return txt.strip()
all_tables, meta = [], []
with pdfplumber.open(input_pdf) as pdf:
    print(f"Total pages: {len(pdf.pages)}")
    for i, page in enumerate(pdf.pages, start=1):
        print(f" Processing page {i} ...")
        tables = page.extract_tables({
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines",
            "intersection_tolerance": 5,
            "snap_tolerance": 3,
            "join_tolerance": 3,
            "edge_min_length": 3,
            "text_tolerance": 2,
        })
        if not tables:
            tables = page.extract_tables({
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
                "text_tolerance": 3,
            })
        for t_id, table in enumerate(tables or [], start=1):
            df = pd.DataFrame(table)
            if df.empty:
                continue
            df = df.dropna(how="all", axis=0).dropna(how="all", axis=1)
            df = df.applymap(normalize_text)
            all_tables.append(df)
            meta.append({"page": i, "rows": df.shape[0], "cols": df.shape[1]})
if all_tables:
    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as w:
        for k, df in enumerate(all_tables, start=1):
            df.to_excel(w, index=False, sheet_name=f"Table_{k:03d}")
        pd.DataFrame(meta).to_excel(w, index=False, sheet_name="Summary")
    print(f" Saved {len(all_tables)} tables → {output_xlsx}")
else:
    print(" No tables detected in PDF.")
