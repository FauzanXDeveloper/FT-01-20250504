import pdfplumber
import pandas as pd
from openpyxl import Workbook

# -----------------------------
# 1. Paths
# -----------------------------
pdf_path = "CHAN WAI WENN.pdf"
excel_output = "extracted_tables.xlsx"

# -----------------------------
# 2. Extract tables using pdfplumber
# -----------------------------
tables = []

with pdfplumber.open(pdf_path) as pdf:
    for pnum, page in enumerate(pdf.pages):
        try:
            extracted = page.extract_tables()
            if extracted:
                for table in extracted:
                    df = pd.DataFrame(table)
                    tables.append((pnum, df))
        except Exception:
            pass

# -----------------------------
# 3. Write all tables to Excel
# -----------------------------
writer = pd.ExcelWriter(excel_output, engine='openpyxl')

for i, (pnum, df) in enumerate(tables):
    # Page number index + 1
    sheet_name = f"Page{pnum+1}_T{i+1}"
    df.to_excel(writer, sheet_name=sheet_name, index=False)

writer.close()

print("Extraction completed.")
print("Excel file saved as:", excel_output)
