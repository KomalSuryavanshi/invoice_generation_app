import pandas as pd
import glob
import openpyxl
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")  # Use glob to read all xlsx files

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_no = filename.split("-")[0]   # filename.split("-") will give out ['10001','2023.1.18'] so add [0] to get just the 10001
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt= f"Invoice No.{invoice_no}")

    pdf.output(f"PDFs/{filename}.pdf")
