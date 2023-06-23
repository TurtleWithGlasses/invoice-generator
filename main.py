import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_number, invoice_date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice: {invoice_number}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}", ln=1)

    # Add header
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]

    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Add rows
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(100,100,100)
        pdf.cell(w=30, h=8, txt=str(row[0]), border=1)
        pdf.cell(w=70, h=8, txt=str(row[1]), border=1)
        pdf.cell(w=30, h=8, txt=str(row[2]), border=1)
        pdf.cell(w=30, h=8, txt=str(row[3]), border=1)
        pdf.cell(w=30, h=8, txt=str(row[4]), border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")