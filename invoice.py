from pathlib import Path
import pandas as pd 
from docxtpl import DocxTemplate
from babel.numbers import format_currency
from docx2pdf import convert

base_dir = Path(__file__).parent
invoiceTemplate = input("Enter the template name ")
invoiceTemplatePath = base_dir/invoiceTemplate

excelPath = base_dir / "dummy_data.xlsx"

output_dir = base_dir / "invoices"
output_dir.mkdir(exist_ok=True)

invoice = pd.read_excel(excelPath, sheet_name="Sheet1")

##################### Invoice Start #######################
invoice = pd.DataFrame(invoice)
print(invoice)