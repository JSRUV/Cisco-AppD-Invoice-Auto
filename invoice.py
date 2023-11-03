from pathlib import Path
import pandas as pd 
from docxtpl import DocxTemplate
from babel.numbers import format_currency
from docx2pdf import convert

base_dir = Path(__file__).parent
invoiceTemplatePath = base_dir / "invoicetemp.docx"

excelPath = base_dir / "dummy_data.xlsx"

output_dir = base_dir / "invoices"
output_dir.mkdir(exist_ok=True)

invoice = pd.read_excel(excelPath, sheet_name="Sheet1")

##################### Invoice Start #######################
invoice = pd.DataFrame(invoice)
print(invoice)

formatted = invoice.loc[:,["client_name", "address", "state", "zipcode", 
                                    "unit_price", "sales_tax", "total", "qty", "item" ]]
formatted_invoices = formatted.to_dict(orient="records")
print(formatted_invoices)

for record in (formatted_invoices):
	doc = DocxTemplate(invoiceTemplatePath)
	doc.render(record)
	output_path = output_dir / f"{record['client_name']}-Invoice.docx"
	doc.save(output_path)
print("---------- Invoice DONE ------------")

convert("invoices/")