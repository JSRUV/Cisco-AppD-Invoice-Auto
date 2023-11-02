from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate
from babel.numbers import format_currency
from docx2pdf import convert

data = {
    "client_name": ["Client A", "Client B", "Client C", "Client D", "Client E", "Client F", "Client G", "Client H", "Client I", "Client J"],
    "address": ["123 Main St", "456 Elm St", "789 Oak St", "101 Pine St", "202 Birch St", "303 Cedar St", "404 Walnut St", "505 Maple St", "606 Spruce St", "707 Fir St"],
    "state": ["CA", "NY", "TX", "FL", "IL", "PA", "OH", "NJ", "GA", "VA"],
    "zipcode": ["90210", "10001", "75001", "33101", "60601", "19103", "44101", "07001", "30301", "22301"],
    "unit_price": [100, 75, 50, 125, 60, 80, 90, 70, 110, 95],
    "sales_tax": [8, 6, 5, 10, 6, 8, 9, 7, 11, 9],
    "total": [108, 81, 55, 135, 66, 88, 99, 77, 121, 104],
    "qty": [2, 3, 1, 4, 2, 5, 1, 3, 2, 4],
    "item": ["Item 1", "Item 2", "Item 3", "Item 4", "Item 5", "Item 6", "Item 7", "Item 8", "Item 9", "Item 10"]
}

df = pd.DataFrame(data)

def create_invoice_documents(template_path, data, output_dir):
    print(f"------------ {output_dir.name} START -----------------")
    data_df = pd.DataFrame(data)

    # Format the data
    data_df["unit_price"] = data_df["unit_price"].apply(lambda x: format_currency(x, currency="USD", locale="en_US"))
    data_df["sales_tax"] = data_df["sales_tax"].apply(lambda x: format_currency(x, currency="USD", locale="en_US"))
    data_df["total"] = data_df["total"].apply(lambda x: format_currency(x, currency="USD", locale="en_US"))

    formatted_invoice = data_df.to_dict(orient="records")

    for record in formatted_invoice:
        doc = DocxTemplate(template_path)
        doc.render(record)
        output_path = output_dir / f"{record['client_name']}-invoice.docx"
        doc.save(output_path)

    print(f"---------- {output_dir.name} DONE ------------")

# Get user input for the template file and output folder
base_dir = Path(__file__).parent
invoice_temp = input("Enter the template filename: ")
invoice_template_path = base_dir / invoice_temp
excel_path = base_dir / "invoice-data.xlsx"
output_dir = base_dir / "Invoices"
output_dir.mkdir(exist_ok=True)

# Create and save invoice documents
create_invoice_documents(invoice_template_path, df, output_dir)

# Convert generated documents to PDF
convert("Invoices/")
