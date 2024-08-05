import pandas as pd
import re
import fitz  # PyMuPDF
import os

# Function to read the supplier sheet in CSV format with the correct delimiter
def read_supplier_sheet(file_path):
    try:
        df = pd.read_csv(file_path, delimiter=';')
        # Convert values from comma to dot
        df['Subtotal (USD)'] = df['Subtotal (USD)'].str.replace(',', '.').astype(float)
        print(f"Sheet {file_path} read successfully.")
        print(f"Columns found: {df.columns}")
        return df
    except Exception as e:
        print(f"Error reading the sheet {file_path}: {e}")
        return None

# Function to extract references and values from PDFs
def extract_reference_and_values_from_invoice(invoice_path):
    try:
        pdf_document = fitz.open(invoice_path)
        text = ""
        
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text += page.get_text()
        
        # Print the text extracted from the PDF
        print(f"Text extracted from invoice {invoice_path}:\n{text[:1000]}...")  # print the first 1000 characters for brevity
        
        references = re.findall(r'REF# (\d+)', text)
        values = re.findall(r'(\d+,\d{2}) USD', text)
        values = [float(value.replace(',', '.')) for value in values]

        print(f"References extracted from invoice {invoice_path}: {references}")
        print(f"Values extracted from invoice {invoice_path}: {values}")
        
        return references, values
    except Exception as e:
        print(f"Error extracting references and values from invoice {invoice_path}: {e}")
        return [], []

# Function to compare the extracted values with the supplier sheet
def compare_values(supplier_df, invoice_references, invoice_values):
    discrepancies = []
    print(f"Comparing {len(invoice_references)} references and {len(invoice_values)} values...")
    for ref, inv_value in zip(invoice_references, invoice_values):
        try:
            supplier_row = supplier_df[supplier_df['Ref ID'] == int(ref)]
            if not supplier_row.empty:
                original_value = supplier_row['Subtotal (USD)'].values[0]
                print(f"Reference {ref}: original value {original_value}, invoice value {inv_value}")
                if original_value != inv_value:
                    discrepancies.append({
                        'Ref ID': ref,
                        'Original Value': original_value,
                        'Invoice Value': inv_value
                    })
            else:
                print(f"Reference {ref} not found in the sheet.")
        except Exception as e:
            print(f"Error comparing values for reference {ref}: {e}")
    return discrepancies

# Function to generate the report in CSV format
def generate_report(discrepancies, report_file):
    try:
        report_df = pd.DataFrame(discrepancies)
        report_df.to_csv(report_file, index=False)
        print(f"Report generated in: {report_file}")
    except Exception as e:
        print(f"Error generating the report: {e}")

# File paths
supplier_file = 'C:/Python/PriceChecker/Automated Daily Shipment.csv'
invoice_files = ['C:/Python/PriceChecker/invoices/CI_0084111684.pdf', 'C:/Python/PriceChecker/invoices/CI_0084113590.pdf']
report_file = 'C:/Python/PriceChecker/discrepancies_report.csv'

# Reading the supplier sheet
supplier_df = read_supplier_sheet(supplier_file)

# Check for required columns
required_columns = ['Ref ID', 'Subtotal (USD)']
if supplier_df is not None:
    print(f'Columns found: {supplier_df.columns}')
    for col in required_columns:
        if col not in supplier_df.columns:
            print(f'Required column not found: {col}')
            exit()

    # Extract and compare data
    all_discrepancies = []
    for invoice_file in invoice_files:
        if os.path.exists(invoice_file):
            invoice_references, invoice_values = extract_reference_and_values_from_invoice(invoice_file)
            discrepancies = compare_values(supplier_df, invoice_references, invoice_values)
            all_discrepancies.extend(discrepancies)
        else:
            print(f'PDF file not found: {invoice_file}')

    # Generate the report in CSV format
    generate_report(all_discrepancies, report_file)
else:
    print("Error reading the sheet, check the file path.")
