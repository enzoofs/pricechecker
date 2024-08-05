import pandas as pd
import re
import fitz  # PyMuPDF
import os

# Função para leitura da planilha do fornecedor em formato CSV com delimitador correto
def read_supplier_sheet(file_path):
    try:
        df = pd.read_csv(file_path, delimiter=';')
        # Converter valores de vírgula para ponto
        df['Subtotal (USD)'] = df['Subtotal (USD)'].str.replace(',', '.').astype(float)
        print(f"Planilha {file_path} lida com sucesso.")
        print(f"Colunas encontradas: {df.columns}")
        return df
    except Exception as e:
        print(f"Erro ao ler a planilha {file_path}: {e}")
        return None

# Função para extração de referências e valores dos PDFs
def extract_reference_and_values_from_invoice(invoice_path):
    try:
        pdf_document = fitz.open(invoice_path)
        text = ""
        
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text += page.get_text()
        
        # Print the text extracted from the PDF
        print(f"Texto extraído da invoice {invoice_path}:\n{text[:1000]}...")  # print the first 1000 characters for brevity
        
        references = re.findall(r'REF# (\d+)', text)
        values = re.findall(r'(\d+,\d{2}) USD', text)
        values = [float(value.replace(',', '.')) for value in values]

        print(f"Referências extraídas da invoice {invoice_path}: {references}")
        print(f"Valores extraídos da invoice {invoice_path}: {values}")
        
        return references, values
    except Exception as e:
        print(f"Erro ao extrair referências e valores da invoice {invoice_path}: {e}")
        return [], []

# Função para comparar os valores extraídos com a planilha do fornecedor
def compare_values(supplier_df, invoice_references, invoice_values):
    discrepancies = []
    print(f"Comparando {len(invoice_references)} referências e {len(invoice_values)} valores...")
    for ref, inv_value in zip(invoice_references, invoice_values):
        try:
            supplier_row = supplier_df[supplier_df['Ref ID'] == int(ref)]
            if not supplier_row.empty:
                original_value = supplier_row['Subtotal (USD)'].values[0]
                print(f"Referência {ref}: valor original {original_value}, valor da invoice {inv_value}")
                if original_value != inv_value:
                    discrepancies.append({
                        'Ref ID': ref,
                        'Original Value': original_value,
                        'Invoice Value': inv_value
                    })
            else:
                print(f"Referência {ref} não encontrada na planilha.")
        except Exception as e:
            print(f"Erro ao comparar valores para a referência {ref}: {e}")
    return discrepancies

# Função para gerar o relatório em formato CSV
def generate_report(discrepancies, report_file):
    try:
        report_df = pd.DataFrame(discrepancies)
        report_df.to_csv(report_file, index=False)
        print(f"Relatório gerado em: {report_file}")
    except Exception as e:
        print(f"Erro ao gerar o relatório: {e}")

# Caminhos dos arquivos
supplier_file = 'C:/Python/PriceChecker/Automated Daily Shipment.csv'
invoice_files = ['C:/Python/PriceChecker/invoices/CI_0084111684.pdf', 'C:/Python/PriceChecker/invoices/CI_0084113590.pdf']
report_file = 'C:/Python/PriceChecker/discrepancies_report.csv'

# Leitura da planilha do fornecedor
supplier_df = read_supplier_sheet(supplier_file)

# Verificação das colunas necessárias
required_columns = ['Ref ID', 'Subtotal (USD)']
if supplier_df is not None:
    print(f'Colunas encontradas: {supplier_df.columns}')
    for col in required_columns:
        if col not in supplier_df.columns:
            print(f'Coluna necessária não encontrada: {col}')
            exit()

    # Extração e comparação dos dados
    all_discrepancies = []
    for invoice_file in invoice_files:
        if os.path.exists(invoice_file):
            invoice_references, invoice_values = extract_reference_and_values_from_invoice(invoice_file)
            discrepancies = compare_values(supplier_df, invoice_references, invoice_values)
            all_discrepancies.extend(discrepancies)
        else:
            print(f'Arquivo PDF não encontrado: {invoice_file}')

    # Gerar o relatório em formato CSV
    generate_report(all_discrepancies, report_file)
else:
    print("Erro na leitura da planilha, verifique o caminho do arquivo.")
