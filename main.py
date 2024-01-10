import os
import sys
import xml.etree.ElementTree as ET
from openpyxl import Workbook

# Define the namespace used in XML
namespace = {'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2',
             'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2'}

excel_data = []


def parsing_function(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    supplier_data_parsing(root)
    invoice_general_data(root)
    customer_data_parsing(root)
    invoice_items(root)


def supplier_data_parsing(root):
    # Find Registration name
    supplier_party_data = root.findall('.//cac:AccountingSupplierParty', namespace)

    for supplier_data in supplier_party_data:
        excel_data.append(supplier_data.find('.//cbc:RegistrationName', namespace).text)
        supplier_tax_scheme = supplier_data.find('.//cac:TaxScheme/cbc:ID', namespace)
        supplier_identification_code = supplier_data.find('.//cac:PartyTaxScheme/cbc:CompanyID', namespace)

    supplier_identification_number = f"{supplier_identification_code.text} ({supplier_tax_scheme.text})"

    excel_data.append(supplier_identification_number)


def customer_data_parsing(root):
    customer_party_data = root.findall('.//cac:AccountingCustomerParty', namespace)

    for customer_data in customer_party_data:
        excel_data.append(customer_data.find('.//cbc:RegistrationName', namespace).text)
        customer_identification_country = customer_data.find('.//cac:Country/cbc:IdentificationCode', namespace)
        customer_company_ID = customer_data.find(".//cbc:CompanyID", namespace)

    if not customer_company_ID.text.startswith("RO"):
        customer_identification_code = customer_identification_country.text + customer_company_ID.text
    else:
        customer_identification_code = customer_company_ID.text
    customer_tax_data = root.findall('.//cac:TaxTotal', namespace)

    for tax_data in customer_tax_data:
        customer_tax_scheme = tax_data.find('.//cac:TaxScheme/cbc:ID', namespace)

    customer_identification_number = f"{customer_identification_code} ({customer_tax_scheme.text})"
    excel_data.append(customer_identification_number)


def invoice_items(root):
    invoice_general_items = root.findall(".//cac:InvoiceLine", namespace)

    start_index_item = len(excel_data)

    invoice_list = []

    for items in invoice_general_items:
        item_name = items.find(".//cac:Item/cbc:Name", namespace)
        item_quantity = items.find(".//cbc:InvoicedQuantity", namespace)
        item_quantity_unit = items.find(".//cbc:InvoicedQuantity", namespace).get('unitCode')
        item_quantity_full = f"{item_quantity.text} {item_quantity_unit}"
        item_percentage = items.find(".//cac:ClassifiedTaxCategory/cbc:Percent", namespace)
        item_price = items.find(".//cac:Price/cbc:PriceAmount", namespace)
        item_price_currency = items.find(".//cac:Price/cbc:PriceAmount", namespace).get('currencyID')
        item_price_full = f"{item_price.text} {item_price_currency}"
        item_total_price = items.find(".//cbc:LineExtensionAmount", namespace)
        item_total_currency = items.find(".//cbc:LineExtensionAmount", namespace).get('currencyID')
        item_total_full = f"{item_total_price.text} {item_total_currency}"
        item_tva = round(float(item_total_price.text) * (float(item_percentage.text) / 100), 2)

        item_tva_full = str(item_tva) + " " + item_price_currency
        item_data = [item_name.text, item_percentage.text, item_quantity_full,
                     item_price_full, item_total_full, item_tva_full]
        end_index = start_index_item + len(item_data)
        excel_data[start_index_item:end_index] = item_data
        ws.append(excel_data)
    excel_data.clear()

        # if first:
            #excel_data.clear()
            #first = False
        #else:
            #invoice_list.append(item_data)

    #for item_data in invoice_list:
     #   row_data = [""] * start_index_item + item_data
     #   ws.append(row_data)


def invoice_general_data(root):
    excel_data.append(root.find(".//cbc:ID", namespace).text)
    excel_data.append(root.find(".//cbc:IssueDate", namespace).text)
    # Lipseste data incarcarii
    excel_data.append("")


def create_excel():
    # Create a new Excel workbook

    global wb
    wb = Workbook()
    global ws
    ws = wb.active

    # Set headers row in Excel
    first_row_headers = [
        'Furnizor', 'Cod fiscal furnizor', 'Numar factura', 'Data emiterii', 'Data incarcare',
        'Cumparator', 'Cod fiscal Cumparator', 'Denumire articol', 'Cota TVA', 'Cantitate',
        'Pret unitar', 'Total net', 'Valoare TVA'
    ]
    ws.append(first_row_headers)

    # Set second row headers with numeric values
    second_row_headers = list(map(str, range(1, len(first_row_headers) + 1)))
    ws.append(second_row_headers)


def parsing_all_xml_files(folder_path):
    # Iterate through files in the specified folder
    create_excel()
    for filename in os.listdir(folder_path):
        if filename.endswith('.xml'):  # Process only XML files
            file_path = os.path.join(folder_path, filename)
            print(f"Processing {filename}...")

            # Call your XML processing function here
            parsing_function(file_path)

    wb.save('invoice_data.xlsx')
    print("Data extracted and saved to 'invoice_data.xlsx'")


parsing_all_xml_files('facturi')
