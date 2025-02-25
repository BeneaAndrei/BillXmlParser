import ctypes
import os
import sys
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from tkinter import Tk, messagebox

# Define the namespace used in XML
namespace = {'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2',
             'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2'}

excel_data = []


def show_message_box(message, type):
    root = Tk()
    root.withdraw()
    if type == 'e':
        messagebox.showwarning("Error", message)
    if type == "i":
        messagebox.showinfo("Success!", message)


def parsing_function(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    supplier_data_parsing(root)
    invoice_general_data(root)
    customer_data_parsing(root)
    invoice_items(root)


def supplier_data_parsing(root):
    supplier_party_data = root.findall('.//cac:AccountingSupplierParty', namespace)

    for supplier_data in supplier_party_data:
        excel_data.append(supplier_data.find('.//cbc:RegistrationName', namespace).text)
        supplier_tax_scheme = supplier_data.find('.//cac:TaxScheme/cbc:ID', namespace)
        supplier_identification_code = supplier_data.find('.//cac:PartyTaxScheme/cbc:CompanyID', namespace)

        city = supplier_data.find('.//cbc:CityName', namespace).text if supplier_data.find('.//cbc:CityName',
                                                                                           namespace) is not None else ""
        county = supplier_data.find('.//cbc:CountrySubentity', namespace).text if supplier_data.find(
            './/cbc:CountrySubentity', namespace) is not None else ""

        excel_data.append(city)
        excel_data.append(county)

        if supplier_tax_scheme is not None:
            supplier_identification_number = f"{supplier_identification_code.text} ({supplier_tax_scheme.text})"
        else:
            supplier_tax_scheme = supplier_data.find('.//cac:PartyLegalEntity/cbc:CompanyID', namespace)
            supplier_identification_number = f"{supplier_tax_scheme.text}"

        excel_data.append(supplier_identification_number)


def customer_data_parsing(root):
    customer_party_data = root.findall('.//cac:AccountingCustomerParty', namespace)

    for customer_data in customer_party_data:
        excel_data.append(customer_data.find('.//cbc:RegistrationName', namespace).text)

        city = customer_data.find('.//cbc:CityName', namespace).text if customer_data.find('.//cbc:CityName',
                                                                                           namespace) is not None else ""
        county = customer_data.find('.//cbc:CountrySubentity', namespace).text if customer_data.find(
            './/cbc:CountrySubentity', namespace) is not None else ""

        excel_data.append(city)
        excel_data.append(county)

        customer_identification_code = customer_data.find('.//cbc:CompanyID', namespace).text
        excel_data.append(customer_identification_code)


def invoice_items(root):
    invoice_general_items = root.findall(".//cac:InvoiceLine", namespace)
    start_index_item = len(excel_data)

    for items in invoice_general_items:
        item_name = items.find(".//cac:Item/cbc:Name", namespace).text
        item_quantity = items.find(".//cbc:InvoicedQuantity", namespace).text
        item_quantity_unit = items.find(".//cbc:InvoicedQuantity", namespace).get('unitCode')
        item_percentage = items.find(".//cac:ClassifiedTaxCategory/cbc:Percent", namespace)
        item_price = items.find(".//cac:Price/cbc:PriceAmount", namespace).text
        item_total_price = items.find(".//cbc:LineExtensionAmount", namespace).text

        if item_percentage is not None:
            item_tva = round(float(item_total_price) * (float(item_percentage.text) / 100), 2)
        else:
            item_tva = 0

        item_data = [item_name, item_percentage.text if item_percentage is not None else "0", item_quantity,
                     item_quantity_unit, item_price, item_total_price, item_tva]

        end_index = start_index_item + len(item_data)
        excel_data[start_index_item:end_index] = item_data
        ws.append(excel_data)
    excel_data.clear()


def invoice_general_data(root):
    excel_data.append(root.find(".//cbc:ID", namespace).text)
    excel_data.append(root.find(".//cbc:IssueDate", namespace).text)
    excel_data.append("")  # Data Incarcare placeholder
    excel_data.append("")  # ID INCARCARE placeholder
    excel_data.append("")  # ID DESCARCARE placeholder
    excel_data.append(root.find(".//cbc:InvoiceTypeCode", namespace).text)  # TIP FACTURA


def create_excel():
    global wb, ws
    wb = Workbook()
    ws = wb.active

    first_row_headers = [
        'Furnizor', 'Localitate Furnizor', 'Judet Furnizor', 'Cod fiscal furnizor',
        'Numar factura', 'Data emiterii', 'Data incarcare', 'ID INCARCARE', 'ID DESCARCARE', 'TIP FACTURA',
        'Cumparator', 'Localitate Cumparator', 'Judet Cumparator', 'Cod fiscal Cumparator',
        'Denumire articol', 'Cota TVA', 'Cantitate', 'UM', 'Pret unitar', 'Total net', 'Valoare TVA'
    ]
    ws.append(first_row_headers)
    ws.append(list(map(str, range(1, len(first_row_headers) + 1))))


def parsing_all_xml_files(folder_path):
    error_flag = False
    create_excel()

    for filename in os.listdir(folder_path):
        if filename.endswith('.xml'):
            file_path = os.path.join(folder_path, filename)
            print(f"Processing {filename}...")
            try:
                parsing_function(file_path)
            except Exception as e:
                show_message_box(f"Error processing {file_path}: {str(e)}", 'e')
                error_flag = True
    try:
        wb.save('invoice_data.xlsx')
        if not error_flag:
            show_message_box("All done", 'i')
        else:
            show_message_box("Finished with some errors", 'i')
    except Exception:
        show_message_box("Error: Verify if the invoice excel is closed and try again", 'e')


parsing_all_xml_files('.')
