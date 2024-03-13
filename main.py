import ctypes
import os
import sys
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from tkinter import  Tk, messagebox
from tkinter.ttk import Progressbar, Label

# Define the namespace used in XML
namespace = {'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2',
             'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2'}

excel_data = []

def show_message_box(message, type):
    root = Tk()
    root.withdraw()  # Hide the root window

    # Show a message box with the given message
    if type == 'e':
        messagebox.showwarning("Error", message)
    if type == "i":
        messagebox.showinfo("Succes!", message)


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
        customer_identification_country = customer_data.find('.//cac:Country/cbc:IdentificationCode', namespace)
        customer_company_ID = customer_data.find(".//cbc:CompanyID", namespace)

    if not customer_company_ID.text.startswith("RO"):
        customer_identification_code = customer_identification_country.text + customer_company_ID.text
    else:
        customer_identification_code = customer_company_ID.text

    customer_tax_data = root.findall('.//cac:TaxTotal', namespace)

    for tax_data in customer_tax_data:
        customer_tax_scheme = tax_data.find('.//cac:TaxScheme/cbc:ID', namespace)

    if customer_tax_scheme is not None and customer_identification_code is not None:
        customer_identification_number = f"{customer_identification_code}"
    elif customer_tax_scheme is None and customer_identification_code is not None:
        customer_identification_number = f"{customer_identification_code.text}"
    elif customer_identification_code is None and customer_tax_scheme is not None:
        customer_identification_number = f"{customer_tax_scheme.text}"
    else:
        customer_identification_number = f" "

    excel_data.append(customer_identification_number)


def invoice_items(root):
    invoice_general_items = root.findall(".//cac:InvoiceLine", namespace)

    start_index_item = len(excel_data)

    invoice_list = []

    for items in invoice_general_items:
        item_name = items.find(".//cac:Item/cbc:Name", namespace)
        item_quantity = items.find(".//cbc:InvoicedQuantity", namespace)
        item_quantity_unit = items.find(".//cbc:InvoicedQuantity", namespace).get('unitCode')
        # item_quantity_full = f"{item_quantity.text} {item_quantity_unit}"
        item_percentage = items.find(".//cac:ClassifiedTaxCategory/cbc:Percent", namespace)
        item_price = items.find(".//cac:Price/cbc:PriceAmount", namespace)
        item_price_currency = items.find(".//cac:Price/cbc:PriceAmount", namespace).get('currencyID')
        item_price_full = f"{item_price.text} {item_price_currency}"
        item_total_price = items.find(".//cbc:LineExtensionAmount", namespace)
        item_total_currency = items.find(".//cbc:LineExtensionAmount", namespace).get('currencyID')
        item_total_full = f"{item_total_price.text} {item_total_currency}"

        if item_percentage is not None:
            item_tva = round(float(item_total_price.text) * (float(item_percentage.text) / 100), 2)
            item_tva_full = str(item_tva) + " " + item_price_currency
            item_data = [item_name.text, item_percentage.text, item_quantity.text, item_quantity_unit,
                         item_price_full, item_total_full, item_tva_full]

        else:
            item_tva_full = "0 RON"
            item_data = [item_name.text, "0", item_quantity.text, item_quantity_unit,
                         item_price_full, item_total_full, item_tva_full]

        end_index = start_index_item + len(item_data)
        excel_data[start_index_item:end_index] = item_data
        ws.append(excel_data)
    excel_data.clear()


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
        'UM', 'Pret unitar', 'Total net', 'Valoare TVA'
    ]
    ws.append(first_row_headers)

    # Set second row headers with numeric values
    second_row_headers = list(map(str, range(1, len(first_row_headers) + 1)))
    ws.append(second_row_headers)


def parsing_all_xml_files(folder_path):
    # Iterate through files in the specified folder
    error_flag = False
    create_excel()
    for filename in os.listdir(folder_path):
        if filename.endswith('.xml'):  # Process only XML files
            file_path = os.path.join(folder_path, filename)
            print(f"Processing {filename}...")

            try:
                # Call your XML processing function here
                parsing_function(file_path)
            except Exception as e:
                show_message_box("There was an error processing" + file_path + "with the error: " + str(e), 'e')
                print("\033[91m {}\033[00m".format(
                    "There was an error processing" + file_path + "with the error: " + str(e)))
                error_flag = True
    try:
        wb.save('invoice_data.xlsx')
        print("\033[92mData extracted and saved to 'invoice_data.xlsx'\033[00m")
        if (error_flag == False):
            show_message_box("All done", 'i')
        else:
            show_message_box("Finished with some errors", 'i')

    except Exception:
        print("\033[91m {}\033[00m".format("There was an error, verify if the invoice excel is closed and try again"))
        show_message_box("There was an error, verify if the invoice excel is closed and try again", 'e')


parsing_all_xml_files('.')
