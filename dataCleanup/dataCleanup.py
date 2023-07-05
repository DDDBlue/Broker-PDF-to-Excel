import os
import pdfplumber
import pandas as pd
from openpyxl import load_workbook
import shutil
from shutil import copy2
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from LinkCrudeResourcesLLC import extract_data_link_crude
from CirtronCommoditiesLLC import extract_data_citron_commodities
from ModernCommoditiesINC import extract_data_modern_commodities

# Define the names & locations of folders and format / data files
BASE_DIR = os.getcwd()
PDF_DIR = os.path.join(BASE_DIR, "pdfs")
DATA_DIR = os.path.join(BASE_DIR, "data")
RESULT_DIR = os.path.join(BASE_DIR, "results")
FORMAT_FILE = os.path.join(BASE_DIR, "format.xlsx")
physical_data_locations = os.path.join(BASE_DIR, "physical_data_locations.xlsx")
recognization = {}

# Copy every excel from one folder to another
def copy_files(source_dir, target_dir):
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)
    
    files = [f for f in os.listdir(source_dir) if os.path.isfile(os.path.join(source_dir, f))]
    
    for file in files:
        if file.endswith('.xlsx'):
            copy2(os.path.join(source_dir, file), target_dir)

# Copy the format.xlsx file to each excel file
def copy_format_to_each_file(target_dir, format_file):
    files = [f for f in os.listdir(target_dir) if os.path.isfile(os.path.join(target_dir, f)) and f.endswith('.xlsx')]

    format_df = pd.read_excel(format_file)
    
    for file in files:
        book = load_workbook(os.path.join(target_dir, file))

        if 'Sheet2' in book.sheetnames:
            sheet = book['Sheet2']  # If 'Sheet2' already exists, select it
        else:
            sheet = book.create_sheet('Sheet2')  # Otherwise, create 'Sheet2'

        for row_index, row in format_df.iterrows():
            for column_index, cell_value in enumerate(row):
                # +1 is added to row and column indices because Openpyxl's index starts from 1, not 0
                sheet.cell(row=row_index+1, column=column_index+1, value=cell_value)
        
        book.save(os.path.join(target_dir, file))

# Helper function
def load_files(directory, extension=None):
    if extension:
        return [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f)) and f.endswith(extension)]
    else:
        return [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]


broker_to_function_map = {
    "Broker Link Crude": extract_data_link_crude,
    "Broker Citron Commodities": extract_data_citron_commodities,
    "Broker Modern Commodities": extract_data_modern_commodities,
    # Add more broker and values here...
}

def identify_broker(sheet):
    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str):
                if 'LINK CRUDE RESOURCES, LLC' in cell:
                    #print("found Broker Link Crude")
                    return 'Broker Link Crude'
                if 'None None' in cell:
                    #print("found Broker Citron Commodities")
                    return 'Broker Citron Commodities'
                if '' in cell:
                    return 'Broker Modern Commodities'
    # If no broker found, return None
    return None

def extract_data(sheet, file_name):
    brokerCompany = identify_broker(sheet)
    
    if brokerCompany is None:
        print(f"Could not identify broker in file: {file_name}")
        return None

    # Use the appropriate extraction function for the identified broker
    extraction_function = broker_to_function_map.get(brokerCompany)
    if extraction_function:
        return extraction_function(sheet)
    else:
        print(f"Could not find extraction function for broker: {brokerCompany}")


# Printing the data collected to the corresponding excel files, including constant data
def update_sheet(sheet, data, filename):
    transaction_date, transaction_type, seller, buyer, pipeline, location, trader, quantityA, quantityB, broker, brokerDocID, \
    pricingDetail, pricingType, paymentTerm, creditTerm, delivery_date_start, delivery_date_end, id_ = data

    sheet['B1'] = transaction_date or ""
    sheet['B3'] = transaction_type or ""
    #print(f"buyer is", buyer)
    #print(f"seller is", seller)
    if buyer and 'PETROCHINA' in buyer and broker == 'LINK CRUDE RESOURCES,LLC':
        #print("we are buying")
        sheet['B2'] = 'Buy'
        sheet['B4'] = buyer
        sheet['B5'] = seller
    elif buyer and 'PETROCHINA' in buyer and broker == 'CITRON COMMODITIES LLC':
        #print(f"we are buying, trader is", trader)
        sheet['B2'] = 'Buy'
        sheet['B4'] = buyer
        sheet['B5'] = seller
        sheet['B14'] = trader or ""
    else:
        #print("we are selling")
        sheet['B2'] = 'Sell'
        sheet['B4'] = seller
        sheet['B5'] = buyer
        sheet['B14'] = trader or ""
        quantityA = -1 * (quantityA or 0)

    try:
        sheet['B8'] = (delivery_date_start.strftime('%m/%d/%Y').lstrip("0").replace("/0", "/") if delivery_date_start else "")
        sheet['B9'] = (delivery_date_end.strftime('%m/%d/%Y').lstrip("0").replace("/0", "/") if delivery_date_end else "")
    except AttributeError:
        sheet['B8'] = ""
        sheet['B9'] = ""

    sheet['B10'] = quantityA or ""
    sheet['B11'] = quantityB or ""
    sheet['B12'] = '±0%'
    sheet['B13'] = 'FIP'
    sheet['B15'] = 'Crude_AM'
    sheet['B17'] = 'Pipeline'
    sheet['B18'] = location or ""  
    sheet['B19'] = pipeline or ""

    try:
        sheet['B22'] = (delivery_date_start.strftime('%Y-%m-%d') if delivery_date_start else "")
    except AttributeError:
        sheet['B22'] = ""

    sheet['B23'] = pricingType or ""
    sheet['B24'] = pricingDetail or ""
    sheet['B25'] = paymentTerm or ""
    sheet['B26'] = creditTerm or ""
    sheet['B28'] = 'N/A'
    sheet['B29'] = 'No Inspection Fee Associated'
    sheet['B30'] = '0'
    sheet['B32'] = 'USD'
    sheet['B33'] = broker or ""
    sheet['B34'] = brokerDocID or ""
    sheet['B35'] = id_ or ""
    
    if not sheet['B14'].value:
        recognization[filename] = False
    if not sheet['B35'].value:
        recognization[filename] = False


def copy_format_to_sheet(format_file_name, sheet):
    format_book = load_workbook(format_file_name)
    format_sheet = format_book.active

    for i, column in enumerate(format_sheet.columns, start=1):
        column_letter = get_column_letter(i)
        if column_letter in format_sheet.column_dimensions:
            format_width = format_sheet.column_dimensions[column_letter].width
            sheet.column_dimensions[column_letter].width = format_width

def cleanup_data(target_dir):
    files = load_files(target_dir)

    for file in files:
        book = load_workbook(os.path.join(target_dir, file))
        for sheet_name in book.sheetnames:
            sheet1 = book[sheet_name]
        if sheet1.max_column == 2:
            for row in sheet1.iter_rows(min_row=2, max_col=2, max_row=sheet1.max_row):
                cell_A = row[0]
                cell_B = row[1]
                cell_A.value = f"{cell_A.value} {cell_B.value}"
                cell_B.value = None
            book.save(os.path.join(target_dir, file))
        data_dict = extract_data(sheet1, file)

        if data_dict is None:
            print(f"Skipping file {file} due to broker identification issue.")
            continue  # Skip to the next file

        if 'Sheet2' in book.sheetnames:
            sheet2 = book['Sheet2']
        else:
            sheet2 = book.create_sheet('Sheet2')

        copy_format_to_sheet('format.xlsx', sheet2)
        update_sheet(sheet2, data_dict, file)
        book.save(os.path.join(target_dir, file))


# PDF Plumber that changes broker pdf to excel
def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = []
        for page in pdf.pages:
            page_text = page.extract_text()
            for line in page_text.split('\n'):
                text.append([line])  # wrap each line in a list to create a 2D list
    return text

# Helper function
def write_to_excel(data, output_file):
    df = pd.DataFrame(data, columns=['Text'])
    df.to_excel(output_file, index=False)

# Function that deletes the data & results folder each time before the new info is recorded
def remove_all_files_from_directory(directory):
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))

# Main(), where all the functions are called and handles the command line
# Simply run 'python dataCleanup.py'
def main():
    for directory in [DATA_DIR, PDF_DIR, RESULT_DIR]:
        if not os.path.exists(directory):
            os.makedirs(directory)
    remove_all_files_from_directory(DATA_DIR)
    remove_all_files_from_directory(RESULT_DIR)

    for file in os.listdir(PDF_DIR):
        if file.endswith('.pdf'):
            pdf_path = os.path.join(PDF_DIR, file)
            text = extract_text_from_pdf(pdf_path)

            output_file = os.path.join(DATA_DIR, file.replace('.pdf', '.xlsx'))
            write_to_excel(text, output_file)
        elif file.endswith('.xlsx'):
            original_wb = load_workbook(os.path.join(PDF_DIR, file))
            for i, sheet in enumerate(original_wb.sheetnames, start=1):
                wb = Workbook()
                new_sheet = wb.active
                for row in original_wb[sheet].iter_rows():
                    new_sheet.append((cell.value for cell in row))
                output_file = os.path.join(DATA_DIR, f"{file.replace('.xlsx', '')}_{i}.xlsx")
                wb.save(output_file)

    cleanup_data(DATA_DIR)
    copy_files(DATA_DIR, RESULT_DIR)
    copy_format_to_each_file(RESULT_DIR, FORMAT_FILE)
    print(DATA_DIR)
    if recognization:
        print(f"Error! Trader or Pipeline may not be Completed in these files: {recognization}")
    else:
        print(f"Success! Trader and Pipeline all Recognized")


if __name__ == "__main__":
    main()