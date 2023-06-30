import os
#import openpyxl
#import glob
import pdfplumber
import pandas as pd
from openpyxl import load_workbook
import shutil
from shutil import copy2
#import sys
#from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import re
from datetime import datetime, timedelta

# Define the names & locations of folders and format / data files
BASE_DIR = os.getcwd()
PDF_DIR = os.path.join(BASE_DIR, "pdfs")
DATA_DIR = os.path.join(BASE_DIR, "data")
RESULT_DIR = os.path.join(BASE_DIR, "results")
FORMAT_FILE = os.path.join(BASE_DIR, "format.xlsx")
physical_data_locations = os.path.join(BASE_DIR, "physical_data_locations.xlsx")
recognization = True

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

# Main body of finding each piece of necesary data and changing them to usable format
def extract_data(sheet):
    (transaction_date, transaction_type, seller, buyer, pipeline, sellerAttn, buyerAttn, 
    quantityA, quantityB, broker, brokerDocID, pricingDetail, pricingType, paymentTerm, 
    creditTerm, delivery_date_start, delivery_date_end, city, state, country, id_, company) = (None,) * 22

    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str):
                if 'Transaction Date:' in cell:
                    transaction_date = cell.split(':')[1].strip()
                elif 'Transaction Type:' in cell:
                    transaction_type = cell.split(':')[1].strip()
                    if transaction_type == 'Exchange':
                        transaction_type = 1
                    elif transaction_type == 'Outright':
                        transaction_type = 0
                    else: transaction_type = -1
                elif 'Seller:' in cell:
                    seller = cell.split(':')[1].strip()
                elif 'Buyer:' in cell:
                    buyer = cell.split(':')[1].strip()
                elif 'Pipeline:' in cell:
                    pipeline = get_pipeline(cell.split(':')[1].strip())
                elif 'F.O.B.:' in cell:
                    city = cell.split(':')[1].strip()
                elif 'Seller Attn:' in cell:
                    sellerAttn = get_name(cell.split(':')[1].strip())
                elif 'Buyer Attn:' in cell:
                    buyerAttn = get_name(cell.split(':')[1].strip())
                elif 'Total Volume:' in cell:
                    quantityA_str = cell.split(':')[1].strip()
                    quantityA_str = re.sub('[^\d,]', '', quantityA_str)
                    quantityA_str = quantityA_str.replace(',', '')
                    quantityA = int(quantityA_str)
                elif 'Barrels' in cell:
                    quantityB = 'BBL'
                elif 'LINK CRUDE RESOURCES, LLC' in cell:
                    broker = 'LINK CRUDE RESOURCES,LLC'
                elif 'Price US$/UNIT:' in cell:
                    pricing_str = cell.split(':')[1].strip()
                    if pricing_str.startswith('$'):
                        # It's a simple price, so remove the dollar sign and convert to float
                        pricing_str = pricing_str.replace('$', '')  
                        pricingDetail = float(pricing_str)
                        pricingType = 'Fixed'
                    else: continue
                elif 'PLUS $' in cell:
                    priceing_str_2 = cell.split('$')[1].strip().strip('.')
                    pricingDetail = 'Wti/EXCHANGE/NYMEX/1ST NRBY/CLOSE +' + priceing_str_2 + ' USD/BBL'
                    pricingType = 'CMA'
                elif 'MINUS $' in cell:
                    priceing_str_2 = cell.split('$')[1].strip().strip('.')
                    pricingDetail = 'Wti/EXCHANGE/NYMEX/1ST NRBY/CLOSE -' + priceing_str_2 + ' USD/BBL'
                    pricingType = 'CMA'
                elif 'BEFORE 20TH OF THE MONTH' in cell:
                    paymentTerm = '20 days after delivery month-end'
                elif 'BUYER\'S CREDIT IS SUBJECT TO SELLER\'S APPROVAL' in cell:
                    creditTerm = 'Seller\'s discretion'
                elif 'Delivery Date:' in cell:
                    delivery_month_year = cell.split(':')[1].strip()
                    delivery_date_start = datetime.strptime(delivery_month_year, '%B %Y')
                    delivery_date_end = datetime(delivery_date_start.year, delivery_date_start.month, 1) + timedelta(days=32)
                    delivery_date_end = delivery_date_end.replace(day=1) - timedelta(days=1)
                elif 'PetroChina International (America) Inc' in cell:
                    company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
                elif 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.' in cell:
                    company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
                elif 'Transaction #:' in cell:
                    brokerDocID = cell.split(':')[1].strip()

        if transaction_date and transaction_type and seller and buyer and pipeline and city and sellerAttn and buyerAttn and quantityA and quantityB \
            and broker and brokerDocID and pricingDetail and pricingType and paymentTerm and creditTerm and delivery_date_start and delivery_date_end:
            break
    
    # Change city name from HOUSTON to Houston, except for ECHO which is recorded as ECHO
    if city == 'ECHO':
        city == city
    else: city = city.title()

    physical_data_locations_df = pd.read_excel('physical_data_locations.xlsx')
    # Filter the data based on city, pipeline and status
    filtered_data = physical_data_locations_df[
        (physical_data_locations_df['city'] == city) &
        (physical_data_locations_df['pipeline_system'] == pipeline) &
        (physical_data_locations_df['status'] == 0) &
        (physical_data_locations_df['booking'] == company)
    ]
    print(f"Filtered data: \n{filtered_data}")
    #print(f"Data: \n{city, pipeline, 0, company}")

    if not filtered_data.empty:
        matched_row = filtered_data.iloc[0]
        state = matched_row['state']
        country = matched_row['country']
        id_ = matched_row['id']
        #print(f"Found matching id")
    else: 
        print(f"Potential Error")
        recognization = False
        filtered_data = physical_data_locations_df[
            (physical_data_locations_df['city'] == city) &
            (physical_data_locations_df['booking'] == company) &
            (physical_data_locations_df['status'] == 0)
        ]
        print(f"Filtered data: \n{filtered_data}")
        if not filtered_data.empty:
            matched_row = filtered_data.iloc[0]
            state = matched_row['state']
            country = matched_row['country']
            id_ = 'no corresponding pipeline implis no correct id'
            pipeline = 'pipeline not found, broker pipeline did not match in database'

    return transaction_date, transaction_type, seller, buyer, pipeline, city, sellerAttn, buyerAttn, quantityA, quantityB, broker, brokerDocID, \
        pricingDetail, pricingType, paymentTerm, creditTerm, delivery_date_start, delivery_date_end, state, country, id_

# Change name of employees to the default one recorded in the system
def get_name(input_str):
    name_to_value = {
        'Chuan Chen': 'chenchuan',
        'Nick Bugos': 'nicholasbugos',
        'Dan Dubeck': 'danieldubeck', 
        'Somename' : 'pennychin',
        'Somename' : 'justintodd',
        'Somename' : 'quynhtran',
        'Somename' : 'jjchen',
        'Somename' : 'yuridashko',
        'Somename' : 'ryanlowey',
        'Somename' : 'brycesturdy',
        'Somename' : 'jameshutchinson',
        'Somename' : 'oscarmarrero'
        # Add more names and values here...
    }

    for name, value in name_to_value.items():
        if name in input_str:
            return value
    recognization = False
    return 'Un-identified Trader'

# Change name of pipelines to the default one recorded in the system
def get_pipeline(input_str):
    name_to_value = {
        'ENBRIDGE TERMINAL': 'Enbridge',
        'ENTERPRISE': 'Enterprise',
        'ZYDECO': 'HOHO',
        'LOCAP': 'LOOP Pipeline',
        'MAGELLAN': 'Magellan East houston',
        'SEAWAY': 'Seaway',
        # Add more names and values here...
    }

    for name, value in name_to_value.items():
        if name in input_str:
            return value
    recognization = False
    return 'Un-identified Pipeline'

# Printing the data collected to the corresponding excel files, including constant data
def update_sheet(sheet, data):
    transaction_date, transaction_type, seller, buyer, pipeline, city, sellerAttn, buyerAttn, quantityA, quantityB, broker, brokerDocID, \
    pricingDetail, pricingType, paymentTerm, creditTerm, delivery_date_start, delivery_date_end, state, country, id_  = data

    sheet['B1'] = transaction_date
    sheet['B3'] = transaction_type
    if buyer is not None and 'PetroChina' in buyer:
        sheet['B2'] = 'Buy'
        sheet['B4'] = buyer
        sheet['B5'] = seller
        sheet['B14'] = buyerAttn
    else:
        sheet['B2'] = 'Sell'
        sheet['B4'] = seller
        sheet['B5'] = buyer
        sheet['B14'] = sellerAttn
        quantityA *= -1
    sheet['B7'] = ''
    sheet['B8'] = delivery_date_start.strftime('%m/%d/%Y').lstrip("0").replace("/0", "/")
    sheet['B9'] = delivery_date_end.strftime('%m/%d/%Y').lstrip("0").replace("/0", "/")
    sheet['B10'] = quantityA
    sheet['B11'] = quantityB
    sheet['B12'] = '±0%'
    sheet['B13'] = 'FIP'
    sheet['B15'] = 'Crude_AM'
    sheet['B16'] = ''
    sheet['B17'] = 'Pipeline'
    sheet['B18'] = f"{city}, {state}, {country}"    
    sheet['B19'] = pipeline
    sheet['B22'] = delivery_date_start.strftime('%Y-%m-%d')
    sheet['B23'] = pricingType
    sheet['B24'] = pricingDetail
    sheet['B25'] = paymentTerm
    sheet['B26'] = creditTerm
    sheet['B28'] = 'N/A'
    sheet['B29'] = 'No Inspection Fee Associated'
    sheet['B30'] = '0'
    sheet['B32'] = 'USD'
    sheet['B33'] = broker
    sheet['B34'] = brokerDocID
    sheet['B35'] = id_

def copy_format_to_sheet(format_file_name, sheet):
    format_book = load_workbook(format_file_name)
    format_sheet = format_book.active

    for i, column in enumerate(format_sheet.columns, start=1):
        column_letter = get_column_letter(i)
        if column_letter in format_sheet.column_dimensions:
            sheet.column_dimensions[column_letter].width = format_sheet.column_dimensions[column_letter].width

def cleanup_data(target_dir):
    files = load_files(target_dir)

    for file in files:
        book = load_workbook(os.path.join(target_dir, file))
        sheet1 = book.active  # Changes here
        data = extract_data(sheet1)

        if 'Sheet2' in book.sheetnames:
            sheet2 = book['Sheet2']
        else:
            sheet2 = book.create_sheet('Sheet2')

        copy_format_to_sheet('format.xlsx', sheet2)

        update_sheet(sheet2, data)

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

    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)

    for pdf in os.listdir(PDF_DIR):
        if pdf.endswith('.pdf'):
            pdf_path = os.path.join(PDF_DIR, pdf)
            text = extract_text_from_pdf(pdf_path)

            output_file = os.path.join(DATA_DIR, pdf.replace('.pdf', '.xlsx'))
            write_to_excel(text, output_file)

    cleanup_data(DATA_DIR)
    copy_files(DATA_DIR, RESULT_DIR)
    copy_format_to_each_file(RESULT_DIR, FORMAT_FILE)
    if recognization == True: print(f"Success! Trader and Pipeline all Recognized")
    else: print(f"Error! Trader or Pipeline may not be Completed")

if __name__ == "__main__":
    main()