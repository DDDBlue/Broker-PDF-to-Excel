import pandas as pd
import re
from datetime import datetime, timedelta
from pairings import get_name, get_pipeline, get_city, month_to_num

def extract_data_sage_refined(sheet):
    (transaction_date, transaction_type, seller, buyer, pipeline, trader, sellerAttn, buyerAttn, deliveryTerm, 
    quantityA, quantityB, quantityC, broker, brokerDocID, pricingDetail, pricingType, premium, paymentTerm, 
    creditTerm, delivery_date_start, delivery_date_end, city, state, location, country, id_, company, team, currency) = ("",) * 29
    broker = 'SAGE REFINED PRODUCTS LTD'
    currency = 'USD'
    creditTerm = 'Seller\'s discretion'
    paymentTerm = '2 business days after(start date=0) ROI (Receipt of Invoice)'
    transaction_type = '0'

    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str):
                if 'Deal ID:' in cell:
                    brokerDocID = cell.split(':')[1].strip()
                elif 'Transaction Date' in cell:
                    transaction_string = cell.split(':')[1].strip()
                    transaction_string = transaction_string.split(' ')
                    transaction_date = month_to_num(transaction_string[0]) + ' ' + transaction_string[1] + ' ' + transaction_string[2]
                    transaction_date = transaction_date.replace(' ', '/')  # Replace spaces with slashes
                    try:
                        datetime_obj = datetime.strptime(transaction_date, "%m/%d/%Y")  # Update the format to match new date string
                        transaction_date = datetime_obj.strftime("%m/%d/%Y")
                    except ValueError:
                        print(f"Could not parse date: {transaction_date}")
                        transaction_date = None  # or assign some default value
                elif 'Buyer:' in cell:
                    buyer = cell.split(':')[1].strip()
                elif 'Seller:' in cell:
                    seller = cell.split(':')[1].strip()
                elif 'Oscar Marrero' in cell:
                    trader = get_name(cell)
                elif 'James Hutchinson' in cell:
                    trader = get_name(cell)
                elif 'Justin Todd' in cell:
                    trader = get_name(cell)
                elif 'Location' in cell:
                    city = cell.split(':')[1].strip()
                    city = get_city(city.split(',')[0].strip())
                elif 'Delivered via:' in cell:
                    pipeline = cell.split(':')[1].strip()
                    pipeline = get_pipeline(pipeline)
                elif 'Quantity:' in cell:
                    quantityA = cell.split(':')[1].strip()
                    quantityA = quantityA.split(' ')[0].replace(',','')
                    quantityA = int (quantityA)
                elif 'Term:' in cell:
                    delivery_month_year = cell.split(':')[1].strip()
                    print(delivery_month_year)
                    try:
                        delivery_date_start = datetime.strptime(delivery_month_year, '%B %Y')
                    except ValueError:
                        print("The date format is incorrect.")
                    delivery_date_end = delivery_date_start + timedelta(days=32)
                    delivery_date_end = delivery_date_end.replace(day=1) - timedelta(days=1)
                elif 'Trade ID' in cell:
                    brokerDocID = cell.split(':')[1].strip()
                elif 'Price:' in cell:
                    pricingType = 'Fixed'
                    pricingDetail = cell.split('Price:')[1].strip()
                    pricingDetail = pricingDetail.split(' ')[0].strip()
                    pricingDetail = pricingDetail.replace('$', '')
                    pricingDetail = pricingDetail.replace('/gal', '')
                elif 'Pricing Info:' in cell:
                    pricingType = 'Average'
                    if 'EFP' in cell:
                        pricingType = 'EFP'
                    pricingDetail = cell.split(':')[1].strip()

        if transaction_date and transaction_type and seller and buyer and pipeline and city and trader and buyerAttn and sellerAttn \
            and quantityA and quantityB and broker and brokerDocID and pricingDetail and pricingType and paymentTerm and creditTerm \
            and delivery_date_start and delivery_date_end:
            break
    # Change city name from HOUSTON to Houston, except for ECHO which is recorded as ECHO
    if city == 'ECHO':
        city == city
    elif city == 'Houston':
        city = 'East Houston'
    elif city == 'Johnson\'s Corner':
        city = 'Johnsons Corner'
    elif city == 'Johnson\'S Corner':
        city = 'Johnsons Corner'
    else: city = city.title()
    if 'PetroChina International (America), Inc.' in seller:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        seller = company
        buyer = buyer.upper()
        team = 'Product_Light'
        quantityB = 'BBL'
        quantityC = '±5%'
        deliveryTerm = 'FIP'
    elif 'PetroChina International (America), Inc.' in buyer:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        buyer = company
        seller = seller.upper()
        team = 'Product_Light'
        quantityB = 'BBL'
        quantityC = '±5%'
        deliveryTerm = 'FIP'
    elif 'PetroChina International (Canada), Trading Ltd.' in seller:
        company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
        seller = company
        buyer = buyer.upper()
        team = 'Product_Light'
        quantityB = 'M3'
        quantityC = '±5%'
        deliveryTerm = 'EXPIPE'
    elif 'PetroChina International (Canada), Trading Ltd.' in buyer:
        company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
        buyer = company
        seller = seller.upper()
        trader = get_name(buyerAttn)
        team = 'Product_Light'
        quantityB = 'M3'
        quantityC = '±5%'
        deliveryTerm = 'EXPIPE'
    physical_data_locations_df = pd.read_excel('physical_data_locations.xlsx')
    # Filter the data based on city, pipeline and status
    filtered_data = physical_data_locations_df[
        (physical_data_locations_df['city'] == city) &
        (physical_data_locations_df['pipeline_system'] == pipeline) &
        (physical_data_locations_df['status'] == 0) &
        (physical_data_locations_df['booking'] == company)
    ]
    #print(f"Filtered data: \n{filtered_data}")
    print(f"Data: \n{city, pipeline, 0, company}")

    if not filtered_data.empty:
        matched_row = filtered_data.iloc[0]
        state = matched_row['state']
        country = matched_row['country']
        location = f"{''.join(city)}, {''.join(state)}, {''.join(country)}"
        id_ = matched_row['id']
        print(f"Found matching id", id_)
    else: 
        print(f"ID Not Found")
        filtered_data = physical_data_locations_df[
            (physical_data_locations_df['city'] == city) &
            (physical_data_locations_df['booking'] == company) &
            (physical_data_locations_df['status'] == 0)
        ]
        #print(f"Filtered data: \n{filtered_data}")
        if not filtered_data.empty:
            matched_row = filtered_data.iloc[0]
            state = matched_row['state']
            country = matched_row['country']
            location = f"{''.join(city)}, {''.join(state)}, {''.join(country)}"
        if not location:
            print(f"location not found, set to city which is {city}")
            location = city
        id_ = 'no corresponding pipeline implis no correct id'
        pipeline = 'pipeline not found, broker pipeline did not match in database'
    return transaction_date, transaction_type, seller, buyer, pipeline, location, trader, quantityA, quantityB, quantityC, broker, brokerDocID, \
        pricingDetail, pricingType, premium, paymentTerm, creditTerm, delivery_date_start, delivery_date_end, id_, team, currency, deliveryTerm or ""
