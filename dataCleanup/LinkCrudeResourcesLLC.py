import pandas as pd
import re
from datetime import datetime, timedelta
from pairings import get_name, get_pipeline

def extract_data_link_crude(sheet):
    (transaction_date, transaction_type, seller, buyer, pipeline, buyerAttn, sellerAttn, trader, 
    quantityA, quantityB, broker, brokerDocID, pricingDetail, pricingType, paymentTerm, 
    creditTerm, delivery_date_start, delivery_date_end, deliver_month, city, state, country, id_, company) = ("",) * 24
                    
    broker = 'LINK CRUDE RESOURCES,LLC'
    
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
                elif 'Seller Attn:' in cell:
                    sellerAttn = get_name(cell.split(':')[1].strip())
                elif 'Buyer Attn:' in cell:
                    buyerAttn = get_name(cell.split(':')[1].strip())
                elif 'Pipeline:' in cell:
                    pipeline = get_pipeline(cell.split(':')[1].strip())
                elif 'F.O.B.:' in cell:
                    city = cell.split(':')[1].strip()
                elif 'Total Volume:' in cell:
                    quantityA_str = cell.split(':')[1].strip()
                    quantityA_str = re.sub('[^\d,]', '', quantityA_str)
                    quantityA_str = quantityA_str.replace(',', '')
                    quantityA = int(quantityA_str)
                elif 'Barrels' in cell:
                    quantityB = 'BBL'
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
                    # Handles full month name and abbreviated month name
                    try:
                        delivery_date_start = datetime.strptime(delivery_month_year, '%B %Y')
                    except ValueError:
                        delivery_date_start = datetime.strptime(delivery_month_year, '%b %Y')
                    try:
                        deliver_month = datetime.strptime(delivery_month_year, '%B %Y')
                    except ValueError:
                        deliver_month = datetime.strptime(delivery_month_year, '%b %Y')
                    delivery_date_end = datetime(delivery_date_start.year, delivery_date_start.month, 1) + timedelta(days=32)
                    delivery_date_end = delivery_date_end.replace(day=1) - timedelta(days=1)
                elif 'Transaction #:' in cell:
                    brokerDocID = cell.split(':')[1].strip()

        if transaction_date and transaction_type and seller and buyer and pipeline and city and trader \
            and quantityA and quantityB and broker and brokerDocID and pricingDetail and pricingType and paymentTerm and creditTerm \
            and delivery_date_start and delivery_date_end and deliver_month:
            break
    if 'PetroChina International (America) Inc' in seller:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        seller = company
        buyer = buyer.upper()
        trader = sellerAttn
    elif 'PetroChina International (America) Inc' in buyer:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        buyer = company
        seller = seller.upper()
        trader = buyerAttn
    elif 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD' in seller:
        company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
        seller = company
        buyer = buyer.upper()
        trader = sellerAttn
    elif 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD' in buyer:
        company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
        buyer = company
        seller = seller.upper()
        trader = buyerAttn
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
    #print(f"Filtered data: \n{filtered_data}")
    #print(f"Data: \n{city, pipeline, 0, company}")

    if not filtered_data.empty:
        matched_row = filtered_data.iloc[0]
        state = matched_row['state']
        country = matched_row['country']
        location = f"{''.join(city)}, {''.join(state)}, {''.join(country)}"
        #print(location)
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
            id_ = 'no corresponding pipeline implis no correct id'
            pipeline = 'pipeline not found, broker pipeline did not match in database'

    return transaction_date, transaction_type, seller, buyer, pipeline, location, trader, quantityA, quantityB, broker, brokerDocID, \
        pricingDetail, pricingType, paymentTerm, creditTerm, delivery_date_start, delivery_date_end, id_
