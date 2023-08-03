import pandas as pd
import re
from datetime import datetime, timedelta
from pairings import get_name, get_pipeline

def extract_data_tp_icap(sheet):
    (transaction_date, transaction_type, seller, buyer, pipeline, trader, deliveryTerm, commodity, 
    quantityA, quantityB, quantityC, broker, brokerDocID, pricingDetail, pricingType, premium, paymentTerm, 
    creditTerm, delivery_date_start, delivery_date_end, city, state, location, country, id_, company, team, currency) = ("",) * 28
    broker = 'TP ICAP (EUROPE)'
    currency = 'USD'
    creditTerm = 'Seller\'s discretion'
    paymentTerm = '5 business days after(start date=0) ROI (Receipt of Invoice)'
    commodity = 'placeholder'
    transaction_type = '0'
    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str):
                if 'Deal ID:' in cell:
                    brokerDocID = cell.split(':')[1].strip()
                elif 'Deal Date' in cell:
                    transaction_date = cell.split('Deal Date:')[1].strip()
                    transaction_date = ' '.join(transaction_date.split()[:-1])  # removes the timezone part
                    try:
                        datetime_obj = datetime.strptime(transaction_date, "%Y-%m-%d %H:%M:%S")
                    except ValueError:
                        pass  # or assign some default value to datetime_obj
                    transaction_date = datetime_obj.strftime("%m/%d/%Y")
                elif 'Buyer:' in cell:
                    buyer = cell.split(':')[1].strip()
                elif 'Seller:' in cell:
                    seller = cell.split(':')[1].strip()
                elif 'To:' in cell:
                    trader = get_name(cell.split(':')[1].strip())
                elif 'Product:' in cell:
                    if 'Propane' in cell:
                        commodity = 'PROPANE'
                    if 'Butane' in cell:
                        commodity = 'BUTANE'
                    if 'Ethane' in cell:
                        commodity = 'ETHANE'
                    if 'EPC' in cell:
                        pipeline = 'Enterprise'
                    if 'Enterprise' in cell:
                        pipeline = 'Enterprise'
                    if 'Energy Transfer' in cell:
                        pipeline = 'Lonestar'
                elif 'Location' in cell:
                    if 'See Notes' in cell:
                        pipeline = pipeline
                    else: 
                        pipeline = cell.split(':')[1].strip()
                elif 'Pipeline' in cell:
                    pipeline = cell.split('Pipeline ')[1].strip()
                    pipeline = get_pipeline(pipeline)
                elif 'Total Volume' in cell:
                    quantityA = quantityA
                elif 'Volume' in cell:
                    quantityA = cell.split(' ')[1].strip()
                    quantityA = quantityA.replace('m3','')
                    quantityA = quantityA.replace(',','')
                    quantityA = int (quantityA)
                elif 'Timing:' in cell:
                    dates = cell.split('Timing:')[1].split('through')  # splits the cell into start and end dates
                    delivery_date_start_str = dates[0].strip()  # strips any leading/trailing whitespace
                    delivery_date_end_str = dates[1].strip()
                    try:
                        delivery_date_start = datetime.strptime(delivery_date_start_str, '%Y-%m-%d')
                        delivery_date_end = datetime.strptime(delivery_date_end_str, '%Y-%m-%d')
                    except ValueError:
                        print("The date format is incorrect.")
                elif 'Trade ID' in cell:
                    brokerDocID = cell.split(':')[1].strip()
                elif 'Price' in cell:
                    pricingType = 'Fixed'
                    pricingDetail = cell.split('Price:')[1].strip()
                    pricingDetail = pricingDetail.split(' ')[0].strip()
                    pricingDetail = pricingDetail.replace('$', '')
                elif 'Additional Information:' in cell:
                    pricingType = 'Average'
                    premium = pricingDetail + ' USD/GAL'
                    pricingDetail = commodity + '/OPIS/EPC/SPOT01/CLOSE +' + premium

        if transaction_date and transaction_type and seller and buyer and pipeline and city and trader and commodity\
            and quantityA and quantityB and broker and brokerDocID and pricingDetail and pricingType and paymentTerm and creditTerm \
            and delivery_date_start and delivery_date_end:
            break
    # Change city name from HOUSTON to Houston, except for ECHO which is recorded as ECHO
    
    if pipeline == 'Enterprise':
        city = 'MONT BELVIEU-EPC'
    if pipeline == 'Lonestar':
        city = 'MONT BELVIEU-LST'
    if 'Petrochina International (America), Inc.' in seller:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        seller = company
        buyer = buyer.upper()
        team = 'PP_AM'
        quantityB = 'BBL'
        quantityC = 'EXACT'
        deliveryTerm = 'FOB'
    elif 'Petrochina International (America), Inc.' in buyer:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        buyer = company
        seller = seller.upper()
        team = 'PP_AM'
        quantityB = 'BBL'
        quantityC = 'EXACT'
        deliveryTerm = 'FOB'
    elif 'Petrochina International (Canada), Trading Ltd.' in seller:
        company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
        seller = company
        buyer = buyer.upper()
        team = 'PP_Canada'
        quantityB = 'BBL'
        quantityC = 'EXACT'
        deliveryTerm = 'FOB'
    elif 'Petrochina International (Canada), Trading Ltd.' in buyer:
        company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
        buyer = company
        seller = seller.upper()
        team = 'PP_Canada'
        quantityB = 'BBL'
        quantityC = 'EXACT'
        deliveryTerm = 'FOB'
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
