import pandas as pd
import re
from datetime import datetime, timedelta
from pairings import get_name, get_pipeline, get_city

def extract_data_pvm_petroleum(sheet):
    (transaction_date, transaction_type, seller, buyer, pipeline, trader, sellerAttn, buyerAttn, deliveryTerm, 
    quantityA, quantityB, quantityC, broker, brokerDocID, pricingDetail, pricingType, premium, paymentTerm, 
    creditTerm, delivery_date_start, delivery_date_end, city, state, location, country, id_, company, team, currency) = ("",) * 29
    broker = 'PVM PETROLEUM MARKETS LLC'
    currency = 'USD'
    creditTerm = 'Seller\'s discretion'
    paymentTerm = '20 days after delivery month-end'

    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str):
                if 'Confirmation of Transaction' in cell:
                    brokerDocID = cell.split('Transaction')[1].strip()
                elif 'Deal Date: ' in cell:
                    transaction_date = cell.split(': ')[1].strip()
                    transaction_date = ' '.join(transaction_date.split()[:-1])  # removes the timezone part
                    try:
                        datetime_obj = datetime.strptime(transaction_date, "%Y-%m-%d %H:%M:%S")
                    except ValueError:
                        pass  # or assign some default value to datetime_obj
                    transaction_date = datetime_obj.strftime("%m/%d/%Y")
                elif 'Buyer:' in cell:
                    buyer = cell.split(': ')[1].strip()
                elif 'Seller:' in cell:
                    seller = cell.split(': ')[1].strip()
                elif 'To:' in cell:
                    trader = cell.split(': ')[1].strip()
                    trader = get_name(trader)
                elif 'Delivery Point: ' in cell:
                    city = cell.split(': ')[1].strip()
                    city = get_city(city.split(',')[0].strip())
                elif 'Total Quantity:' in cell:
                    quantityA = cell.split(': ')[1].strip()
                    quantityA = quantityA.split(' ')[0].replace(',','')
                    quantityA = int (quantityA)

                elif 'Period:' in cell:
                    period = cell.split(': ')[1].strip()
                    start_date_str, end_date_str = period.split(' through ')
                    try:
                        delivery_date_start = datetime.strptime(start_date_str, '%Y-%m-%d')
                        delivery_date_end = datetime.strptime(end_date_str, '%Y-%m-%d')
                    except ValueError:
                        print("The date format is incorrect.")
                elif 'Price:' in cell:
                    if 'Argus' in cell:
                        pricingType = 'Average'
                        city2 = city
                        if city2 == 'East Houston':
                            city2 = 'Houston'
                        if 'plus' in cell:
                            premium = cell.split('$')[1].strip()
                            premium = premium + ' USD/BBL'
                            pricingDetail = 'Wti/ARGUS/' + city2 + '/SPOT01/CLOSE/Flat Price/Weighted Average +' + premium
                        else:
                            premium = '0 USD/BBL'
                            pricingDetail = 'Wti/ARGUS/' + city2 + '/SPOT01/CLOSE/Flat Price/Simple Average +' + premium
                    else:
                        pricingDetail = cell.split(': ')[1].strip()
                        pricingDetail = pricingDetail.split(' ')[0].strip().replace('$','')
                        if pricingType == 'EFP':
                            pricingDetail = 'Fixed: ' + pricingDetail + ' USD/BBL   /'
                        else:
                            pricingType = 'Fixed'
                      
                elif 'EFP' in cell:
                    pricingType = 'EFP'

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
    if 'Petrochina International (America), Inc.' in seller:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        seller = company
        buyer = buyer.upper()
        team = 'Product_Light'
        quantityB = 'BBL'
        quantityC = '±0%'
        deliveryTerm = 'FIP'
    elif 'Petrochina International (America), Inc.' in buyer:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        buyer = company
        seller = seller.upper()
        team = 'Product_Light'
        quantityB = 'BBL'
        quantityC = '±0%'
        deliveryTerm = 'FIP'
    elif 'Petrochina International (Canada), Trading Ltd.' in seller:
        company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
        seller = company
        buyer = buyer.upper()
        team = 'Product_Light'
        quantityB = 'M3'
        quantityC = '±5%'
        deliveryTerm = 'EXPIPE'
    elif 'Petrochina International (Canada), Trading Ltd.' in buyer:
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
