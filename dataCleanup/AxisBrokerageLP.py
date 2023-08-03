import pandas as pd
import re
from datetime import datetime, timedelta
from pairings import get_name, get_pipeline, get_city

def extract_data_axis_brokerage(sheet):
    (transaction_date, transaction_type, seller, buyer, pipeline, trader, sellerAttn, buyerAttn, deliveryTerm, 
    quantityA, quantityB, quantityC, broker, brokerDocID, pricingDetail, pricingType, premium, paymentTerm, 
    creditTerm, delivery_date_start, delivery_date_end, city, state, location, country, id_, company, team, currency) = ("",) * 29
    broker = 'AXIS BROKERAGE LP'
    currency = 'USD'
    creditTerm = 'Seller\'s discretion'
    paymentTerm = '20 days after delivery month-end'
    transaction_type = '1'
    perDay = False
    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str):
                if 'Confirmation of Transaction' in cell:
                    brokerDocID = cell.split('(')[1].strip()
                    brokerDocID = brokerDocID.replace(')','')
                elif 'Date : ' in cell:
                    transaction_string = cell.split(': ')[1].strip()
                    try:
                        datetime_obj = datetime.strptime(transaction_string, "%m/%d/%Y")  # Update the format to match new date string
                        transaction_date = datetime_obj.strftime("%m/%d/%Y")
                    except ValueError:
                        print(f"Could not parse date: {transaction_date}")
                elif 'Buyer :' in cell:
                    buyer = cell.split(': ')[1].strip()
                elif 'Seller :' in cell:
                    seller = cell.split(': ')[1].strip()
                elif 'To :' in cell:
                    trader = cell.split(': ')[1].strip()
                    trader = get_name(trader)
                elif 'FIP : ' in cell:
                    city = cell.split(': ')[1].strip()
                    city = get_city(city.split(',')[0].strip())
                elif 'Delivery Method : ' in cell:
                    pipeline = cell.split(': ')[1].strip()
                    pipeline = get_pipeline(pipeline)
                elif 'Volume :' in cell:
                    quantityA = cell.split(': ')[1].strip()
                    quantityA = quantityA.split(' ')[0].replace(',','')
                    quantityA = int (quantityA)
                    if 'per Day' in cell:
                        perDay = True
                elif 'Period :' in cell:
                    period = cell.split(': ')[1].strip()
                    start_date_str, end_date_str = period.split(' through ')
                    try:
                        delivery_date_start = datetime.strptime(start_date_str, '%m/%d/%Y')
                        delivery_date_end = datetime.strptime(end_date_str, '%m/%d/%Y')
                    except ValueError:
                        print("The date format is incorrect.")
                    if delivery_date_start.month == 12:
                        next_month_first_day = delivery_date_start.replace(year=delivery_date_start.year+1, month=1, day=1)
                    else:
                        next_month_first_day = delivery_date_start.replace(month=delivery_date_start.month+1, day=1)

                   # Number of days in current month
                    number_of_days = (next_month_first_day - delivery_date_start).days
                elif 'Trade ID' in cell:
                    brokerDocID = cell.split(': ')[1].strip()
                elif 'Price :' in cell:
                    if 'Argus' in cell:
                        pricingType = 'Average'
                        city2 = city
                        if city2 == 'East Houston':
                            city2 = 'Houston'
                        pricingDetail = 'Wti/ARGUS/' + city2 + '/SPOT01/CLOSE/Flat Price/Weighted Average +'
                        premium = '0 USD/BBL'
                        pricingDetail = pricingDetail + premium 
                    else:
                        pricingType = 'Fixed'
                        pricingDetail = cell.split(': ')[1].strip()
                        pricingDetail = pricingDetail.split(' ')[0].strip().replace('$','')
                elif '$+' in cell:
                    premium = cell.split('$')[1].strip()
                    premium = premium.split(' ')[0].strip()
                    premium = premium + ' USD/BBL'
                    pricingDetail = pricingDetail.replace('0 USD/BBL', '')
                    pricingDetail = pricingDetail + premium 

        if transaction_date and transaction_type and seller and buyer and pipeline and city and trader and buyerAttn and sellerAttn \
            and quantityA and quantityB and broker and brokerDocID and pricingDetail and pricingType and paymentTerm and creditTerm \
            and delivery_date_start and delivery_date_end:
            break
    if perDay == True:
        quantityA = quantityA * number_of_days
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
        quantityC = '±0%'
        deliveryTerm = 'FIP'
    elif 'PetroChina International (America), Inc.' in buyer:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        buyer = company
        seller = seller.upper()
        team = 'Product_Light'
        quantityB = 'BBL'
        quantityC = '±0%'
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
