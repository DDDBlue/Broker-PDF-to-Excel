import pandas as pd
import re
from datetime import datetime, timedelta
from pairings import get_name, get_pipeline, get_city

def extract_data_modern_commodities(sheet):
    (transaction_date, transaction_type, seller, buyer, pipeline, buyerAttn, sellerAttn, trader, 
    quantityA, quantityB, quantityC, broker, brokerDocID, pricingDetail, pricingType, premium, paymentTerm, 
    creditTerm, delivery_date_start, delivery_date_end, city, state, location, country, id_, company, team, currency) = ("",) * 29
                    
    broker = 'MODERN COMMODITIES INC.'
    paymentTerm = '20 days after delivery month-end'
    creditTerm = 'Seller\'s discretion'
    pricingDetail = 'Wti/EXCHANGE/NYMEX/1ST NRBY/CLOSE'
    currency = 'USD'

    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str):
                if 'Executed Timestamp :' in cell:
                    timestamp_str = cell.split(':')[1].strip()
                    # Try with full format first
                    try:
                        timestamp = datetime.strptime(timestamp_str, '%Y-%m-%d %I:%M:%S %p')
                    except ValueError:
                        # Try with date and hour if full format fails
                        try:
                            timestamp = datetime.strptime(timestamp_str, '%Y-%m-%d %H')
                        except ValueError:
                            # Fallback to date only if above format fails
                            timestamp = datetime.strptime(timestamp_str, '%Y-%m-%d')
                    month, day, year = timestamp.strftime('%m/%d/%Y').split('/')
                    transaction_date = f"{int(month)}/{day}/{year}"
                elif 'Transaction Type:' in cell:
                    transaction_type = cell.split(':')[1].strip()
                    if transaction_type == 'Exchange':
                        transaction_type = 1
                    elif transaction_type == 'Outright':
                        transaction_type = 0
                    else: transaction_type = -1
                elif 'Offer Legal Name :' in cell:
                    seller = cell.split(':')[1].strip()
                elif 'Bid Legal Name :' in cell:
                    buyer = cell.split(':')[1].strip()
                elif 'Offer Trader :' in cell:
                    sellerAttn = cell.split(':')[1].strip()
                elif 'Bid Trader : ' in cell:
                    buyerAttn = cell.split(':')[1].strip()
                elif 'Pipeline/Terminal : ' in cell:
                    pipeline = get_pipeline(cell.split(': ')[1].strip())
                    #print(pipeline)
                elif 'Location :' in cell:
                    city = get_city(cell.split(':')[1].strip())
                    #print(city)
                elif 'Volume :' in cell:
                    quantityA_str = cell.split(':')[1].strip()
                    quantityA = int(quantityA_str)
                elif 'bbls/day' in cell:
                    quantityB = 'BBL'
                    quantityA = quantityA * number_of_days
                elif 'CMA' in cell:
                    pricingType = 'CMA'
                elif 'Price :' in cell:
                    premium = cell.split(':')[1].strip()
                elif 'Term Start :' in cell:
                    delivery_month_year = cell.split(':')[1].strip()
                    try:
                        delivery_date_start = datetime.strptime(delivery_month_year, '%Y-%m-%d')
                        formatted_date_start = delivery_date_start.strftime('%m/%d/%Y')  # convert to 'M/D/YYYY' format
                    except ValueError:
                        print("The date format is incorrect.")
                        return
                    delivery_date_end = datetime(delivery_date_start.year, delivery_date_start.month, 1) + timedelta(days=32)
                    delivery_date_end = delivery_date_end.replace(day=1) - timedelta(days=1)

                    if delivery_date_start.month == 12:
                        next_month_first_day = delivery_date_start.replace(year=delivery_date_start.year+1, month=1, day=1)
                    else:
                        next_month_first_day = delivery_date_start.replace(month=delivery_date_start.month+1, day=1)

                   # Number of days in current month
                    number_of_days = (next_month_first_day - delivery_date_start).days
                elif 'Spread Trade Number :' in cell:
                    brokerDocID = brokerDocID
                elif 'Trade Number :' in cell:
                    brokerDocID = cell.split(':')[1].strip()

        if transaction_date and transaction_type and seller and buyer and pipeline and city and trader \
            and quantityA and quantityB and broker and brokerDocID and pricingDetail and pricingType and paymentTerm and creditTerm \
            and delivery_date_start and delivery_date_end:
            break
    if 'PetroChina International (America), Inc' in seller:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        seller = company
        buyer = buyer.upper()
        trader = get_name(sellerAttn)
        team = 'Crude_AM'
        quantityC = '±0%'
    elif 'PetroChina International (America), Inc' in buyer:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        buyer = company
        seller = seller.upper()
        trader = get_name(buyerAttn)
        team = 'Crude_AM'
        quantityC = '±0%'
    elif 'PETROCHINA INTERNATIONAL (CANADA), TRADING LTD' in seller:
        company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
        seller = company
        buyer = buyer.upper()
        trader = get_name(sellerAttn)
        team = 'Crude_Canada'
        quantityC = '±5%'
    elif 'PETROCHINA INTERNATIONAL (CANADA), TRADING LTD' in buyer:
        company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
        buyer = company
        seller = seller.upper()
        trader = get_name(buyerAttn)
        team = 'Crude_Canada'
        quantityC = '±5%'
    
    if pricingType == 'CMA' and quantityB == 'BBL':
        premium = str(premium) + ' USD/BBL'
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
        if not location:
            print(f"location not found, set to city which is {city}")
            location = city
        id_ = 'no corresponding pipeline implis no correct id'
        pipeline = 'pipeline not found, broker pipeline did not match in database'

    return transaction_date, transaction_type, seller, buyer, pipeline, location, trader, quantityA, quantityB, quantityC, broker, brokerDocID, \
        pricingDetail, pricingType, premium, paymentTerm, creditTerm, delivery_date_start, delivery_date_end, id_, team, currency or ""
