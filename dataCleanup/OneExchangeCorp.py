import pandas as pd
import re
from datetime import datetime, timedelta
from pairings import get_name, get_pipeline

def extract_data_one_exchange(sheet):
    (transaction_date, transaction_type, seller, buyer, pipeline, trader, deliveryTerm, 
    quantityA, quantityB, quantityC, broker, brokerDocID, pricingDetail, pricingType, premium, paymentTerm, 
    creditTerm, delivery_date_start, delivery_date_end, city, state, location, country, id_, company, team, currency) = ("",) * 27
    broker = 'ONE EXCHANGE CORP.'
    deliveryTerm = 'EXPIPE'
    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str):
                if 'Trade Date:' in cell:
                    transaction_date = cell.split(':')[1].strip()
                elif 'Buyer:' in cell:
                    parts = cell.split("Seller:")
                    buyer = parts[0].replace("Buyer: ", "").strip()
                    seller = parts[1].strip()
                elif 'Trader:' in cell:
                    parts = cell.split(" Trader: ")
                    trader1 = get_name(parts[0].replace("Trader: ", "").strip())
                    trader2 = get_name(parts[1].strip())
                    if trader1 == 'Un-identified Trader':
                        trader = trader2
                    else: trader = trader1
                elif 'Transportation:' in cell:
                    pipeline = cell.split(":")[1].strip()
                    if 'Crude' in cell:
                        pipeline2 = pipeline.split('Crude')[0].strip()
                        pipeline = pipeline2
                    pipeline = get_pipeline(pipeline)
                elif 'Delivery Point:' in cell:
                    parts = cell.split(':')
                    city = parts[2].strip(" ")
                    if '|' in city:
                        city = city.split('| ')
                        city = city[1]
                elif 'Quantity:' in cell:
                    match = re.search("Quantity: ([\d,]+)", cell)
                    if match:
                        quantityA = int(match.group(1).replace(',', ''))
                    if 'M3' in cell:
                        quantityB = 'M3'
                    elif 'BBL' in cell:
                        quantityB = 'BBL'
                elif 'seller to pay' in cell:
                    creditTerm = 'Seller\'s discretion'
                elif 'Start Date:' in cell:
                    match_start = re.search("Start Date: ([\w\s\d,]+) End Date:", cell)
                    match_end = re.search("End Date: ([\w\s\d,]+)", cell)
                    if match_start and match_end:
                        date_start = match_start.group(1).strip()
                        date_end = match_end.group(1).strip()
                        try:
                            delivery_date_start = datetime.strptime(date_start, '%B %d, %Y')  # convert to datetime object
                            delivery_date_end = datetime.strptime(date_end, '%B %d, %Y')  # convert to datetime object
                        except ValueError:
                            print("The date format is incorrect.")
                elif 'Transaction ID:' in cell:
                    brokerDocID = cell.split(':')[1].strip()
                elif 'Price:' in cell:
                    match = re.search(r"Price: (-?\$[0-9.]+)", cell)
                    if match:
                        premium = float(match.group(1).replace('$', ''))
                    if 'USD/CAD' in cell:
                        currency = 'USD/CAD'
                    elif 'USD' in cell:
                        currency = 'USD'
                    elif 'CAD' in cell:
                        currency = 'CAD'
                elif 'Calendar average of' in cell:
                    pricingType = 'CMA'
                    premium = str(premium) + ' USD/BBL'
                    pricingDetail = 'Wti/EXCHANGE/NYMEX/1ST NRBY/CLOSE ' + premium
                elif 'Calendar Month Average' in cell:
                    pricingType = 'Complex'
                    pricingDetail = '-'
                    premium = ''
                elif 'Gibson Terminal' in cell:
                    pipeline = 'Gibson T19'

        if transaction_date and transaction_type and seller and buyer and pipeline and city and trader \
            and quantityA and quantityB and broker and brokerDocID and pricingDetail and pricingType and paymentTerm and creditTerm \
            and delivery_date_start and delivery_date_end:
            break
    # Change city name from HOUSTON to Houston, except for ECHO which is recorded as ECHO
    if city == 'ECHO':
        city == city
    elif city == 'Houston':
        city = 'East Houston'
    else: city = city.title()

    if 'PetroChina International (America) Inc' in seller:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        seller = company
        buyer = buyer.upper()
        team = 'Crude_AM'
        quantityC = '±0%'
    elif 'Petrochina International (America) Inc' == buyer:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        buyer = company
        seller = seller.upper()
        team = 'Crude_AM'
        quantityC = '±0%'
    elif 'PetroChina International (Canada) Trading Ltd.' in seller:
        company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
        seller = company
        buyer = buyer.upper()
        team = 'Crude_Canada'
        quantityC = '±5%'
    elif 'PetroChina International (Canada) Trading Ltd.' in buyer:
        company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
        buyer = company
        seller = seller.upper()
        team = 'Crude_Canada'
        quantityC = '±5%'
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
