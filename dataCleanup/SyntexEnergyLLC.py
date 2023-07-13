import pandas as pd
import re
from datetime import datetime, timedelta
from pairings import get_name, get_pipeline

def extract_data_syntex_energy(sheet):
    (transaction_date, transaction_type, seller, buyer, pipeline, trader, sellerAttn, buyerAttn, deliveryTerm, 
    quantityA, quantityB, quantityC, broker, brokerDocID, pricingDetail, pricingType, premium, paymentTerm, 
    creditTerm, delivery_date_start, delivery_date_end, city, state, location, country, id_, company, team, currency) = ("",) * 29
    broker = 'SYNTEXENERGY LLC'
    deliveryTerm = 'EXPIPE'
    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str):
                if 'Trade Time' in cell:
                    transaction_date = ' '.join(cell.split(':')[1].strip().split(' ')[:3])
                elif 'Buyer Company :' in cell:
                    buyer = cell.split(':')[1].strip()
                elif 'Seller Company :' in cell:
                    seller = cell.split(':')[1].strip()
                elif 'Seller Trader' in cell:
                    sellerAttn = get_name(cell.split(':')[1].strip())
                elif 'Buyer Trader' in cell:
                    buyerAttn = get_name(cell.split(':')[1].strip())
                elif 'Hub' in cell:
                    parts = cell.split(':')[1].strip()
                    pipeline = parts.split('-')[1].strip()
                    print(pipeline)
                    pipeline = get_pipeline(pipeline)
                    print(pipeline)
                    city = parts.split('-')[0].strip()
                elif 'Volume :' in cell:
                    quantityA = cell.split(':')[1].strip()
                    quantityA = quantityA.split(' ')[0].strip()
                    quantityA = int (quantityA)
                elif 'Term Start' in cell:
                    delivery_month_year = cell.split(':')[1].strip()
                    try:
                        # Adjust the format to match 'August 01, 2023'
                        delivery_date_start = datetime.strptime(delivery_month_year, '%B %d, %Y')
                    except ValueError:
                        print("The date format is incorrect.")
                        return
                    delivery_date_end = datetime(delivery_date_start.year, delivery_date_start.month, 1) + timedelta(days=32)
                    delivery_date_end = delivery_date_end.replace(day=1) - timedelta(days=1)
                elif 'Trade ID' in cell:
                    brokerDocID = cell.split(':')[1].strip()
                elif 'Price' in cell:
                    if 'USD' in cell:
                        currency = 'USD'
                    elif 'CAD' in cell:
                        currency = 'CAD'
                    premium = cell.split(':')[1].strip()
                    premium = premium + ' USD/BBL'
                elif 'Index' in cell:
                    if 'CMA' in cell: 
                        pricingType = 'CMA'
                        pricingDetail = 'Wti/EXCHANGE/NYMEX/1ST NRBY/CLOSE ' + premium
                elif 'Calendar Month Average' in cell:
                    pricingType = 'Complex'
                    pricingDetail = '-'
                    premium = ''

        if transaction_date and transaction_type and seller and buyer and pipeline and city and trader and buyerAttn and sellerAttn \
            and quantityA and quantityB and broker and brokerDocID and pricingDetail and pricingType and paymentTerm and creditTerm \
            and delivery_date_start and delivery_date_end:
            break
    # Change city name from HOUSTON to Houston, except for ECHO which is recorded as ECHO
    if city == 'ECHO':
        city == city
    elif city == 'Houston':
        city = 'East Houston'
    else: city = city.title()

    if 'PetroChina International (America), Inc.' in seller:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        seller = company
        buyer = buyer.upper()
        trader = sellerAttn
        team = 'Crude_AM'
        quantityB = 'BBL'
        quantityC = '±0%'
    elif 'PetroChina International (America), Inc.' == buyer:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        buyer = company
        seller = seller.upper()
        trader = buyerAttn
        team = 'Crude_AM'
        quantityB = 'BBL'
        quantityC = '±0%'
    elif 'PetroChina International (Canada) Trading Ltd.' in seller:
        company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
        seller = company
        buyer = buyer.upper()
        trader = sellerAttn
        team = 'Crude_Canada'
        quantityB = 'M3'
        quantityC = '±5%'
    elif 'PetroChina International (Canada) Trading Ltd.' in buyer:
        company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
        buyer = company
        seller = seller.upper()
        trader = buyerAttn
        team = 'Crude_Canada'
        quantityB = 'M3'
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
