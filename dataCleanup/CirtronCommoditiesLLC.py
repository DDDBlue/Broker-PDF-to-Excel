import pandas as pd
import re
from datetime import datetime, timedelta
from pairings import get_name, get_pipeline

def extract_data_citron_commodities(sheet):
    (transaction_date, transaction_type, seller, buyer, pipeline, trader, 
    quantityA, quantityB, quantityC, broker, brokerDocID, pricingDetail, pricingType, premium, paymentTerm, 
    creditTerm, delivery_date_start, delivery_date_end, deliver_month, city, state, country, id_, company, team, currency) = ("",) * 26

    petrochina_found = False
    broker = 'CITRON COMMODITIES LLC'
    paymentTerm = '20 days after delivery month-end'
    currency = 'USD'

    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str):
                if 'Date:' in cell:
                    transaction_date = cell.split(':')[1].strip()
                elif 'Transaction Type:' in cell:
                    transaction_type = cell.split(':')[1].strip()
                    if transaction_type == 'Exchange':
                        transaction_type = 1
                    elif transaction_type == 'Outright':
                        transaction_type = 0
                    else: transaction_type = -1
                elif 'Seller: Petrochina International' in cell:
                    petrochina_found = True
                    seller = cell.split(':')[1].strip()
                elif 'Buyer: Petrochina International' in cell:
                    petrochina_found = True
                    buyer = cell.split(':')[1].strip()
                elif 'Seller:' in cell and seller == "":
                    seller = cell.split(':')[1].strip()
                elif 'Buyer:' in cell and buyer == "":
                    buyer = cell.split(':')[1].strip()
                elif 'Trader:' in cell and petrochina_found and trader == "":
                    trader = get_name(cell.split(':')[1].strip())
                elif 'Pipeline/Terminal' in cell:
                    #pipeline = cell.split('Pipeline/Terminal')[1].strip()
                    #print(pipeline)
                    pipeline = get_pipeline(cell.split('Pipeline/Terminal')[1].strip())
                    #print(pipeline)
                elif 'Delivery Location:' in cell:
                    location = cell.split(":")[1].strip()  # assuming location follows "Location:"
                    city = location.split(",")[0]  # get city before comma
                elif 'bpd' in cell:
                    quantityA_str = cell.split(':')[1].strip()
                    quantityA_str = re.sub('[^\d,]', '', quantityA_str)
                    quantityA_str = quantityA_str.replace(',', '')
                    quantityA = int(quantityA_str)
                    quantityB = 'BBL'
                elif 'EFP' in cell: 
                    pricingType = 'EFP'
                elif 'settlement price for the dates' in cell:
                    pricingType = 'Average'
                    if '+0.00' in cell: 
                        pricingDetail = 'Wti/ARGUS/CUSHING/SPOT01/CLOSE/Flat Price/Simple Average +0 USD/BBL'
                        premium = '0 USD/BBL'
                    elif '+0.01' in cell: 
                        pricingDetail = 'Wti/ARGUS/MidLand/SPOT01/CLOSE/Flat Price/Weighted Average +0.01 USD/BBL'
                        premium = '0.01 USD/BBL'
                elif 'Price: ' in cell:
                    if pricingType != 'Average' and pricingType != 'CMA': 
                        pricingType = 'Fixed'
                    price_pattern = r'(\d+\.\d+)'  # matches any number with a decimal in between
                    match = re.search(price_pattern, cell)
                    if match:
                        pricingDetail = float(match.group(1))  # convert the matched string to float
                    else:
                        pricingDetail = None
                elif 'BEFORE 20TH OF THE MONTH' in cell:
                    paymentTerm = '20 days after delivery month-end'
                elif 'seller to pay' in cell:
                    creditTerm = 'Seller\'s discretion'
                elif 'Timing:' in cell:
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

                    if delivery_date_start.month == 12:
                        next_month_first_day = delivery_date_start.replace(year=delivery_date_start.year+1, month=1, day=1)
                    else:
                        next_month_first_day = delivery_date_start.replace(month=delivery_date_start.month+1, day=1)

                   # Number of days in current month
                    number_of_days = (next_month_first_day - delivery_date_start).days
                elif 'Petrochina International (America) Inc' in cell:
                    company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
                elif 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.' in cell:
                    company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
                elif 'Transaction #:' in cell:
                    brokerDocID = cell.split(':')[1].strip()

        if transaction_date and transaction_type and seller and buyer and pipeline and city and trader \
            and quantityA and quantityB and broker and brokerDocID and pricingDetail and pricingType and paymentTerm and creditTerm \
            and delivery_date_start and delivery_date_end and deliver_month:
            break
    
    quantityA = quantityA * (number_of_days)
    # Change city name from HOUSTON to Houston, except for ECHO which is recorded as ECHO
    if city == 'ECHO':
        city == city
    elif city == 'Houston':
        city = 'East Houston'
    else: city = city.title()

    if 'Petrochina International (America), Inc.' == seller:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        seller = company
        buyer = buyer.upper()
        team = 'Crude_AM'
        quantityC = '±0%'
    elif 'Petrochina International (America), Inc.' == buyer:
        company = 'PETROCHINA INTERNATIONAL (AMERICA), INC.'
        buyer = company
        seller = seller.upper()
        team = 'Crude_AM'
        quantityC = '±0%'
    elif 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.' == seller:
        company = 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.'
        seller = company
        buyer = buyer.upper()
        team = 'Crude_Canada'
        quantityC = '±5%'
    elif 'PETROCHINA INTERNATIONAL (CANADA) TRADING LTD.' == buyer:
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
            id_ = 'no corresponding pipeline implis no correct id'
            pipeline = 'pipeline not found, broker pipeline did not match in database'

    return transaction_date, transaction_type, seller, buyer, pipeline, location, trader, quantityA, quantityB, quantityC, broker, brokerDocID, \
        pricingDetail, pricingType, premium, paymentTerm, creditTerm, delivery_date_start, delivery_date_end, id_, team, currency or ""
