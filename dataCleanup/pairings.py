import pandas as pd

import re
from datetime import datetime, timedelta

# Change name of employees to the default one recorded in the system
def get_name(input_str):
    name_to_value = {
        'Chuan Chen': 'chenchuan',
        'Nick Bugos': 'nicholasbugos',
        'Nicholas Bugos': 'nicholasbugos',
        'Dan Dubeck': 'danieldubeck',
        'Daniel Dubeck': 'danieldubeck',
        'Penny Chin': 'pennychin',
        'Justin Todd': 'justintodd',
        'Quynh Tran': 'quynhtran',
        'JJ Chen': 'jjchen',
        'Yuri Dashko': 'yuridashko',
        'Ryan Lowey': 'ryanlowey',
        'Bryce Sturdy': 'brycesturdy',
        'James Hutchinson': 'jameshutchinson',
        'Oscar Marrero': 'oscarmarrero',
        'Zhang Qing': 'zhangqing',
        'David Velasquez': 'davidvelasquez',
        'Justin Amoah': 'justinamoah'
        # Add more names and values here...
    }

    for name, value in name_to_value.items():
        if name in input_str:
            return value
    return 'Un-identified Trader'



# Change name of pipelines to the default one recorded in the system
def get_pipeline(input_str):
    name_to_value = {
        'ENBRIDGE TERMINAL': 'Enbridge',
        'NDPL': 'Enbridge North Dakota',
        'ENTERPRISE': 'Enterprise',
        'Enterprise Pipeline': 'Enterprise',
        'ZYDECO': 'HOHO',
        'LOCAP': 'LOOP Pipeline',
        'Loop': 'LOOP Pipeline',
        'MAGELLAN': 'Magellan East houston',
        'SEAWAY': 'Seaway',
        'LOOP': 'LOOP Pipeline',
        'Magellan/Houston': 'Magellan East houston',
        'Magellan Terminal': 'Magellan East houston',
        'Peace': 'Peace Pipeline',
        'AOSPL': 'Alberta Oil Sands Pipeline',
        'Pembina': 'Pembina Pipeline',
        'Enb T@S': 'Enbridge Transfer At Source',
        'ETAS': 'Enbridge Transfer At Source',
        'Transfer at Source': 'Enbridge Transfer At Source',
        'Enbridge Transfer': 'Enbridge Transfer At Source',
        'Enb TAS': 'Enbridge Transfer At Source',
        'Gibson': 'Gibson T19',
        'GT19': 'Gibson T19',
        'Fort Sask': 'Fort Sask Pipeline',
        'FSPL': 'Fort Sask Pipeline',
        'CLP-H': 'IPF Pipeline',
        'Cold Lake': 'IPF Pipeline',
        'Gibson Terminal': 'Gibson T19',
        'Mustang': 'SAX',
        'Market Link': 'Marketlink',
        'Federated': 'Swan Hills Pipeline',
        'Dakota Access Pipeline': 'DAPL',
        'Guernsey Hub': 'Guernsey HUB',
        'Guernsey': 'Guernsey HUB',
        'Colonial Pipeline-Non Alabama Origin': 'Colonial',
        'Colonial Pipeline': 'Colonial',

        # Add more names and values here...
    }

    for name, value in name_to_value.items():
        if name in input_str:
            return value
    return input_str


# Change name of city to the default one recorded in the system
def get_city(input_str):
    name_to_value = {
        'JOCO': 'Johnsons Corner',
        'Basis Linden': 'Linden',
        'Magellan East Houston': 'East Houston',
        # Add more names and values here...
    }

    for name, value in name_to_value.items():
        if name in input_str:
            return value
    return input_str

def month_to_num(month):
    month_dict = {
        'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 
        'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08', 
        'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
    }
    return month_dict.get(month, 'Invalid month')