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
        'Somename': 'pennychin',
        'Somename': 'justintodd',
        'Somename': 'quynhtran',
        'Somename': 'jjchen',
        'Yuri Dashko': 'yuridashko',
        'Ryan Lowey': 'ryanlowey',
        'Bryce Sturdy': 'brycesturdy',
        'Somename': 'jameshutchinson',
        'Somename': 'oscarmarrero',
        'Zhang Qing': 'zhangqing',
        'David Velasquez': 'davidvelasquez'
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
        'ZYDECO': 'HOHO',
        'LOCAP': 'LOOP Pipeline',
        'MAGELLAN': 'Magellan East houston',
        'SEAWAY': 'Seaway',
        'LOOP': 'LOOP Pipeline',
        'Magellan/Houston': 'Magellan East houston',
        'Peace': 'Peace Pipeline',
        'AOSPL': 'Alberta Oil Sands Pipeline',
        'Pembina': 'Pembina Pipeline',
        'Enb T@S': 'Enbridge Transfer At Source',
        'ETAS': 'Enbridge Transfer At Source',
        'Transfer at Source': 'Enbridge Transfer At Source',
        'Gibson': 'Gibson T19',
        'GT19': 'Gibson T19',
        'Fort Sask': 'Fort Sask Pipeline',
        'CLP-H': 'IPF Pipeline',
        'Gibson Terminal': 'Gibson T19',
        'Mustang': 'SAX',
        'Market Link': 'Marketlink'

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
        # Add more names and values here...
    }

    for name, value in name_to_value.items():
        if name in input_str:
            return value
    return input_str