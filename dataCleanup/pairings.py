import pandas as pd

import re
from datetime import datetime, timedelta

# Change name of employees to the default one recorded in the system
def get_name(input_str):
    name_to_value = {
        'Chuan Chen': 'chenchuan',
        'Nick Bugos': 'nicholasbugos',
        'Dan Dubeck': 'danieldubeck',
        'Daniel Dubeck': 'danieldubeck',
        'Somename' : 'pennychin',
        'Somename' : 'justintodd',
        'Somename' : 'quynhtran',
        'Somename' : 'jjchen',
        'Somename' : 'yuridashko',
        'Somename' : 'ryanlowey',
        'Somename' : 'brycesturdy',
        'Somename' : 'jameshutchinson',
        'Somename' : 'oscarmarrero'
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
        'ENTERPRISE': 'Enterprise',
        'ZYDECO': 'HOHO',
        'LOCAP': 'LOOP Pipeline',
        'MAGELLAN': 'Magellan East houston',
        'SEAWAY': 'Seaway',
        'LOOP': 'LOOP Pipeline',
        'Magellan/Houston': 'Magellan East houston'
        # Add more names and values here...
    }

    for name, value in name_to_value.items():
        if name in input_str:
            return value
    return input_str