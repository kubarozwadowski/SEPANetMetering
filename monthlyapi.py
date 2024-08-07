import requests
import json
from typing import List, Dict

# EIA API Key
api_key = 'UhKbugjHqnLMfOetbnKlXUd1wVI42kvFjWV8Bdy6'
metadata = 'https://api.eia.gov/v2/electricity/state-electricity-profiles/net-metering?api_key=UhKbugjHqnLMfOetbnKlXUd1wVI42kvFjWV8Bdy6'
base_url = 'https://api.eia.gov/v2/electricity/state-electricity-profiles/net-metering/data'


def fetch_data(technology: str, start: int, end: int, state: str = None) -> dict:
    params = {
        'api_key': api_key,
        'data[0]': 'capacity',
        'data[1]': 'customers',
        'facets[technology][]': technology,
        'start': start,
        'end': end
    }
    
    if state:
        params['facets[state][]'] = state
    
    response = requests.get(base_url, params=params)
    if response.status_code == 200:
        return response.json()['response']['data']
    else:
        print(f"Error: Unable to fetch data (status code {response.status_code})")
        return []
    
#specific technology 1 year
#https://api.eia.gov/v2/electricity/state-electricity-profiles/net-metering/data?api_key=UhKbugjHqnLMfOetbnKlXUd1wVI42kvFjWV8Bdy6&data[0]=capacity&data[1]=customers&facets[technology][]=Photovoltaic&start=2016&end=2016
def technology_year(technology: str, year: int) -> dict:
    data = fetch_data(technology, year, year)
    
    sectors = {}
    for item in data:
        sector = item['sector']
        state = item['state']
        capacity = float(item['capacity']) if item['capacity'] is not None else 0.0
        customers = int(item['customers']) if item['customers'] is not None else 0
        
        if sector not in sectors:
            sectors[sector] = {
                'capacities': [],
                'customers': [],
                'data': [],
                'max_capacity': {'value': 0, 'state': None},
                'min_capacity': {'value': float('inf'), 'state': None},
                'max_customers': {'value': 0, 'state': None},
                'min_customers': {'value': float('inf'), 'state': None}
            }
        
        sectors[sector]['capacities'].append(capacity)
        sectors[sector]['customers'].append(customers)
        sectors[sector]['data'].append(item)
        
        if state != 'US':  # Ensure max/min values are not calculated for 'US'
            if capacity > sectors[sector]['max_capacity']['value']:
                sectors[sector]['max_capacity'] = {'value': capacity, 'state': state}
            if capacity < sectors[sector]['min_capacity']['value']:
                sectors[sector]['min_capacity'] = {'value': capacity, 'state': state}
            if customers > sectors[sector]['max_customers']['value']:
                sectors[sector]['max_customers'] = {'value': customers, 'state': state}
            if customers < sectors[sector]['min_customers']['value']:
                sectors[sector]['min_customers'] = {'value': customers, 'state': state}
    
    for sector in sectors:
        sectors[sector]['average_capacity'] = sum(sectors[sector]['capacities']) / len(sectors[sector]['capacities'])
        sectors[sector]['average_customers'] = sum(sectors[sector]['customers']) / len(sectors[sector]['customers'])
    
    return {
        'sectors': sectors,
        'metadata': {
            'capacity_units': data[0]['capacity-units'],
            'customers_units': data[0]['customers-units'],
            'technology': technology,
            'year': year
        }
    }

#specific technology range 
# #https://api.eia.gov/v2/electricity/state-electricity-profiles/net-metering/data?api_key=UhKbugjHqnLMfOetbnKlXUd1wVI42kvFjWV8Bdy6&data[0]=capacity&data[1]=customers&facets[technology][]=Photovoltaic&start=2016&end=2018   
def technology_range(technology: str, start: int, end: int) -> dict:
    data = fetch_data(technology, start, end)
    
    sectors = {}
    for item in data:
        sector = item['sector']
        state = item['state']
        year = item['period']
        capacity = float(item['capacity']) if item['capacity'] is not None else 0.0
        customers = int(item['customers']) if item['customers'] is not None else 0
        
        if sector not in sectors:
            sectors[sector] = {
                'capacities': [],
                'customers': [],
                'data': [],
                'max_capacity': {'value': 0, 'state': None},
                'min_capacity': {'value': float('inf'), 'state': None},
                'max_customers': {'value': 0, 'state': None},
                'min_customers': {'value': float('inf'), 'state': None}
            }
        
        sectors[sector]['capacities'].append(capacity)
        sectors[sector]['customers'].append(customers)
        sectors[sector]['data'].append(item)
        
        if state != 'US':  # Ensure max/min values are not calculated for 'US'
            if capacity > sectors[sector]['max_capacity']['value']:
                sectors[sector]['max_capacity'] = {'value': capacity, 'state': state}
            if capacity < sectors[sector]['min_capacity']['value']:
                sectors[sector]['min_capacity'] = {'value': capacity, 'state': state}
            if customers > sectors[sector]['max_customers']['value']:
                sectors[sector]['max_customers'] = {'value': customers, 'state': state}
            if customers < sectors[sector]['min_customers']['value']:
                sectors[sector]['min_customers'] = {'value': customers, 'state': state}
    
    for sector in sectors:
        sectors[sector]['average_capacity'] = sum(sectors[sector]['capacities']) / len(sectors[sector]['capacities'])
        sectors[sector]['average_customers'] = sum(sectors[sector]['customers']) / len(sectors[sector]['customers'])
    
    return {
        'sectors': sectors,
        'metadata': {
            'capacity_units': data[0]['capacity-units'],
            'customers_units': data[0]['customers-units'],
            'technology': technology,
            'start_year': start,
            'end_year': end
        }
    }

#specific technology + state 1 year
#https://api.eia.gov/v2/electricity/state-electricity-profiles/net-metering/data?api_key=UhKbugjHqnLMfOetbnKlXUd1wVI42kvFjWV8Bdy6&data[0]=capacity&data[1]=customers&facets[technology][]=Photovoltaic&start=2016&end=2016&facets[state][]=DC
def state_technology_year(technology: str, year: int, state: str) -> dict:
    data = fetch_data(technology, year, year, state)
    
    sectors = {}
    for item in data:
        sector = item['sector']
        capacity = float(item['capacity']) if item['capacity'] is not None else 0.0
        customers = int(item['customers']) if item['customers'] is not None else 0
        
        if sector not in sectors:
            sectors[sector] = {
                'capacities': [],
                'customers': [],
                'data': [],
                'max_capacity': {'value': 0, 'state': None},
                'min_capacity': {'value': float('inf'), 'state': None},
                'max_customers': {'value': 0, 'state': None},
                'min_customers': {'value': float('inf'), 'state': None}
            }
        
        sectors[sector]['capacities'].append(capacity)
        sectors[sector]['customers'].append(customers)
        sectors[sector]['data'].append(item)
        
        if capacity > sectors[sector]['max_capacity']['value']:
            sectors[sector]['max_capacity'] = {'value': capacity, 'state': state}
        if capacity < sectors[sector]['min_capacity']['value']:
            sectors[sector]['min_capacity'] = {'value': capacity, 'state': state}
        if customers > sectors[sector]['max_customers']['value']:
            sectors[sector]['max_customers'] = {'value': customers, 'state': state}
        if customers < sectors[sector]['min_customers']['value']:
            sectors[sector]['min_customers'] = {'value': customers, 'state': state}
    
    for sector in sectors:
        sectors[sector]['average_capacity'] = sum(sectors[sector]['capacities']) / len(sectors[sector]['capacities'])
        sectors[sector]['average_customers'] = sum(sectors[sector]['customers']) / len(sectors[sector]['customers'])
    
    return {
        'sectors': sectors,
        'metadata': {
            'capacity_units': data[0]['capacity-units'],
            'customers_units': data[0]['customers-units'],
            'technology': technology,
            'year': year,
            'state': state
        }
    }

#specific technology + state range
#https://api.eia.gov/v2/electricity/state-electricity-profiles/net-metering/data?api_key=UhKbugjHqnLMfOetbnKlXUd1wVI42kvFjWV8Bdy6&data[0]=capacity&data[1]=customers&facets[technology][]=Photovoltaic&start=2016&end=2018&facets[state][]=MA
def state_technology_range(technology: str, start: int, end: int, state: str) -> dict:
    data = fetch_data(technology, start, end, state)
    
    sectors = {}
    for item in data:
        sector = item['sector']
        year = item['period']
        capacity = float(item['capacity']) if item['capacity'] is not None else 0.0
        customers = int(item['customers']) if item['customers'] is not None else 0
        
        if sector not in sectors:
            sectors[sector] = {
                'capacities': [],
                'customers': [],
                'data': [],
                'max_capacity': {'value': 0, 'state': None},
                'min_capacity': {'value': float('inf'), 'state': None},
                'max_customers': {'value': 0, 'state': None},
                'min_customers': {'value': float('inf'), 'state': None}
            }
        
        sectors[sector]['capacities'].append(capacity)
        sectors[sector]['customers'].append(customers)
        sectors[sector]['data'].append(item)
        
        if capacity > sectors[sector]['max_capacity']['value']:
            sectors[sector]['max_capacity'] = {'value': capacity, 'state': state}
        if capacity < sectors[sector]['min_capacity']['value']:
            sectors[sector]['min_capacity'] = {'value': capacity, 'state': state}
        if customers > sectors[sector]['max_customers']['value']:
            sectors[sector]['max_customers'] = {'value': customers, 'state': state}
        if customers < sectors[sector]['min_customers']['value']:
            sectors[sector]['min_customers'] = {'value': customers, 'state': state}
    
    for sector in sectors:
        sectors[sector]['average_capacity'] = sum(sectors[sector]['capacities']) / len(sectors[sector]['capacities'])
        sectors[sector]['average_customers'] = sum(sectors[sector]['customers']) / len(sectors[sector]['customers'])
    
    return {
        'sectors': sectors,
        'metadata': {
            'capacity_units': data[0]['capacity-units'],
            'customers_units': data[0]['customers-units'],
            'technology': technology,
            'start_year': start,
            'end_year': end,
            'state': state
        }
    }


# Example usage
technology = 'Photovoltaic'
start = 2015
end = 2018
state = 'IN'
result = state_technology_range(technology, start, end, state)
print(result['sectors'].get('COM'))