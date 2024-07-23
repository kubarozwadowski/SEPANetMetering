import requests
import json

# EIA API Key
api_key = 'UhKbugjHqnLMfOetbnKlXUd1wVI42kvFjWV8Bdy6'
metadata = 'https://api.eia.gov/v2/electricity/state-electricity-profiles/net-metering?api_key=UhKbugjHqnLMfOetbnKlXUd1wVI42kvFjWV8Bdy6'


def fetch_and_print_data(state, end_year, technology):
    url = 'https://api.eia.gov/v2/electricity/state-electricity-profiles/net-metering/data'
    params = {
        'api_key': api_key,
        'data[0]': 'capacity',
        'data[1]': 'customers',
        'facets[state][]': state,
        'facets[technology][][]': technology,
        'end': end_year
    }

    print("Requesting data with parameters:", params)
    
    response = requests.get(url, params=params)
    print("Response status code:", response.status_code)
    
    if response.status_code == 200:
        data = response.json()['response']['data']
        print("Response JSON data:", json.dumps(data, indent=3))
    else:
        print(f"Error: Unable to fetch data (status code {response.status_code})")

# Example usage
state = 'IN'
fetch_and_print_data(state, 2015, 'Wind')