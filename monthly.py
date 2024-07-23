import openpyxl
from openpyxl import Workbook, load_workbook
import matplotlib.pyplot as plt
from netMetering import col_to_index, is_valid_integer


wb = load_workbook('net_metering2024.xlsx')
monthly_totals_states = wb['Monthly_Totals-States']
monthly_totals_us = wb['Monthly_Totals-US']

# UhKbugjHqnLMfOetbnKlXUd1wVI42kvFjWV8Bdy6

def calculate_stats_month(sheet, month, start_column):
    rows = list(sheet.iter_rows(values_only=True))
    data = rows[4:]  # Data starts on row 5

    # Filter by month
    filtered_data = [row for row in data if row[1] == month]

    if not filtered_data:
        print(f"No data found for month {month}")
        return None

    # Define the categories in the specified order
    categories = ['Residential', 'Commercial', 'Industrial', 'Transportation', 'Total']
    start_idx = col_to_index(start_column) - 1  # Convert column letter to index (0-based)
    
    stats = {}

    for idx, category in enumerate(categories):
        column_index = start_idx + idx
        valid_entries = [(row[2], row[column_index]) for row in filtered_data if row[column_index] is not None]
        valid_entries = [(state, value) for state, value in valid_entries if is_valid_integer(value)]
        

        if valid_entries:
            max_value = max(valid_entries, key=lambda x: x[1])
            min_value = min(valid_entries, key=lambda x: x[1])
            avg_value = sum(int(value) for state, value in valid_entries) / len(valid_entries)

            stats[category] = {
                'max': {'value': max_value[1], 'state': max_value[0]},
                'min': {'value': min_value[1], 'state': min_value[0]},
                'avg': avg_value
            }
        else:
            stats[category] = {
                'max': {'value': None, 'state': None},
                'min': {'value': None, 'state': None},
                'avg': None
            }

    # Print the stats for each category
    for category in categories:
        print(f"Category: {category}")
        print(f"  Max: {stats[category]['max']['value']} (State: {stats[category]['max']['state']})")
        print(f"  Min: {stats[category]['min']['value']} (State: {stats[category]['min']['state']})")
        print(f"  Avg: {stats[category]['avg']}")

    return stats

def calculate_stats_uptomonth_state(sheet, month, state, start_column):
    rows = list(sheet.iter_rows(values_only=True))
    data = rows[4:]  # Data starts on row 5

    # Filter by month and state, ensuring month is not None
    filtered_data = [row for row in data if row[1] is not None and row[1] <= month and row[2] == state]

    if not filtered_data:
        print(f"No data found for month {month} and state {state}")
        return None

    # Define the categories in the specified order
    categories = ['Residential', 'Commercial', 'Industrial', 'Transportation', 'Total']
    start_idx = col_to_index(start_column) - 1  # Convert column letter to index (0-based)
    
    stats = {}

    for idx, category in enumerate(categories):
        column_index = start_idx + idx
        valid_entries = [(row[1], row[2], row[column_index]) for row in filtered_data if row[column_index] is not None]
        valid_entries = [(month, state, value) for month, state, value in valid_entries if is_valid_integer(value)]
        
        if valid_entries:
            max_value = max(valid_entries, key=lambda x: x[2])
            min_value = min(valid_entries, key=lambda x: x[2])
            avg_value = sum(int(value) for _, _, value in valid_entries) / len(valid_entries)
            current_value = valid_entries[-1][2]  # Get the current month's value

            stats[category] = {
                'max': {'value': max_value[2], 'state': max_value[1], 'month': max_value[0]},
                'min': {'value': min_value[2], 'state': min_value[1], 'month': min_value[0]},
                'avg': avg_value,
                'current': current_value
            }
        else:
            stats[category] = {
                'max': {'value': None, 'state': None, 'month': None},
                'min': {'value': None, 'state': None, 'month': None},
                'avg': None,
                'current': None
            }

    return stats

def calculate_stats_uptomonth(sheet, month, start_column):
    rows = list(sheet.iter_rows(values_only=True))
    data = rows[4:]  # Data starts on row 5

    # Filter by month, ensuring month is not None
    filtered_data = [row for row in data if row[1] is not None and row[1] <= month]

    if not filtered_data:
        print(f"No data found for month {month}")
        return None

    categories = ['Residential', 'Commercial', 'Industrial', 'Transportation', 'Total']
    start_idx = col_to_index(start_column) - 1  # Convert column letter to index (0-based)
    
    stats = {}

    for idx, category in enumerate(categories):
        column_index = start_idx + idx
        valid_entries = [(row[1], row[2], row[column_index]) for row in filtered_data if row[column_index] is not None]
        valid_entries = [(month, state, value) for month, state, value in valid_entries if is_valid_integer(value)]
        
        if valid_entries:
            max_value = max(valid_entries, key=lambda x: x[2])
            min_value = min(valid_entries, key=lambda x: x[2])
            avg_value = sum(int(value) for _, _, value in valid_entries) / len(valid_entries)
            current_value = valid_entries[-1][2]  # Get the current month's value

            stats[category] = {
                'max': {'value': max_value[2], 'state': max_value[1], 'month': max_value[0]},
                'min': {'value': min_value[2], 'state': min_value[1], 'month': min_value[0]},
                'avg': avg_value,
                'current': current_value
            }
        else:
            stats[category] = {
                'max': {'value': None, 'state': None, 'month': None},
                'min': {'value': None, 'state': None, 'month': None},
                'avg': None,
                'current': None
            }

    return stats


print(calculate_stats_uptomonth(monthly_totals_states, 3, 'J'))
