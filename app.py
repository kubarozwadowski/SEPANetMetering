import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import numpy as np
import plotly.express as px
import folium
from streamlit_folium import folium_static
import geopandas as gpd
from openpyxl import load_workbook
from monthly import calculate_stats_month, calculate_stats_uptomonth, calculate_stats_uptomonth_state

wb = load_workbook('net_metering2024.xlsx')
monthly_totals_states = wb['Monthly_Totals-States']
monthly_totals_us = wb['Monthly_Totals-US']

st.set_page_config(page_title="EIA Form Analysis", page_icon=":bar_chart:")


categories_mapping = {
    "Photovoltaic": {
        "Capacity MW": "E",
        "Customers": "J",
        "Virtual Capacity MW": "O",
        "Virtual Customers": "T",
        "Energy Sold Back MWh": "Y"
    },
    "Wind": {
        "Capacity MW": "BH",
        "Customers": "BM",
        "Energy Sold Back MWh": "BR"
    },
    "Battery": {
        "PV-Paired Battery Capacity (MW)": "AD",
        "PV-Paired Installations": "AI",
        "PV-Paired Energy Capacity (MWh)": "AN",
        "Not PV-Paired Battery Capacity (MW)": "AS",
        "PV-Paired Installations": "AX",
        "PV-Paired Energy Capacity (MWh)": "BC"
    },
    "Other": {
        "Capacity MW": "BW",
        "Customers": "CB",
        "Energy Sold Back MWh": "CG"
    },
    "All Technologies": {
        "Capacity MW": "CL",
        "Customers": "CQ",
        "Energy Sold Back MWh": "CV"
    }
}

month_mapping = {
    "January" : 1,
    "February" : 2,
    "March" : 3,
    "April" : 4
}

function_mapping = {
    "Calculate statistics for a specific month" : calculate_stats_month,
    "Calculate statistics up to and including a specific month": calculate_stats_uptomonth,
    "Calculate statistics by state up to and including a specific month": calculate_stats_uptomonth_state
}

def create_blank_us_map(title):
    # Create a folium map centered on the US
    m = folium.Map(location=[37.8, -96], zoom_start=4)
    folium_static(m)

st.sidebar.title("User Input")
with st.sidebar:


    function = st.selectbox("Select a Function", ["Calculate statistics for a specific month", "Calculate statistics up to and including a specific month", "Calculate statistics by state up to and including a specific month"])
    'You Selected: ', function


    category = st.selectbox('Select a Category', ["Photovoltaic", "Battery", "Wind", "Other", "All Technologies"])
    'You Selected: ', category

    subcategory = None
    if category:
        subcategory = st.selectbox('Select a Subcategory', list(categories_mapping[category].keys()))
        st.write('You Selected:', subcategory)


    month = st.selectbox('Select a Month', ["January", "February", "March", "April"])
    'You Selected: ', month

    if function == "Calculate statistics by state up to and including a specific month":
        state = st.selectbox('Select a State', [
            "AK", "AL", "AR", "AZ", "CA", "CO", "CT", "DE", "FL", "GA",
            "HI", "IA", "ID", "IL", "IN", "KS", "KY", "LA", "MA", "MD",
            "ME", "MI", "MN", "MO", "MS", "MT", "NC", "ND", "NE", "NH",
            "NJ", "NM", "NV", "NY", "OH", "OK", "OR", "PA", "RI", "SC",
            "SD", "TN", "TX", "UT", "VA", "VT", "WA", "WI", "WV", "WY"])
        'You Selected: ', state


    #show stats + maps + graphs on button press "Calculate"
    calculate = st.button('Calculate Stats')

infolink = "https://www.eia.gov/electricity/data/eia861m/#:~:text=Description%3A%20The,established%20imputation%20procedures."
st.logo("SEPALogo.png")
st.title("EIA-861M Net Metering Data")
st.subheader("Interactive Analysis of Monthly Net Metering Data Across Various Technologies and States")
st.subheader("[More Information](%s)" % infolink)
st.write("Use the sidebar on the left to adjust preferences (Drag to resize).")

if calculate:
        
        month_num = month_mapping[month]
        
        selected_function = function_mapping[function]
        
        start_column = categories_mapping[category][subcategory]
        
        
        if function == "Calculate statistics by state up to and including a specific month":
            stats = selected_function(monthly_totals_states, month_num, state, start_column)
        else:
            stats = selected_function(monthly_totals_states, month_num, start_column)

        if stats:
            st.write("## Statistics")
            data = {
                'Category': [],
                'Max Value': [],
                'Min Value': [],
                'Avg Value': []
            }

            for category, values in stats.items():
                data['Category'].append(category)
                data['Max Value'].append(values['max']['value'])
                data['Min Value'].append(values['min']['value'])
                data['Avg Value'].append(values['avg'])

       

            df = pd.DataFrame(data)
            category_order = ['Residential', 'Commercial', 'Industrial', 'Transportation', 'Total']
            df['Category'] = pd.Categorical(df['Category'], categories=category_order, ordered=True)
            st.dataframe(df)
            df = df.sort_values('Category')
            st.scatter_chart(df.set_index('Category'), color=['#FFA500', '#FF0000', '#0000FF'])

            if function == "Calculate statistics up to and including a specific month":
                all_months_data = []
                for m in range(1, month_num + 1):
                    monthly_stats = selected_function(monthly_totals_states, m, start_column)  # Problem fixed
                    if monthly_stats:
                        for category, values in monthly_stats.items():
                            all_months_data.append({
                                'Month': m,
                                'Category': category,
                                'Value': values['current']  # Use the current value
                            })
                df_all_months = pd.DataFrame(all_months_data)
                fig = px.line(df_all_months, x='Month', y='Value', color='Category', markers=True)
                st.plotly_chart(fig)

            if function == "Calculate statistics by state up to and including a specific month":
                all_months_data = []
                for m in range(1, month_num + 1):
                    monthly_stats = selected_function(monthly_totals_states, m, state, start_column)  # Problem fixed
                    if monthly_stats:
                        for category, values in monthly_stats.items():
                            all_months_data.append({
                                'Month': m,
                                'Category': category,
                                'Value': values['current']  # Use the current value
                            })
                df_all_months = pd.DataFrame(all_months_data)
                fig = px.line(df_all_months, x='Month', y='Value', color='Category', markers=True)
                st.plotly_chart(fig)

        if function == "Calculate statistics for a specific month" or function == "Calculate statistics up to and including a specific month":
            for cat in ['Residential', 'Commercial', 'Industrial', 'Transportation', 'Total']:
                st.write(f"### {cat} Average Values")
                create_blank_us_map(f"{cat} - Average Value per State")



