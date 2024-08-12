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
from monthly import calculate_stats_month, calculate_stats_uptomonth, calculate_stats_uptomonth_state, calculate_stats_range_months_state
from monthlyapi import technology_year, technology_range, state_technology_range, state_technology_year
from calendar import month_abbr
from datetime import datetime
import plotly.graph_objects as go

#wb = load_workbook('net_metering2024.xlsx')
#monthly_totals_states = wb['Monthly_Totals-States']
#monthly_totals_us = wb['Monthly_Totals-US']

wb = load_workbook('NetMeteringAll.xlsx')
monthly_totals_states = wb['Sheet1']

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
    "Jan": 1,
    "Feb": 2,
    "Mar": 3,
    "Apr": 4,
    "May": 5,
    "Jun": 6,
    "Jul": 7,
    "Aug": 8,
    "Sep": 9,
    "Oct": 10,
    "Nov": 11,
    "Dec": 12
}
#for displaying month name in tables
month_reverse_mapping = {v: k for k, v in month_mapping.items()}

function_mapping = {
    "Calculate statistics for a specific month (US Total)" : calculate_stats_month,
    "Calculate statistics up to and including a specific month (US Total)": calculate_stats_uptomonth,
    "Calculate statistics by state up to and including a specific month": calculate_stats_uptomonth_state,
    "Calculate statistics by state for a range of months": calculate_stats_range_months_state
}

def create_blank_us_map(title):
    # Create a folium map centered on the US
    m = folium.Map(location=[37.8, -96], zoom_start=4)
    folium_static(m)

st.sidebar.title("User Input")
with st.sidebar:

    filter_method = st.radio(
        "Filter Data By:",
        ["***Month***", "***Year***"],
        index=None
    )

    if filter_method == "***Month***":
        function = st.selectbox("Select a Function", ["Calculate statistics for a specific month (US Total)", "Calculate statistics by state for a range of months"], index=None)
        'You Selected: ', function

        category = st.selectbox('Select a Technology', ["Photovoltaic", "Battery", "Wind", "Other", "All Technologies"], index=None)
        'You Selected: ', category

        subcategory = None
        if category:
            subcategory = st.selectbox('Select a Sector', list(categories_mapping[category].keys()), index=None)
            st.write('You Selected:', subcategory)
    

        if function == "Calculate statistics for a specific month (US Total)":
            with st.expander('Please Select a Month and Year'):
                this_year = datetime.now().year
                this_month = datetime.now().month
                year = st.selectbox("", range(2024, 2011, -1))
                month_abbr = month_abbr[1:]
                report_month_str = st.radio("", month_abbr, index=this_month - 1, horizontal=True)
                month = month_abbr.index(report_month_str) + 1

            # Result
            st.text(f'You Selected: {report_month_str} {year} ')

        if function == "Calculate statistics by state for a range of months":
            with st.expander('Please Select a Start Month and Year'):
                this_year = datetime.now().year
                this_month = datetime.now().month
                start_year = st.selectbox("", range(2024, 2011, -1), key="start_year")
                start_month_abbr = list(month_abbr)[1:]
                start_report_month_str = st.radio("", start_month_abbr, index=this_month - 1, horizontal=True, key="start_month")
                start_month = start_month_abbr.index(start_report_month_str) + 1

            # Result
            st.text(f'Start Date: {start_report_month_str} {start_year} ')

            with st.expander('Please Select an End Month and Year'):
                this_year = datetime.now().year
                this_month = datetime.now().month
                end_year = st.selectbox("", range(2024, 2011, -1), key="end_year")
                end_month_abbr = list(month_abbr)[1:]
                end_report_month_str = st.radio("", end_month_abbr, index=this_month - 1, horizontal=True, key="end_month")
                end_month = end_month_abbr.index(end_report_month_str) + 1

            # Result
            st.text(f'End Date: {end_report_month_str} {end_year} ')

            
            state = st.selectbox('Select a State', [
                "AK", "AL", "AR", "AZ", "CA", "CO", "CT", "DE", "FL", "GA",
                "HI", "IA", "ID", "IL", "IN", "KS", "KY", "LA", "MA", "MD",
                "ME", "MI", "MN", "MO", "MS", "MT", "NC", "ND", "NE", "NH",
                "NJ", "NM", "NV", "NY", "OH", "OK", "OR", "PA", "RI", "SC",
                "SD", "TN", "TX", "UT", "VA", "VT", "WA", "WI", "WV", "WY"], index=None)
            'You Selected: ', state


    if filter_method == "***Year***": 
        function_api = st.selectbox("Select a Function", ["Calculate statistics for a specific year (US Total)", "Calculate statistics for a range of years (US Total)", "Calculate statistics by state for a specific year", "Calculate statistics by state for a range of years"], index=None)
        'You Selected: ', function_api

        category = st.selectbox('Select a Technology', ["Photovoltaic", "Wind", "Other", "All Technologies"], index=None)
        'You Selected: ', category

        if function_api == "Calculate statistics for a specific year (US Total)":
            year = st.slider("Select a Year", 2013, 2022, 2018)
            'You Selected: ', year

        if function_api == "Calculate statistics for a range of years (US Total)":
            values = st.slider("Select a range of values", 2013, 2022, (2016, 2019))
            st.write("Start Year:", values[0])
            st.write("End Year:", values[1])
            start_year = values[0]
            end_year = values[1]

        if function_api == "Calculate statistics by state for a specific year":
            year = st.slider("Select a Year", 2013, 2022, 2018)
            'You Selected: ', year

            state = st.selectbox('Select a State', [
                "AK", "AL", "AR", "AZ", "CA", "CO", "CT", "DE", "FL", "GA",
                "HI", "IA", "ID", "IL", "IN", "KS", "KY", "LA", "MA", "MD",
                "ME", "MI", "MN", "MO", "MS", "MT", "NC", "ND", "NE", "NH",
                "NJ", "NM", "NV", "NY", "OH", "OK", "OR", "PA", "RI", "SC",
                "SD", "TN", "TX", "UT", "VA", "VT", "WA", "WI", "WV", "WY"], index=None)
            'You Selected: ', state

        if function_api == "Calculate statistics by state for a range of years":
            values = st.slider("Select a range of values", 2013, 2022, (2016, 2019))
            st.write("Start Year:", values[0])
            st.write("End Year:", values[1])
            start_year = values[0]
            end_year = values[1]

            state = st.selectbox('Select a State', [
                "AK", "AL", "AR", "AZ", "CA", "CO", "CT", "DE", "FL", "GA",
                "HI", "IA", "ID", "IL", "IN", "KS", "KY", "LA", "MA", "MD",
                "ME", "MI", "MN", "MO", "MS", "MT", "NC", "ND", "NE", "NH",
                "NJ", "NM", "NV", "NY", "OH", "OK", "OR", "PA", "RI", "SC",
                "SD", "TN", "TX", "UT", "VA", "VT", "WA", "WI", "WV", "WY"], index=None)
            'You Selected: ', state

    #show stats + maps + graphs on button press "Calculate"
    calculate = st.button('Calculate Stats')


infolink = "https://www.eia.gov/electricity/data/eia861m/#:~:text=Description%3A%20The,established%20imputation%20procedures."
st.logo("SEPALogo.png")
st.title("EIA-861M Net Metering Data")
st.subheader("Interactive Analysis of Monthly Net Metering Data Across Various Technologies and States")
st.subheader("[More Information](%s)" % infolink)
st.write("Use the sidebar on the left to adjust preferences (Drag to resize).")
st.write("The data on this website is current as of ***May 2024***.")
st.subheader(":red[Important notes about the data:]")
with st.expander("Due to the varied formatting of the data, some values may be missing, especially for older months. Click to see details."):
    st.write("Because of the nature of how the data is collected, monthly data over a large range may be slow to display graphs")
    st.write("No virtual data between Jan 2011 and Dec 2016")
    st.write("No data for Tennessee in Jan 2011")
    st.write("No data for Alabama and Missippi between Jan 2011 and March 2012")
    st.write("No data about the battery technology between Jan 2011 and July 2023")



if calculate:
        
    if filter_method == "***Month***":
        if function == "Calculate statistics for a specific month (US Total)":
            month_num = month_mapping[report_month_str]
            start_column = categories_mapping[category][subcategory]
            stats, raw_data = calculate_stats_month(monthly_totals_states, year, month_num, start_column)
            
            if stats:
                st.write(f"## Statistics for {year} - {report_month_str}")

                # Prepare data for plotting
                plot_data = {
                    'State': ['US'],
                    'Residential': [],
                    'Commercial': [],
                    'Industrial': [],
                    'Transportation': [],
                    'Total': []
                }

                for category, values in stats.items():
                    plot_data[category].append(values['us_value'])

                df_plot = pd.DataFrame(plot_data)

                # Plot the data
                fig = go.Figure()
                for category in df_plot.columns[1:]:  # Skip the 'State' column
                    fig.add_trace(go.Bar(
                        x=df_plot['State'],
                        y=df_plot[category],
                        name=category
                    ))

                fig.update_layout(
                    title=f'Statistics for {year} - {report_month_str}',
                    xaxis_tickfont_size=14,
                    yaxis=dict(
                        title='Capacity (MW)',
                        titlefont_size=16,
                        tickfont_size=14,
                    ),
                    legend=dict(
                        x=0,
                        y=1.0,
                        bgcolor='rgba(255, 255, 255, 0)',
                        bordercolor='rgba(255, 255, 255, 0)'
                    ),
                    barmode='group',
                    bargap=0.15,
                    bargroupgap=0.1
                )

                st.plotly_chart(fig)

                # Display the data as a table
                st.write("## Tabular Data")
                st.table(df_plot)



        if function == "Calculate statistics by state for a range of months":
            start_column = categories_mapping[category][subcategory]

            # Call the calculate_stats_range_months_state function with the start and end year and month inputs
            stats, raw_data = calculate_stats_range_months_state(monthly_totals_states, start_year, end_year, start_month, end_month, state, start_column)

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
                    data['Max Value'].append(f"{values['max']['value']} ({values['max']['year']}-{month_reverse_mapping[values['max']['month']]})")
                    data['Min Value'].append(f"{values['min']['value']} ({values['min']['year']}-{month_reverse_mapping[values['min']['month']]})")
                    data['Avg Value'].append(values['avg'])

                df = pd.DataFrame(data)
                category_order = ['Residential', 'Commercial', 'Industrial', 'Transportation', 'Total']
                df['Category'] = pd.Categorical(df['Category'], categories=category_order, ordered=True)
                st.table(df)

                # Determine y-axis label based on subcategory
                y_axis_label = 'Value'
                if 'Capacity MW' in subcategory or 'Virtual Capacity MW' in subcategory:
                    y_axis_label = 'Capacity (MW)'
                elif 'Energy Sold Back MWh' in subcategory:
                    y_axis_label = 'Energy Sold Back (MWh)'

                # Create line graph for the range of months and years
                all_months_data = []
                for year in range(start_year, end_year + 1):
                    for month in range(1, 13):
                        if (year == start_year and month < start_month) or (year == end_year and month > end_month):
                            continue
                        monthly_stats, raw_month_data = calculate_stats_month(monthly_totals_states, year, month, start_column)
                        if monthly_stats:
                            for category, values in monthly_stats.items():
                                all_months_data.append({
                                    'Year': year,
                                    'Month': month,
                                    'Category': category,
                                    'Value': values['current']
                                })

                df_all_months = pd.DataFrame(all_months_data)
                df_all_months['Year-Month'] = df_all_months.apply(lambda row: f"{row['Year']}-{month_reverse_mapping[row['Month']]}", axis=1)
                fig = px.line(df_all_months, x='Year-Month', y='Value', color='Category', markers=True,
                            labels={'Value': y_axis_label})
                fig.update_xaxes(tickmode='array', tickvals=df_all_months['Year-Month'], ticktext=df_all_months['Year-Month'])
                st.plotly_chart(fig)

                # Create tabular data for the graph
                tabular_data = {
                    'Year-Month': df_all_months['Year-Month'].unique()
                }
                for category in category_order:
                    tabular_data[category] = df_all_months[df_all_months['Category'] == category]['Value'].values

                df_tabular = pd.DataFrame(tabular_data)
                st.write("## Tabular Data for Graph")
                st.table(df_tabular)



    if filter_method == "***Year***":
        if function_api == "Calculate statistics for a specific year (US Total)":
            result = technology_year(category, year)
            if result:
                for sector, data in result['sectors'].items():
                    sector_name = data['data'][0]['sectorName'].capitalize()
                    sector_title = sector_name if sector_name != "All sectors" else "All Sectors"
                    st.write(f"### {sector_title} ({year})")
                    
                    summary_df = pd.DataFrame({
                        'Metric': ['Max Capacity', 'Min Capacity', 'Average Capacity', 
                                'Max Customers', 'Min Customers', 'Average Customers'],
                        'Value': [f"{data['max_capacity']['value']} ({data['max_capacity']['state']})", 
                                f"{data['min_capacity']['value']} ({data['min_capacity']['state']})", 
                                data['average_capacity'], 
                                f"{data['max_customers']['value']} ({data['max_customers']['state']})", 
                                f"{data['min_customers']['value']} ({data['min_customers']['state']})", 
                                data['average_customers']]
                    })
                    
                    st.table(summary_df)
                    
                    # Create bar graphs for capacity and customers
                    detailed_df = pd.DataFrame(data['data'])
                    fig_capacity = px.bar(detailed_df, x='stateName', y='capacity', color='sector',
                                        title=f'Capacity by Sector for {year}', labels={'capacity': 'Capacity (MWh)'})
                    fig_customers = px.bar(detailed_df, x='stateName', y='customers', color='sector',
                                        title=f'Customers by Sector for {year}', labels={'customers': 'Number of Customers'})
                    
                    st.plotly_chart(fig_capacity)
                    st.plotly_chart(fig_customers)

        if function_api == "Calculate statistics for a range of years (US Total)":
            result = technology_range(category, start_year, end_year)
            if result:
                for sector, data in result['sectors'].items():
                    sector_name = data['data'][0]['sectorName'].capitalize()
                    sector_title = sector_name if sector_name != "All sectors" else "All Sectors"
                    st.write(f"### {sector_title} ({start_year} - {end_year})")
                    
                    summary_df = pd.DataFrame({
                        'Metric': ['Max Capacity', 'Min Capacity', 'Average Capacity', 
                                'Max Customers', 'Min Customers', 'Average Customers'],
                        'Value': [f"{data['max_capacity']['value']} ({data['max_capacity']['state']})", 
                                f"{data['min_capacity']['value']} ({data['min_capacity']['state']})", 
                                data['average_capacity'], 
                                f"{data['max_customers']['value']} ({data['max_customers']['state']})", 
                                f"{data['min_customers']['value']} ({data['min_customers']['state']})", 
                                data['average_customers']]
                    })
                    
                    st.table(summary_df)
                    
                    # Prepare data for trend graph
                    yearly_data = [item for item in data['data'] if item['state'] == 'US']
                    trend_df = pd.DataFrame(yearly_data)

                    # Ensure unique data points
                    trend_df = trend_df.drop_duplicates().dropna()

                    # Sort data by year
                    trend_df = trend_df.sort_values(by='period')

                    # Create line graphs for trends
                    fig_capacity_trend = px.line(trend_df, x='period', y='capacity', title='Total Capacity Trend',
                                                labels={'capacity': 'Capacity (MWh)', 'period': 'Year'})
                    fig_capacity_trend.update_xaxes(dtick=1)
                    fig_customers_trend = px.line(trend_df, x='period', y='customers', title='Total Customers Trend',
                                                labels={'customers': 'Number of Customers', 'period': 'Year'})
                    fig_customers_trend.update_xaxes(dtick=1)
                    
                    st.plotly_chart(fig_capacity_trend)
                    st.plotly_chart(fig_customers_trend)

        if function_api == "Calculate statistics by state for a specific year":
            result = state_technology_year(category, year, state)
            if result:
                for sector, data in result['sectors'].items():
                    sector_name = data['data'][0]['sectorName'].capitalize()
                    sector_title = sector_name if sector_name != "All sectors" else "All Sectors"
                    st.write(f"### {sector_title} ({state}, {year})")
                    
                    summary_df = pd.DataFrame({
                        'Metric': ['Max Capacity', 'Min Capacity', 'Average Capacity', 
                                'Max Customers', 'Min Customers', 'Average Customers'],
                        'Value': [f"{data['max_capacity']['value']} ({data['max_capacity']['state']})", 
                                f"{data['min_capacity']['value']} ({data['min_capacity']['state']})", 
                                data['average_capacity'], 
                                f"{data['max_customers']['value']} ({data['max_customers']['state']})", 
                                f"{data['min_customers']['value']} ({data['min_customers']['state']})", 
                                data['average_customers']]
                    })
                    
                    st.table(summary_df)
                    
                    # Create bar graphs for capacity and customers
                    detailed_df = pd.DataFrame(data['data'])
                    fig_capacity = px.bar(detailed_df, x='stateName', y='capacity', color='sector',
                                        title=f'Capacity by Sector for {state} ({year})', labels={'capacity': 'Capacity (MWh)'})
                    fig_customers = px.bar(detailed_df, x='stateName', y='customers', color='sector',
                                        title=f'Customers by Sector for {state} ({year})', labels={'customers': 'Number of Customers'})
                    
                    st.plotly_chart(fig_capacity)
                    st.plotly_chart(fig_customers)

        if function_api == "Calculate statistics by state for a range of years":
            result = state_technology_range(category, start_year, end_year, state)
            if result:
                for sector, data in result['sectors'].items():
                    sector_name = data['data'][0]['sectorName'].capitalize()
                    sector_title = sector_name if sector_name != "All sectors" else "All Sectors"
                    st.write(f"### {sector_title} ({state}, {start_year} - {end_year})")
                    
                    summary_df = pd.DataFrame({
                        'Metric': ['Max Capacity', 'Min Capacity', 'Average Capacity', 
                                'Max Customers', 'Min Customers', 'Average Customers'],
                        'Value': [f"{data['max_capacity']['value']} ({data['max_capacity']['state']})", 
                                f"{data['min_capacity']['value']} ({data['min_capacity']['state']})", 
                                data['average_capacity'], 
                                f"{data['max_customers']['value']} ({data['max_customers']['state']})", 
                                f"{data['min_customers']['value']} ({data['min_customers']['state']})", 
                                data['average_customers']]
                    })
                    
                    st.table(summary_df)
                    
                    # Prepare data for trend graph
                    yearly_data = [item for item in data['data'] if item['state'] == 'US']
                    trend_df = pd.DataFrame(yearly_data)

                    # Ensure unique data points
                    trend_df = trend_df.drop_duplicates().dropna()

                    # Sort data by year
                    trend_df = trend_df.sort_values(by='period')

                    # Create line graphs for trends
                    fig_capacity_trend = px.line(trend_df, x='period', y='capacity', title='Total Capacity Trend',
                                                labels={'capacity': 'Capacity (MWh)', 'period': 'Year'})
                    fig_capacity_trend.update_xaxes(dtick=1)
                    fig_customers_trend = px.line(trend_df, x='period', y='customers', title='Total Customers Trend',
                                                labels={'customers': 'Number of Customers', 'period': 'Year'})
                    fig_customers_trend.update_xaxes(dtick=1)
                    
                    st.plotly_chart(fig_capacity_trend)
                    st.plotly_chart(fig_customers_trend)