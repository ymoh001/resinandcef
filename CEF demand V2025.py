# -*- coding: utf-8 -*-
"""
Created on Thu Jan  2 10:47:14 2025

@author: ymohdzaifullizan
"""

import pandas as pd
import time

# Start the timer
start_time = time.time()

# Reference path for the CEF List file
cef_file_path = r'C:\Users\ymohdzaifullizan\OneDrive - Dyson\Year 2 rotation - E&O\Paint\8. August 25\CEF Aug 2025.xlsx'

# Years to process
years = [2023, 2024, 2025, 2026]

# Create a dictionary to store CEF DataFrames for each year
cef_dfs = {}

# Iterate over each year to read the demand data and process CEF
for year in years:
    # Read the CEF Data
    CEF = pd.read_excel(cef_file_path)
    
    # Read the demand data for the current year
    demand_file_path = f'C:\\Users\\ymohdzaifullizan\\OneDrive - Dyson\\Year 2 rotation - E&O\\Shipment details\\Shipment Details 29 Sept 25 ({year}).xlsx'
    demand = pd.read_excel(demand_file_path, sheet_name='preprocess')

    # Initialize columns for Total Demand and months in CEF data with zeros
    CEF['Total Demand'] = 0
    for month in range(1, 13):
        CEF[str(month)] = 0

    # Iterate over each row in CEF data and sum corresponding values from demand_data
    for index, row in CEF.iterrows():
        part_number = row['Part_No']
        cm = row['CM']

        # Filter demand_data based on Partid and Vendor
        filtered_demand = demand[(demand['Partid'] == part_number) & (demand['Vendor'] == cm)]

        # Initialize a list of months
        months = [str(i) for i in range(1, 13)]

        # Sum Total Demand and each month's demand (handling missing columns gracefully)
        total_demand = filtered_demand['Total Demand'].sum() if 'Total Demand' in filtered_demand.columns else 0

        monthly_demands = {month: filtered_demand[month].sum() if month in filtered_demand.columns else 0 for month in months}

        # Update CEF_Jan with the summed values
        CEF.at[index, 'Total Demand'] = total_demand
        for month in months:
            CEF.at[index, month] = monthly_demands[month]

    # Store the CEF_Jan DataFrame for the current year in the dictionary
    cef_dfs[year] = CEF
    print(f"Successfully processed data for {year}.")

# Save all DataFrames to one Excel file with separate sheets
output_file_path = r'C:\Users\ymohdzaifullizan\OneDrive - Dyson\Year 2 rotation - E&O\Paint\9. September 25\PaintSept2023-2026 W40.xlsx'
with pd.ExcelWriter(output_file_path) as writer:
    for year, df in cef_dfs.items():
        df.to_excel(writer, sheet_name=str(year), index=False)
        print(f"Saved CEF data for {year} to {output_file_path}.")

# End the timer and calculate the duration
end_time = time.time()
execution_time = end_time - start_time

# Use divmod to convert seconds into minutes and seconds
minutes, seconds = divmod(execution_time, 60)

print("All data successfully saved in a single Excel file.")
print(f"Time taken: {int(minutes)} minutes and {seconds:.2f} seconds")