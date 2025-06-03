# -*- coding: utf-8 -*-
"""
Created on Thu Jan  2 11:30:00 2025

@author: ymohdzaifullizan
"""

import pandas as pd
import time

# Start the timer
start_time = time.time()

# Reference paths for Resin list
resin_file_path = r"C:\Users\ymohdzaifullizan\OneDrive - Dyson\Year 2 rotation - E&O\Resin\Resin + Demand\5. Resin - Demand (May 2025)\Resin_May'25.xlsx"

# Years to process
years = [2023, 2024, 2025, 2026]

# Create a dictionary to store resin DataFrames for each year
resin_jan_dfs = {}

# Iterate over each year to read the demand data and process resin
for year in years:
    # Read the resin Data
    Resin_Jan = pd.read_excel(resin_file_path)
     
    # Read the demand data for the current year
    demand_file_path = f'C:\\Users\\ymohdzaifullizan\\OneDrive - Dyson\\Year 2 rotation - E&O\\Shipment details\\Shipment Details 03 June 25 ({year}).xlsx'
    demand = pd.read_excel(demand_file_path, sheet_name='preprocess')

    # Initialize columns for Total Demand and months in resin data with zeros
    Resin_Jan['Total Demand'] = 0
    for month in range(1, 13):
        Resin_Jan[str(month)] = 0

    # Iterate over each row in resin data and sum corresponding values from demand_data
    for index, row in Resin_Jan.iterrows():
        part_number = row['Part_Number_No_Rev']
        cm = row['CM']

        # Filter demand_data based on Partid and Vendor
        filtered_demand = demand[(demand['Partid'] == part_number) & (demand['Vendor'] == cm)]

        # Initialize a list of months
        months = [str(i) for i in range(1, 13)]

        # Sum Total Demand and each month's demand (handling missing columns gracefully)
        total_demand = filtered_demand['Total Demand'].sum() if 'Total Demand' in filtered_demand.columns else 0

        monthly_demands = {month: filtered_demand[month].sum() if month in filtered_demand.columns else 0 for month in months}

        # Update Resin with the summed values
        Resin_Jan.at[index, 'Total Demand'] = total_demand
        for month in months:
            Resin_Jan.at[index, month] = monthly_demands[month]

    # Store the CEF_Jan DataFrame for the current year in the dictionary
    resin_jan_dfs[year] = Resin_Jan
    print(f"Successfully processed data for {year}.")

# Save all DataFrames to one Excel file with separate sheets
output_file_path = r'C:\Users\ymohdzaifullizan\OneDrive - Dyson\Year 2 rotation - E&O\Resin\Resin + Demand\6. Resin - Demand (June 2025)\ResinJune2023-2026 W23.xlsx'
with pd.ExcelWriter(output_file_path) as writer:
    for year, df in resin_jan_dfs.items():
        df.to_excel(writer, sheet_name=str(year), index=False)
        print(f"Saved Resin data for {year} to {output_file_path}.")

# End the timer and calculate the duration
end_time = time.time()
execution_time = end_time - start_time

# Use divmod to convert seconds into minutes and seconds
minutes, seconds = divmod(execution_time, 60)

print("All data successfully saved in a single Excel file.")
print(f"Time taken: {int(minutes)} minutes and {seconds:.2f} seconds")