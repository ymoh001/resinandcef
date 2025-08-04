'''after downloading the shipment file from Tableau, please manually save it as an excel file and then run this script
to remove manual preprocessing. I know sepatutnya buat straight from the csv but its more trouble than its worth so just buat mcm ni je.
Then can run Resin+Demand/CEF+Demand'''

import pandas as pd
import time  # Import the time module

# Start the timer
start_time = time.time()

# Load the Excel file, skipping the first two rows and setting the third row as headers
file_path = r'C:\Users\ymohdzaifullizan\OneDrive - Dyson\Year 2 rotation - E&O\Shipment details\Shipment Details 04 Aug 25 (2026).xlsx'
df = pd.read_excel(file_path, header=2)  # Set header to the third row directly

# Step 1: Print the headers to verify
print("Headers after processing:")
print(df.columns)

# Step 2: Ensure the headers for months 1 to 12 are integers or strings
df.columns = df.columns.astype(str)

# Step 3: Add a 'Total Demand' column summing columns 1 to 12 (representing months)
columns_to_sum = [str(i) for i in range(1, 13)]  # Convert the numbers 1-12 to strings to match months - 2025(1,13) - 2026(1,13)
df['Total Demand'] = df[columns_to_sum].sum(axis=1)

# Step 4: Replace values in the 'Vendor' column based on the provided table
replacement_dict = {
    "MEIBAN TECHNOLOGIES (MALAYSIA) SDN": "Meiban Technologies Sdn. Bhd",
    "SYARIKAT SIN KWANG PLASTIC": "SKP",
    "FLEXTRONICS MARKETING (L) LTD.": "Flex ZH",
    "V.S. INDUSTRY BERHAD": "VSI",
    "FLEXTRONICS INTERNATIONAL ASIA": "Flex",
    "KINPO ELECTRONICS": "Kinpo",
    "HI-P PHILIPPINES TECHNOLOGY CORP.": "Hi-P",
    "JABIL CIRCUIT SDN BHD": "Jabil MY",
    "Jabil Inc.": "Jabil MX",
    "Flextronics International Europe B.": "Flex",
    "FLEXTRONICS INTERNATIONAL ASIA PACI": "Flex",
    "Hi-P (Xiamen) Precision Plastic": "Hi-P Xiamen",
    "FLEXTRONICS SHAH ALAM SDN BHD": "Flex",
    "Wingtech Group (Hong Kong) Limited": "Luxshare", # Wingtech change to Luxshare
    "BCM Ltd": "BCM",
    "Dyson OPL - Downtons": "OPL",
    "PHICOUSTIC SYSTEMS (HK) LTD": "Phicoustic",
    "VS INDUSTRY PHILIPPINES  INC.": "VSI PH",
    "Luxshare Precision Limited": "Luxshare"
}

df['Vendor'] = df['Vendor'].replace(replacement_dict)

# Step 5: Save the processed data to a new sheet named 'preprocess'
with pd.ExcelWriter(file_path, mode='a', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='preprocess', index=False)
    
    # End the timer and calculate the duration
end_time = time.time()
execution_time = end_time - start_time

# Use divmod to convert seconds into minutes and seconds
minutes, seconds = divmod(execution_time, 60)

print("Processing and replacement completed!")
print(f"Time taken: {int(minutes)} minutes and {seconds:.2f} seconds")