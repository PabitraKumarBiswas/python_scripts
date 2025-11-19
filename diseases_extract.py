# Import necessary libraries
import pandas as pd
import os

# Define file paths
excel_path = '/Users/pabitra/Documents/RIMES/Bamis Data Share.xlsx'
output_folder = '/Users/pabitra/Documents/RIMES/Python Codes /extracted info'
output_file = os.path.join(output_folder, 'Onion_Disease_Info.xlsx')

# Ensure the output folder exists
os.makedirs(output_folder, exist_ok=True)

# Read the specific sheet
sheet_name = 'Other Source - Diseases Informa'
df = pd.read_excel(excel_path, sheet_name=sheet_name)

# Display basic info
print("✅ Excel file loaded successfully!")
print(f"Total rows in sheet: {len(df)}")

# Check if column B exists (assuming host name column)
host_col = df.columns[1]  # Column B is index 1
print(f"Host column detected: {host_col}")

# Filter rows where host column contains 'Onion' or 'onion'
filtered_df = df[df[host_col].astype(str).str.contains('onion', case=False, na=False)]

# Display filtered results
print(f"✅ Total rows containing 'Onion': {len(filtered_df)}")
display(filtered_df.head())

# Save filtered data to Excel in the target folder
filtered_df.to_excel(output_file, index=False)

print(f"\n✅ Onion-related rows successfully saved to:\n{output_file}")
