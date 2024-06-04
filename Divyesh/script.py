import os
import pandas as pd

# Folder containing Excel files
folder_path = './data'

# List to store DataFrames
dataframes = []

# Loop through all Excel files in the folder
excel_files = os.listdir(folder_path)
for i, filename in enumerate(excel_files):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(folder_path, filename)

        # Read the Excel file
        df = pd.read_excel(file_path)  # Adjust sheet name if needed

        # Filter the DataFrame
        filtered_df = df[df['c name'].str.startswith('HEATWAVE', na=False)]

        # Append the filtered DataFrame to the list
        dataframes.append(filtered_df)

# Create a new Excel file
output_file_path = './tp5.xlsx'
with pd.ExcelWriter(output_file_path) as writer:
    # Loop through the filtered DataFrames and save each to a new sheet
    for i, df in enumerate(dataframes):
        df.to_excel(writer, sheet_name=f'sheet{i + 1}', index=False)

print(f"Filtered data saved to {output_file_path}")