import pandas as pd

# List of paths to your Excel and CSV files
file_paths = [
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/01_lgu_indicators_long (3).csv',
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/Accessibility Indicators.xlsx',
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/BUA Indicators.xlsx',
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/Night Time Lights.xlsx',
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/NPL Risk.xlsx',
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/muni_dashboard.xlsx'
]

# Initialize an empty DataFrame to hold the combined legend data
combined_legend_df = pd.DataFrame()

# Initialize the starting index for indicator_id and category_id
indicator_id_start = 1
category_id_start = 1

# Process each file
for file_path in file_paths:
    if file_path.endswith('.csv'):
        # Read the CSV file into a pandas DataFrame
        df = pd.read_csv(file_path)
        sheets = [df]  # Treat CSV as a single sheet
    else:
        # Read all sheets in the Excel file into a dictionary of DataFrames
        xls = pd.ExcelFile(file_path)
        sheets = [xls.parse(sheet_name) for sheet_name in xls.sheet_names]

    for df in sheets:
        # Check if the required columns exist
        required_columns = ['Variable Name', 'Category']
        if not all(column in df.columns for column in required_columns):
            print(f"Skipping sheet in file {file_path} due to missing columns.")
            continue

        # Fill missing Abstract column with empty strings if not present
        if 'Abstract' not in df.columns:
            df['Abstract'] = ''

        # Modify the Category column for muni_dashboard
        if 'muni_dashboard.xlsx' in file_path:
            df['Category'] = df['Category'] + ' - Census 2021'

        # Extract unique values for indicators and categories
        indicators = df['Variable Name'].unique()
        categories = df['Category'].unique()  # Ensure the column name matches exactly with the file

        # Create dictionaries to map each unique indicator and category to a unique ID
        indicator_id_map = {indicator: idx + indicator_id_start for idx, indicator in enumerate(indicators)}
        category_id_map = {category: idx + category_id_start for idx, category in enumerate(categories)}

        # Update the starting index for the next batch of indicator_ids and category_ids
        indicator_id_start += len(indicators)
        category_id_start += len(categories)

        # Map the IDs to the DataFrame
        df['indicator_id'] = df['Variable Name'].map(indicator_id_map)
        df['category_id'] = df['Category'].map(category_id_map)

        # Create the legend DataFrame with the specified columns
        legend_df = df[['Variable Name', 'indicator_id', 'Category', 'category_id', 'Abstract']].drop_duplicates()

        # Append to the combined legend DataFrame
        combined_legend_df = pd.concat([combined_legend_df, legend_df], ignore_index=True)

# Save the combined legend DataFrame to a new Excel file
combined_legend_df.to_excel('/Users/awfasano/Desktop/Nepal Scripts/inputs/legend.xlsx', index=False)

print('Combined legend has been saved to /Users/awfasano/Desktop/Nepal Scripts/inputs/legend.xlsx')
