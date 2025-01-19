import pandas as pd

# Path to the Fulltitles Excel file
fulltitles_file_path = '/Users/awfasano/Desktop/Nepal Scripts/inputs/Fulltitles.xlsx'

# List of paths to your indicator Excel files
indicator_file_paths = [
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/Accessibility Indicators.xlsx',
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/BUA Indicators.xlsx',
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/Night Time Lights.xlsx',
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/NPL Risk.xlsx'
]

# Load the Fulltitles Excel file into a pandas DataFrame
fulltitles_df = pd.read_excel(fulltitles_file_path)

# Initialize an empty DataFrame to hold the combined legend data
combined_legend_df = pd.DataFrame()

# Process each indicator file
for file_path in indicator_file_paths:
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(file_path)

    # Extract unique values for indicators and categories
    indicators = df['Variable Name'].unique()
    categories = df['Category'].unique()  # Ensure the column name matches exactly with the Excel file

    # Create dictionaries to map each unique indicator and category to a unique ID
    indicator_id_map = {indicator: idx + 1 for idx, indicator in enumerate(indicators)}
    category_id_map = {category: idx + 1 for idx, category in enumerate(categories)}

    # Map the IDs to the DataFrame
    df['indicator_id'] = df['Variable Name'].map(indicator_id_map)
    df['category_id'] = df['Category'].map(category_id_map)

    # Merge with the Fulltitles DataFrame on 'Variable Name'
    df = df.merge(fulltitles_df[['Variable Name', 'Fulltitle']], on='Variable Name', how='left')

    # Save the updated DataFrame back to the original Excel file
    df.to_excel(file_path, index=False)

    # Create the legend DataFrame with the specified columns
    legend_df = df[['Variable Name', 'indicator_id', 'Fulltitle', 'Category', 'category_id']].drop_duplicates()

    # Append to the combined legend DataFrame
    combined_legend_df = pd.concat([combined_legend_df, legend_df], ignore_index=True)

# Save the combined legend DataFrame to a new Excel file
combined_legend_file_path = '/Users/awfasano/Desktop/Nepal Scripts/inputs/combined_legend_with_fulltitles.xlsx'
combined_legend_df.to_excel(combined_legend_file_path, index=False)

# Save the matched variable names and full titles to a separate sheet in the same Excel file
matched_df = fulltitles_df[fulltitles_df['Variable Name'].isin(combined_legend_df['Variable Name'])]

with pd.ExcelWriter(combined_legend_file_path, mode='a', engine='openpyxl') as writer:
    matched_df.to_excel(writer, sheet_name='Matched_Fulltitles', index=False)

print(f'Combined legend with full titles and matched full titles have been saved to {combined_legend_file_path}')
