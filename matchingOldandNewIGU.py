import pandas as pd

# Load the Excel files into DataFrames
geospatial_df = pd.read_excel('/Users/awfasano/Downloads/geospatialData.xlsx')
legend_df = pd.read_excel('/Users/awfasano/Downloads/legend.xlsx')

# Perform the merge operation
merged_df = geospatial_df.merge(legend_df[['Variable Name']], on='Variable Name', how='left', indicator=True)

# Add a column to indicate the match status
merged_df['Match Status'] = merged_df['_merge'].apply(lambda x: 'Matched' if x == 'both' else 'Not Matched')

# Drop the merge indicator column
merged_df.drop(columns=['_merge'], inplace=True)

# Save the updated DataFrame back to the Excel file
merged_df.to_excel('/Users/awfasano/Downloads/geospatial_matched.xlsx', index=False)

print("Matching process completed and saved to 'geospatial_matched.xlsx'")
