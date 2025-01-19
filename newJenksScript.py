import pandas as pd
import jenkspy

# Function to calculate Jenks breaks for each variable
def calculate_jenks_breaks_for_all_variables(df, data_column, variable_column, num_classes):
    unique_variables = df[variable_column].unique()
    jenks_breaks = {}

    for variable in unique_variables:
        filtered_data = df[df[variable_column] == variable][data_column].dropna()
        if not filtered_data.empty and filtered_data.nunique() > 1:
            data = filtered_data.values
            try:
                breaks = jenkspy.jenks_breaks(data, n_classes=num_classes)
                if len(set(breaks)) > 1:
                    jenks_breaks[variable] = breaks
                    print(f"Jenks breaks for {variable}: Low={breaks[0]}, Medium={breaks[1]}, High={breaks[2]}")
                else:
                    jenks_breaks[variable] = None
            except Exception as e:
                print(f"Failed to calculate Jenks breaks for {variable}: {e}")
                jenks_breaks[variable] = None
        else:
            jenks_breaks[variable] = None
    return jenks_breaks

# Function to apply classifications based on Jenks breaks
def apply_classification(df, breaks_dict, data_column, variable_column):
    for variable, breaks in breaks_dict.items():
        if breaks:
            mask = df[variable_column] == variable
            try:
                df.loc[mask, 'class'] = pd.cut(
                    df.loc[mask, data_column],
                    bins=breaks,
                    labels=['Low', 'Medium', 'High'],
                    include_lowest=True)
            except Exception as e:
                print(f"Error applying classification for {variable}: {e}")

# Paths to the Excel and CSV files
file_paths = [
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/Accessibility Indicators.xlsx',
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/BUA Indicators.xlsx',
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/Night Time Lights.xlsx',
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/NPL Risk.xlsx',
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/muni_dashboard.xlsx',
    '/Users/awfasano/Desktop/Nepal Scripts/inputs/geospatialData.xlsx'
]

# Process each file
results = {}
unmatched_variables = set()

for file_path in file_paths:
    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
            sheets = [df]  # Treat CSV as a single sheet
        else:
            xls = pd.ExcelFile(file_path)
            sheets = [xls.parse(sheet_name) for sheet_name in xls.sheet_names]

        combined_df = pd.DataFrame()
        for sheet_index, sheet_df in enumerate(sheets):
            df = sheet_df if isinstance(sheet_df, pd.DataFrame) else sheet_df

            # Ensure the required columns are present
            required_columns = ['Variable Name', 'Value']
            missing_columns = [col for col in required_columns if col not in df.columns]

            if missing_columns:
                print(f"Skipping sheet in file {file_path} due to missing columns: {missing_columns}")
                print(f"Columns present in the dataframe: {df.columns.tolist()}")
                continue

            # Ensure Value column is numeric
            df['Value'] = pd.to_numeric(df['Value'], errors='coerce')
            df.dropna(subset=['Value'], inplace=True)

            num_classes = 3
            breaks_dict = calculate_jenks_breaks_for_all_variables(df, 'Value', 'Variable Name', num_classes)

            # Apply classification only for variables with non-null breaks
            filtered_breaks_dict = {k: v for k, v in breaks_dict.items() if v}
            apply_classification(df, filtered_breaks_dict, 'Value', 'Variable Name')

            # Add columns for Low, Medium, High breaks
            df['Low'] = df['Variable Name'].map(lambda x: filtered_breaks_dict[x][0] if x in filtered_breaks_dict else None)
            df['Medium'] = df['Variable Name'].map(lambda x: filtered_breaks_dict[x][1] if x in filtered_breaks_dict else None)
            df['High'] = df['Variable Name'].map(lambda x: filtered_breaks_dict[x][2] if x in filtered_breaks_dict else None)

            # Combine sheets from muni_dashboard.xlsx
            if file_path.endswith('muni_dashboard.xlsx'):
                combined_df = pd.concat([combined_df, df], ignore_index=True)
            else:
                # Store results for each individual file
                sheet_name = file_path.split('/')[-1].split('.')[0]
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]
                results[sheet_name] = df

            # Identify and print unmatched variables
            for variable in df['Variable Name'].unique():
                unmatched_variables.add(variable)

        if not combined_df.empty:
            results['muni_dashboard'] = combined_df

    except Exception as e:
        print(f"Error processing file {file_path}: {e}")

# Write the processed data to Excel
with pd.ExcelWriter('/Users/awfasano/Desktop/Nepal Scripts/inputs/geospatialData_with_jenks1.xlsx') as writer:
    for sheet_name, df in results.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# Print unmatched variables
if unmatched_variables:
    print("Unmatched variables:")
    for variable in unmatched_variables:
        print(variable)
else:
    print("All variables processed successfully.")
