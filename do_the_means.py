import pandas as pd
import os
import sys


def add_mean_sheet(file_path):
    # Read all sheets into a dictionary of DataFrames
    all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')

    # Identify numeric columns for mean calculation
    num_cols = pd.concat(all_sheets.values(), sort=False).select_dtypes(include='number').columns

    # Calculate mean for each variable across all sheets
    mean_df = pd.concat(all_sheets.values())[num_cols].groupby(level=0).mean()

    # Write the 'mean' sheet back to the original Excel file, overwriting the existing one
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        # Write 'mean' sheet
        mean_df.to_excel(writer, sheet_name='mean', index=False)


def add_transposed_mean_sheet(file_path):
    # Read the 'mean' sheet
    all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')

    # Check if 'mean' sheet exists
    if 'mean' in all_sheets:
        # Check if 'transposed_mean' sheet exists
        if 'transposed_mean' not in all_sheets:
            # Transpose 'mean' sheet
            transposed_mean_df = all_sheets['mean'].T.reset_index()

            # Rename the columns to include the variable names
            transposed_mean_df.columns = ['Variable'] + list(transposed_mean_df.columns[1:])

            # Write the 'transposed mean' sheet back to the original Excel file, overwriting the existing one
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
                # Write 'transposed_mean' sheet at the beginning (first row, first column)
                transposed_mean_df.to_excel(writer, sheet_name='transposed_mean', index=False, startrow=0, startcol=0)


def extract_and_store_variables(file_path, output_dict):
    # Read the 'transposed_mean' sheet
    transposed_mean_df = pd.read_excel(file_path, sheet_name='transposed_mean', engine='openpyxl')

    # Iterate through each variable in the 'transposed_mean' sheet
    for index, row in transposed_mean_df.iterrows():
        variable = str(row['Variable'])  # Convert to string

        if variable and variable not in output_dict:
            output_dict[variable] = pd.DataFrame(columns=['File'] + list(transposed_mean_df.columns[1:]))

        # Extract the variable data for the current row
        file_data = pd.DataFrame({
            'File': os.path.basename(file_path),
            **{col: [row[col]] for col in transposed_mean_df.columns[1:]}
        })

        # Concatenate the DataFrame to the output_dict[variable]
        if output_dict[variable].empty:
            output_dict[variable] = file_data
        else:
            output_dict[variable] = pd.concat([output_dict[variable], file_data], ignore_index=True)


# Specify the path of the target directory
target_directory = r'C:\Users\User\Desktop\CopyOfResults2023-extraction\BilateralSquat'

# Check if the directory exists
if not os.path.exists(target_directory):
    print(f"The specified directory '{target_directory}' does not exist.")
    sys.exit(1)

# Create a dictionary to store data for each variable
variable_data_dict = {}

# Iterate through Excel files in the current directory
for file_name in os.listdir(target_directory):
    if file_name.endswith('.xlsx') and 'squat' in file_name:
        file_path = os.path.join(target_directory, file_name)

        # Check if 'mean' and 'transposed_mean' sheets exist in the file
        with pd.ExcelFile(file_path) as xls:
            if 'mean' not in xls.sheet_names:
                add_mean_sheet(file_path)
            if 'transposed_mean' not in xls.sheet_names:
                add_transposed_mean_sheet(file_path)

        # Extract variables and store them in the dictionary
        extract_and_store_variables(file_path, variable_data_dict)


# Create a new Excel file for storing variable data
output_excel_path = os.path.join(target_directory, 'BilateralSquat/squat_variable_data.xlsx')
with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='w') as writer:
    # Write each variable's data to a separate sheet in the Excel file
    for variable, data in variable_data_dict.items():
        data.to_excel(writer, sheet_name=variable, index=False)
