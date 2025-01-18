import pandas as pd
import numpy as np

tolerance = 1e-10  # Tolerance for near-zero comparisons

def clean_data(dfData):
    amount_column = 'Amount in local currency'  # Replace this with the exact column name

    clean_df = dfData.copy()

    # Sort the DataFrame by the 'Amount in local currency' column

    clean_df = clean_df.sort_values(by=amount_column)

    # Reset index for ease of row removal
    clean_df = clean_df.reset_index(drop=True)

    # Iterate through the sorted DataFrame to identify and remove rows that form pairs summing to near-zero
    indices_to_remove = set()
    i = 0
    j = len(clean_df) - 1

    while i < j:
        sum_val = clean_df.iloc[i][amount_column] + clean_df.iloc[j][amount_column]

        if np.isclose(sum_val, 0, atol=tolerance):
            indices_to_remove.add(i)
            indices_to_remove.add(j)
            i += 1
            j -= 1
        elif sum_val < 0:
            i += 1
        else:
            j -= 1

    # Remove rows forming pairs that sum to near-zero
    clean_df = clean_df.drop(list(indices_to_remove))

    return clean_df

# Read the Excel file
file_path = '/home/spectral/Documents/120 series period 4.xlsx'  # Replace with the path to your file
xl = pd.ExcelFile(file_path)

# Create a Pandas Excel writer
output_file = '/home/spectral/Documents/120 cleaned series period 4.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Loop through each sheet in the Excel file
    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)  # Read each sheet
        original_sum = df['Amount in local currency'].sum()  # Sum before cleaning

        cleaned_data = clean_data(df)  # Clean the data for the sheet

        cleaned_sum = cleaned_data['Amount in local currency'].sum()  # Sum after cleaning

        # Verify if the sum remains the same after cleaning
        if np.isclose(original_sum, cleaned_sum, atol=tolerance):
            print(f"Sum before cleaning in {sheet_name}: {original_sum}")
            print(f"Sum after cleaning in {sheet_name}: {cleaned_sum}")

            # Save the cleaned data to the Excel file
            cleaned_data.to_excel(writer, sheet_name=sheet_name, index=False)

        else:
            print(f"Error: Sum mismatch in sheet '{sheet_name}' after cleaning")
