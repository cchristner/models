import pandas as pd
import os

def process_excel(file_path, rows_to_select):
    # Read the entire Excel sheet
    df = pd.read_excel(file_path, sheet_name="model", header=None)

    # Identify columns with 'q' and 'k' in row 1 (index 0)
    x_columns = df.iloc[0].isin(['q', 'k'])
    x_columns[0] = True  # Always include the first column (A)

    # Filter the DataFrame to keep only the columns with 'x' and column A
    df_filtered = df.loc[:, x_columns]

    # Select the desired rows
    selected_rows = df_filtered.iloc[rows_to_select]

    # Transpose the DataFrame
    df_transposed = selected_rows.T

    # Use row 3 (index 0 after transposition) as column names
    df_final = df_transposed.iloc[1:]  # Exclude the first row (now used as column names)
    df_final.columns = df_transposed.iloc[0]

    # Reset the index
    df_final.reset_index(drop=True, inplace=True)

    return df_final

# Set up the base path
base_path = r"C:\Users\corey\OneDrive\models"

# Process DE data
de_file_path = os.path.join(base_path, "DE", "DE.xlsx")
de_rows = [2, 10, 31, 36, 37]  # Rows 3, 11, 32, 37, and 38
de_df = process_excel(de_file_path, de_rows)

# Process CNH data
cnh_file_path = os.path.join(base_path, "CNH", "CNH.xlsx")
cnh_rows = [1, 2, 24, 28, 30]  # Rows 2, 3, 25, 29, and 31
cnh_df = process_excel(cnh_file_path, cnh_rows)

# Process AGCO data
agco_file_path = os.path.join(base_path, "AGCO", "AGCO.xlsx")
agco_rows = [1, 31, 33, 35]  # Rows 2, 32, 34, 36
agco_df = process_excel(agco_file_path, agco_rows)

# Process TITN data
titn_file_path = os.path.join(base_path, "TITN", "TITN.xlsx")
titn_rows = [1, 13, 18, 23, 23, 43, 53]  # Rows 2, 14, 19, 24, 44, 54
titn_df = process_excel(titn_file_path, titn_rows)

# Print the results
print("DE Data:")
print(de_df)
print("\nCNH Data:")
print(cnh_df)
print("\nAGCO Data:")
print(agco_df)
print("\nTITN Data:")
print(titn_df)