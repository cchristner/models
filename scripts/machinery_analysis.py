import pandas as pd
import os

def process_excel(file_path, rows_to_select):
    # Read the entire Excel sheet
    df = pd.read_excel(file_path, sheet_name="model", header=None)

    # Identify columns with 'q' or 'k' in row 1 (index 0)
    q_k_columns = df.iloc[0].isin(['q'])
    q_k_columns[0] = True  # Always include the first column (A)

    # Filter the DataFrame to keep only the columns with 'q' or 'k' and column A
    df_filtered = df.loc[:, q_k_columns]

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

def analyze_column(df, column_label):
    column_data = pd.to_numeric(df[column_label], errors='coerce').dropna()
    baseline = column_data.iloc[0]
    highest_value = column_data.max()
    drop_back_value = column_data[column_data.idxmax():].min()
    percent_up = ((highest_value - baseline) / baseline) * 100
    percent_down = ((drop_back_value - highest_value) / highest_value) * 100

    return f"Baseline: {baseline:.2f}\n" \
           f"Highest Value: {highest_value:.2f}\n" \
           f"Dropped Back To: {drop_back_value:.2f}\n" \
           f"Percent Up: {percent_up:.2f}%\n" \
           f"Percent Down: {percent_down:.2f}%"

# Set up the base path
base_path = r"C:\Users\corey\OneDrive\models"

# Process DE data
de_file_path = os.path.join(base_path, "DE", "DE.xlsx")
de_rows = [2, 10, 31, 36, 37]  # Rows 3, 11, 32, 37, and 38
de_df = process_excel(de_file_path, de_rows)

# Process CNH data
cnh_file_path = os.path.join(base_path, "CNH", "CNH.xlsx")
cnh_rows = [2, 24, 28, 30]  # Rows 3, 25, 29, and 31
cnh_df = process_excel(cnh_file_path, cnh_rows)

# Process AGCO data
agco_file_path = os.path.join(base_path, "AGCO", "AGCO.xlsx")
agco_rows = [1, 31, 33, 35]  # Rows 2, 32, 34, 36
agco_df = process_excel(agco_file_path, agco_rows)

# Analyze R&D column for each DataFrame
print("DE Analysis:")
de_analysis = analyze_column(de_df, "RnD")
print(de_analysis)

print("\nCNH Analysis:")
cnh_analysis = analyze_column(cnh_df, "RnD")
print(cnh_analysis)

print("\nAGCO Analysis:")
agco_analysis = analyze_column(agco_df, "RnD")
print(agco_analysis)
