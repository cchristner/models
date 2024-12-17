import pandas as pd
import os
import openpyxl

# Define the base path
base_path = "C:/users/corey/onedrive/models"

# Define the path to the DE.xlsx file
de_file_path = os.path.join(base_path, "DE", "DE.xlsx")

# Read the "Scripting" sheet from the DE.xlsx file
df_de = pd.read_excel(de_file_path, sheet_name="Scripting", engine='openpyxl')

# Transpose the DataFrame
df_de_transposed = df_de.transpose()

# Print the transposed DataFrame to the console
print(df_de_transposed)
