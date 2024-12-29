import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np

# Set the working directory to the script location
script_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(script_dir)
data_path = os.path.join(parent_dir, 'scratch_work', 'machinerydata.xlsx')

# Read the Excel file without using headers
df = pd.read_excel(data_path, header=None)

# Extract data for each company
de_data = df.iloc[2:9][[0, 4]].copy()
cnh_data = df.iloc[13:20][[0, 4]].copy()
agco_data = df.iloc[25:32][[0, 4]].copy()

# Convert R&D values to numeric type
de_data[4] = pd.to_numeric(de_data[4], errors='coerce')
cnh_data[4] = pd.to_numeric(cnh_data[4], errors='coerce')
agco_data[4] = pd.to_numeric(agco_data[4], errors='coerce')

# Take natural log of R&D values
de_log = np.log(de_data[4])
cnh_log = np.log(cnh_data[4])
agco_log = np.log(agco_data[4])

# Create the plot
plt.figure(figsize=(12, 6))

# Plot each company's log-transformed data with specified styles
plt.plot(de_data[0], de_log, color='green', linewidth=2, label='DE')
plt.plot(cnh_data[0], cnh_log, color='red', linewidth=2, label='CNH')
plt.plot(agco_data[0], agco_log, color='red', linestyle=':', linewidth=2, label='AGCO')

# Customize the plot
plt.title('Log-transformed R&D Spending Comparison', pad=20)
plt.xlabel('Quarter')
plt.ylabel('Log(R&D Spending)')
plt.grid(True, linestyle='--', alpha=0.7)
plt.legend()

# Rotate x-axis labels for better readability
plt.xticks(rotation=45)

# Adjust layout to prevent label cutoff
plt.tight_layout()

# Save the plot in the scratch_work folder
output_path = os.path.join(parent_dir, 'scratch_work', 'rnd_log_comparison.png')
plt.savefig(output_path)
plt.close()

# Print the data to verify values
print("\nDE Log-transformed R&D Values:")
print(pd.DataFrame({'Quarter': de_data[0], 'Log R&D': de_log}))
print("\nCNH Log-transformed R&D Values:")
print(pd.DataFrame({'Quarter': cnh_data[0], 'Log R&D': cnh_log}))
print("\nAGCO Log-transformed R&D Values:")
print(pd.DataFrame({'Quarter': agco_data[0], 'Log R&D': agco_log}))