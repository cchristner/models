import pandas as pd
import matplotlib.pyplot as plt
import os

# Set the working directory to the script location
script_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(script_dir)
data_path = os.path.join(parent_dir, 'scratch_work', 'machinerydata.xlsx')

# Read the Excel file without using headers
df = pd.read_excel(data_path, header=None)

# Extract data for each company
# Using numeric indices: 0 for column A, 4 for column E (0-based indexing)
de_data = df.iloc[2:9][[0, 4]].copy()
cnh_data = df.iloc[13:20][[0, 4]].copy()
agco_data = df.iloc[25:32][[0, 4]].copy()

# Convert R&D values (column 4) to numeric type
de_data[4] = pd.to_numeric(de_data[4], errors='coerce')
cnh_data[4] = pd.to_numeric(cnh_data[4], errors='coerce')
agco_data[4] = pd.to_numeric(agco_data[4], errors='coerce')

# Function to calculate percentage change from baseline
def calculate_relative_change(series):
    baseline = series.iloc[0]
    return ((series - baseline) / baseline) * 100

# Calculate relative changes
de_changes = calculate_relative_change(de_data[4])
cnh_changes = calculate_relative_change(cnh_data[4])
agco_changes = calculate_relative_change(agco_data[4])

# Create the plot
plt.figure(figsize=(12, 6))

# Plot each company's data with specified styles
plt.plot(de_data[0], de_changes, color='green', linewidth=2, label='DE')
plt.plot(cnh_data[0], cnh_changes, color='red', linewidth=2, label='CNH')
plt.plot(agco_data[0], agco_changes, color='red', linestyle=':', linewidth=2, label='AGCO')

# Customize the plot
plt.title('Relative Change in R&D Spending from Baseline (Q1 2023)', pad=20)
plt.xlabel('Quarter')
plt.ylabel('Percent Change from Baseline (%)')
plt.grid(True, linestyle='--', alpha=0.7)
plt.legend()

# Rotate x-axis labels for better readability
plt.xticks(rotation=45)

# Adjust layout to prevent label cutoff
plt.tight_layout()

# Save the plot
output_path = os.path.join(script_dir, 'rnd_relative_changes.png')
plt.savefig(output_path)
plt.close()

# Print the data to verify values
print("\nDE Data:")
print(de_data)
print("\nCNH Data:")
print(cnh_data)
print("\nAGCO Data:")
print(agco_data)