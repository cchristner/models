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
de_data = df.iloc[2:9][[0, 2]].copy()
cnh_data = df.iloc[13:20][[0, 2]].copy()
agco_data = df.iloc[25:32][[0, 2]].copy()

# Convert sales values to numeric type
de_data[2] = pd.to_numeric(de_data[2], errors='coerce')
cnh_data[2] = pd.to_numeric(cnh_data[2], errors='coerce')
agco_data[2] = pd.to_numeric(agco_data[2], errors='coerce')

# Create the plot
plt.figure(figsize=(12, 6))

# Plot each company's sales data
plt.plot(de_data[0], de_data[2], marker='o', color='green', linewidth=2, label='DE')
plt.plot(cnh_data[0], cnh_data[2], marker='o', color='red', linewidth=2, label='CNH')
plt.plot(agco_data[0], agco_data[2], marker='o', color='blue', linewidth=2, label='AGCO')

# Customize the plot
plt.title('Quarterly Sales Data', pad=20)
plt.xlabel('Quarter')
plt.ylabel('Sales')
plt.grid(True, linestyle='--', alpha=0.7)
plt.legend()

# Rotate x-axis labels for better readability
plt.xticks(rotation=45)

# Adjust layout to prevent label cutoff
plt.tight_layout()

# Save the plot
output_path = os.path.join(parent_dir, 'scratch_work', 'sales_data.png')
plt.savefig(output_path)
plt.close()

# Print the data to verify values
print("\nDE Sales Data:")
print(pd.DataFrame({'Quarter': de_data[0], 'Sales': de_data[2]}))
print("\nCNH Sales Data:")
print(pd.DataFrame({'Quarter': cnh_data[0], 'Sales': cnh_data[2]}))
print("\nAGCO Sales Data:")
print(pd.DataFrame({'Quarter': agco_data[0], 'Sales': agco_data[2]}))
