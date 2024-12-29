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

# Function to calculate quarter-over-quarter percentage change
def calculate_qoq_change(series):
    # Calculate percentage change between consecutive quarters
    pct_changes = series.pct_change() * 100
    # First value will be NaN, set it to 0 since it's the starting point
    pct_changes.iloc[0] = 0
    return pct_changes

# Calculate quarter-over-quarter changes
de_changes = calculate_qoq_change(de_data[4])
cnh_changes = calculate_qoq_change(cnh_data[4])
agco_changes = calculate_qoq_change(agco_data[4])

# Create the plot
plt.figure(figsize=(12, 6))

# Plot each company's data with specified styles
plt.plot(de_data[0], de_changes, color='green', linewidth=2, label='DE')
plt.plot(cnh_data[0], cnh_changes, color='red', linewidth=2, label='CNH')
plt.plot(agco_data[0], agco_changes, color='red', linestyle=':', linewidth=2, label='AGCO')

# Customize the plot
plt.title('Quarter-over-Quarter Change in R&D Spending', pad=20)
plt.xlabel('Quarter')
plt.ylabel('Percent Change from Previous Quarter (%)')
plt.grid(True, linestyle='--', alpha=0.7)
plt.legend()

# Rotate x-axis labels for better readability
plt.xticks(rotation=45)

# Adjust layout to prevent label cutoff
plt.tight_layout()

# Save the plot
output_path = os.path.join(parent_dir, 'scratch_work', 'rnd_qoq_changes.png')
plt.savefig(output_path)
plt.close()

# Print the data to verify values
print("\nDE Quarter-over-Quarter Changes:")
print(pd.DataFrame({'Quarter': de_data[0], 'Change': de_changes}))
print("\nCNH Quarter-over-Quarter Changes:")
print(pd.DataFrame({'Quarter': cnh_data[0], 'Change': cnh_changes}))
print("\nAGCO Quarter-over-Quarter Changes:")
print(pd.DataFrame({'Quarter': agco_data[0], 'Change': agco_changes}))