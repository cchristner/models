import pandas as pd
import os
from typing import List, Dict

# Configuration dictionary for company data
COMPANY_CONFIG = {
    'DE': {
        'folder': 'DE',
        'file': 'DE.xlsx',
        'rows': [2, 10, 31, 36, 37]  # Rows 3, 11, 32, 37, and 38
    },
    'CNH': {
        'folder': 'CNH',
        'file': 'CNH.xlsx',
        'rows': [1, 2, 24, 28, 30]  # Rows 2, 3, 25, 29, and 31
    },
    'AGCO': {
        'folder': 'AGCO',
        'file': 'AGCO.xlsx',
        'rows': [1, 30, 32, 34]  # Rows 2, 31, 35, 37
    },
    'TITN': {
        'folder': 'TITN',
        'file': 'TITN.xlsx',
        'rows': [1, 13, 18, 23, 23, 43, 53]  # Rows 2, 14, 19, 24, 44, 54
    }
}

def process_excel(file_path: str, rows_to_select: List[int]) -> pd.DataFrame:
    """
    Process an Excel file by selecting specific rows and columns.
    
    Args:
        file_path (str): Path to the Excel file
        rows_to_select (List[int]): List of row indices to select
        
    Returns:
        pd.DataFrame: Processed DataFrame
    """
    try:
        # Read the Excel sheet
        df = pd.read_excel(file_path, sheet_name="model", header=None)
        
        # Identify columns with 'q' and 'k' in row 1 (index 0)
        x_columns = df.iloc[0].isin(['q', 'k'])
        x_columns[0] = True  # Always include the first column (A)
        
        # Filter and process the DataFrame
        df_filtered = df.loc[:, x_columns]
        selected_rows = df_filtered.iloc[rows_to_select]
        df_transposed = selected_rows.T
        
        # Format the final DataFrame
        df_final = df_transposed.iloc[1:]
        df_final.columns = df_transposed.iloc[0]
        df_final.reset_index(drop=True, inplace=True)
        
        return df_final
    
    except Exception as e:
        print(f"Error processing {file_path}: {str(e)}")
        raise

def process_all_companies(base_path: str) -> Dict[str, pd.DataFrame]:
    """
    Process Excel files for all companies.
    
    Args:
        base_path (str): Base path for all company folders
        
    Returns:
        Dict[str, pd.DataFrame]: Dictionary of processed DataFrames for each company
    """
    company_data = {}
    
    for company, config in COMPANY_CONFIG.items():
        try:
            file_path = os.path.join(base_path, config['folder'], config['file'])
            company_data[company] = process_excel(file_path, config['rows'])
        except Exception as e:
            print(f"Error processing {company} data: {str(e)}")
            continue
    
    return company_data

def write_to_excel(company_data: Dict[str, pd.DataFrame], output_path: str, macro_data_path: str) -> None:
    """
    Write processed data to Excel file with company names as headers.
    
    Args:
        company_data (Dict[str, pd.DataFrame]): Dictionary of processed DataFrames
        output_path (str): Path for output Excel file
        macro_data_path (str): Path to the macro data Excel file
    """
    try:
        with pd.ExcelWriter(output_path) as writer:
            current_row = 0
            
            for company, df in company_data.items():
                # Write company name
                pd.DataFrame([company]).to_excel(
                    writer,
                    sheet_name='Sheet1',
                    startrow=current_row,
                    index=False,
                    header=False
                )
                
                # Write company data
                df.to_excel(
                    writer,
                    sheet_name='Sheet1',
                    startrow=current_row + 1,
                    index=False
                )
                
                # Update row counter for next company
                current_row += len(df) + 3  # +3 for spacing and company name
                
            # Read and write the macro data
            if os.path.exists(macro_data_path):
                macro_data_df = pd.read_excel(macro_data_path)
                # Write macro data
                pd.DataFrame(["Macro Data"]).to_excel(
                    writer,
                    sheet_name='Sheet1',
                    startrow=current_row,
                    index=False,
                    header=False
                )
                macro_data_df.to_excel(
                    writer,
                    sheet_name='Sheet1',
                    startrow=current_row + 1,
                    index=False
                )
                print(f"Macro data successfully added from {macro_data_path}")
                
        print(f"Data successfully saved to {output_path}")
        
    except Exception as e:
        print(f"Error writing to Excel: {str(e)}")
        raise

def main():
    """Main function to run the Excel processing pipeline."""
    try:
        # Set up the base path
        base_path = r"C:\Users\corey\OneDrive\models"
        output_file = os.path.join(base_path, "scratch_work", "machinerydata.xlsx")
        macro_data_file = os.path.join(base_path, "macro_data.xlsx")
        
        # Process all company data
        company_data = process_all_companies(base_path)
        
        # Print results
        for company, df in company_data.items():
            print(f"\n{company} Data:")
            print(df)
        
        # Write results to Excel
        write_to_excel(company_data, output_file, macro_data_file)
        
    except Exception as e:
        print(f"Error in main execution: {str(e)}")
        raise

if __name__ == "__main__":
    main()
