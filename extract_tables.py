import pandas as pd
import os
import warnings
warnings.filterwarnings('ignore')

# Read the Excel file, specifically the "Serviços" sheet
df = pd.read_excel('Serviços maple bear zona norte.xls', sheet_name='Serviços', header=None)

# Rename columns
df.columns = ['Data', 'Nota', 'Cliente', 'Valor_Contabil'] if len(df.columns) >= 4 else df.columns

# Remove rows with empty dates or where 'Data' is in the Date column
df = df[df['Data'].notna()]
df = df[df['Data'].astype(str).str.lower() != 'data']

# Get unique client names
unique_clients = df['Cliente'].unique()

# Create output directory if it doesn't exist
output_dir = 'extracted_tables'
os.makedirs(output_dir, exist_ok=True)

# Extract tables for each client
for client in unique_clients:
    if pd.notna(client):  # Skip if client name is NaN
        # Get rows for current client
        client_df = df[df['Cliente'] == client]
        
        # If there's data for this client, process it
        if not client_df.empty:
            # Create safe filename by removing invalid characters
            safe_filename = "".join(x for x in str(client) if x.isalnum() or x in (' ', '-', '_'))
            
            # Count occurrences of each date
            date_counts = client_df['Data'].value_counts()
            has_duplicates = (date_counts > 1).any()
            
            if has_duplicates:
                duplicate_rows = client_df[client_df.duplicated(subset=['Data'], keep='first')]
                complementary_rows = client_df[~client_df.index.isin(duplicate_rows.index)]
                
                if not duplicate_rows.empty:
                    output_file = os.path.join(output_dir, f'{safe_filename}_filho1.xlsx')
                    duplicate_rows.to_excel(output_file, index=False)
                    print(f'Saved unique dates table for client: {client}')
                
                if not complementary_rows.empty:
                    output_file = os.path.join(output_dir, f'{safe_filename}_filho2.xlsx')
                    complementary_rows.to_excel(output_file, index=False)
                    print(f'Saved duplicate dates table for client: {client}')
            else:
                # Save single file if no duplicates
                output_file = os.path.join(output_dir, f'{safe_filename}.xlsx')
                client_df.to_excel(output_file, index=False)
                print(f'Saved table for client: {client}')