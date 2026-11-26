import pandas as pd
from tkinter import Tk, filedialog
from collections import Counter
from datetime import date
# Initialize Tkinter root
root = Tk()
root.withdraw()  # Hide the root window

# Prompt the user to select the Excel files
stock_file_path = filedialog.askopenfilename(title="Select Stock File", filetypes=[("Excel files", "*.xlsx *.xls")])
sale_file_path = filedialog.askopenfilename(title="Select Sold File", filetypes=[("Excel files", "*.xlsx *.xls")])

# Define the current date string for the file name
date_today = date.today().strftime('%d-%m')

# Check if a stock file is selected
if stock_file_path:
    df_stock = pd.read_excel(stock_file_path)
    # Stock file Replace the null rows value 
    df_stock['Tinch'] = df_stock['Tinch'].fillna('NONE')
    df_stock['Luster'] = df_stock['Luster'].fillna('NONE')
    df_stock['Luster'] = df_stock['Luster'].replace({'NORMAL': 'EX', 'NONE': 'EX'})
else:
    df_stock = None

# Check if a sale file is selected
if sale_file_path:
    df_Sale = pd.read_excel(sale_file_path)
    # Sale file Replace the null rows value 
    df_Sale['TINGE'] = df_Sale['TINGE'].fillna('NONE')
    df_Sale['LUSTER'] = df_Sale['LUSTER'].fillna('NONE')
    df_Sale['BLACK INCLUSION'] = df_Sale['BLACK INCLUSION'].fillna('NO BLK')
    df_Sale['LUSTER'] = df_Sale['LUSTER'].replace({'NORMAL': 'EX', 'NONE': 'EX'})
    df_Sale['CUT'] = df_Sale['CUT'].replace({'EX': 'EXCELLENT', 'VG': 'V. GOOD'})
    df_Sale['POLISH'] = df_Sale['POLISH'].replace({'EX': 'EXCELLENT', 'VG': 'V. GOOD'})
    df_Sale['SYMMETRY'] = df_Sale['SYMMETRY'].replace({'EX': 'EXCELLENT', 'VG': 'V. GOOD'})
    df_Sale['FLUOR'] = df_Sale['FLUOR'].replace({'FNT': 'FAINT'})
else:
    df_Sale = None

# Proceed only if at least one file is selected
if df_stock is None and df_Sale is None:
    print("No files selected. Please select at least one file to proceed.")
else:
    # Proceed with processing based on which files are available
    # Read the match file
    df_match = pd.read_excel(r"G:\Certificate\Python Output\Do Not Delete\Do Not Delete this file ( Buying File ) 1.00 UP.xlsx")
    
    # Ensure columns 'Stock', 'Sold', and 'FINAL CUSTOMER' are initialized correctly
    df_match['Stock'] = 0
    df_match['Sold'] = 0
    df_match['FINAL CUSTOMER'] = ''  # Explicitly set as string

    # Define the carat ranges
    carat_ranges = [(1.00, 1.09), (1.10, 1.49)]

    # Define method to assign carat ranges
    def assign_carat_range(carat, carat_ranges):
        for lower, upper in carat_ranges:
            if lower <= carat <= upper:
                return f'{lower:.2f}-{upper:.2f}'
        return None

    # Process stock file if available
    if df_stock is not None:
        df_stock['CARAT_RANGE'] = df_stock['Cts'].apply(lambda x: assign_carat_range(x, carat_ranges))
        grouped_df = df_stock.groupby(
            ['Shape', 'Cert', 'CARAT_RANGE', 'COLOR', 'Clarity', 'CUT', 'POLISH', 
             'Symetry', 'FLRC', 'Natts No', 'Tinch', 'Luster']
        ).size().reset_index(name='Stock')

        grouped_df.rename(columns={'CARAT_RANGE': 'SIZE'}, inplace=True)
        grouped_df['CUT'] = grouped_df.apply(
            lambda row: '3EX' if row['CUT'] == 'EXCELLENT' and row['POLISH'] == 'EXCELLENT' 
                             and row['Symetry'] == 'EXCELLENT' else row['CUT'], axis=1)

        # Apply filters to grouped_df (stock)
        filtered_grouped_df = grouped_df[
            (grouped_df['Shape'] == 'ROUND') &
            (grouped_df['Tinch'] == 'NONE') &
            (grouped_df['Luster'] == 'EX') & 
            (grouped_df['Cert'] == 'GIA') &
            grouped_df['FLRC'].isin(['NONE', 'FAINT']) &
            grouped_df['Symetry'].isin(['EXCELLENT', 'V. GOOD']) &
            grouped_df['POLISH'].isin(['EXCELLENT', 'V. GOOD']) &
            grouped_df['CUT'].isin(['EXCELLENT', '3EX', 'V. GOOD']) &
            grouped_df['COLOR'].isin(['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']) &
            grouped_df['Clarity'].isin(['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']) &
            grouped_df['Natts No'].isin(['NO BLK', 'CENTER - L', 'SIDE - L', 'S & C - L'])
        ]

        # Update Stock count in match file
        for i, match_row in df_match.iterrows():
            for j, grouped_row in filtered_grouped_df.iterrows():
                if (match_row['SIZE'] == grouped_row['SIZE'] and
                    match_row['COLOR'] == grouped_row['COLOR'] and
                    match_row['CLARITY'] == grouped_row['Clarity'] and
                    match_row['CUT'] == grouped_row['CUT'] and
                    match_row['POL'] == grouped_row['POLISH'] and
                    match_row['SYM'] == grouped_row['Symetry'] and
                    match_row['FLUOR'] == grouped_row['FLRC']):
                    df_match.at[i, 'Stock'] += grouped_row['Stock']
                    print(f"Matched Row in Grouped Data (index {j}):\n{grouped_row}")

    # Process sale file if available
    if df_Sale is not None:
        df_Sale['CARAT_RANGE'] = df_Sale['CARAT'].apply(lambda x: assign_carat_range(x, carat_ranges))
        grouped_df_Sale = df_Sale.groupby(
            ['SHAPE', 'CERT', 'CARAT_RANGE', 'COLOR', 'CLARITY', 'CUT', 'POLISH', 
             'SYMMETRY', 'FLUOR', 'BLACK INCLUSION', 'TINGE', 'LUSTER', 'FINAL CUSTOMER']
        ).size().reset_index(name='Sold')

        grouped_df_Sale.rename(columns={'CARAT_RANGE': 'SIZE'}, inplace=True)
        grouped_df_Sale['CUT'] = grouped_df_Sale.apply(
            lambda row: '3EX' if row['CUT'] == 'EXCELLENT' and row['POLISH'] == 'EXCELLENT' 
                             and row['SYMMETRY'] == 'EXCELLENT' else row['CUT'], axis=1)

        # Apply filters to grouped_df_Sale (sale)
        filtered_grouped_df_Sale = grouped_df_Sale[
            (grouped_df_Sale['SHAPE'] == 'ROUND') &
            (grouped_df_Sale['TINGE'] == 'NONE') &
            (grouped_df_Sale['LUSTER'] == 'EX') & 
            (grouped_df_Sale['CERT'] == 'GIA') &
            grouped_df_Sale['FLUOR'].isin(['NONE', 'FAINT']) &
            grouped_df_Sale['SYMMETRY'].isin(['EXCELLENT', 'V. GOOD']) &
            grouped_df_Sale['POLISH'].isin(['EXCELLENT', 'V. GOOD']) &
            grouped_df_Sale['CUT'].isin(['EXCELLENT', '3EX', 'V. GOOD']) &
            grouped_df_Sale['COLOR'].isin(['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']) &
            grouped_df_Sale['CLARITY'].isin(['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']) &
            grouped_df_Sale['BLACK INCLUSION'].isin(['NO BLK', 'CENTER - L', 'SIDE - L', 'S & C - L'])
        ]

        # Update Sale count and FINAL CUSTOMER in match file
        for i, match_row in df_match.iterrows():
            customer_counter = Counter()
            for j, grouped_row in filtered_grouped_df_Sale.iterrows():
                if (match_row['SIZE'] == grouped_row['SIZE'] and
                    match_row['COLOR'] == grouped_row['COLOR'] and
                    match_row['CLARITY'] == grouped_row['CLARITY'] and
                    match_row['CUT'] == grouped_row['CUT'] and
                    match_row['POL'] == grouped_row['POLISH'] and
                    match_row['SYM'] == grouped_row['SYMMETRY'] and
                    match_row['FLUOR'] == grouped_row['FLUOR']):
                    df_match.at[i, 'Sold'] += grouped_row['Sold']
                    customer_counter[grouped_row['FINAL CUSTOMER']] += grouped_row['Sold']
                    print(f"Matched Row in Grouped Data (index {j}):\n{grouped_row}")

            # Format FINAL CUSTOMER output
            customer_list = [f"{customer}- {count}" for customer, count in customer_counter.items()]
            df_match.at[i, 'FINAL CUSTOMER'] = ', '.join(customer_list)
            date_today = date.today().strftime('%d-%m')
    # Save the updated match file
    output_file_path = f'G:\\Certificate\\Python Output\\Buy 1 UP {date_today}.xlsx'
    df_match.to_excel(output_file_path, index=False)
