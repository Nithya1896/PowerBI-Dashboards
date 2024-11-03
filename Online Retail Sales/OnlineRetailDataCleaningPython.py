# -*- coding: utf-8 -*-
"""
Created on Sun Oct 13 13:24:48 2024

@author: nithy
"""

import pandas as pd

# Load the data from both sheets
file_path = 'C:/Users/nithy/OneDrive/Documents/Power BI Practice/online_retail_II.xlsx'
sheets = pd.read_excel(file_path, sheet_name=['Year 2009-2010', 'Year 2010-2011']) 

# Define a function for data cleaning
def clean_data(df):
    # Remove rows with Quantity less than 1
    df = df[df['Quantity'] >= 1]

    # Remove rows with Unit Price less than or equal to 0
    df = df[df['Price'] > 0]

    # Remove rows with missing values in specified columns
    df = df.dropna(subset=['Invoice', 'StockCode', 'Description', 'Customer ID'])

    # Drop unnecessary columns
    df = df.drop(columns=['Invoice', 'StockCode', 'Description'])

    return df

# Apply the cleaning function to each sheet
cleaned_sheets = {sheet_name: clean_data(sheet_data) for sheet_name, sheet_data in sheets.items()}

# Save cleaned data back to a new Excel file with multiple sheets
with pd.ExcelWriter('C:/Users/nithy/OneDrive/Documents/Power BI Practice/cleaned_online_retail_data.xlsx', engine='openpyxl') as writer:
    for sheet_name, data in cleaned_sheets.items():
        data.to_excel(writer, sheet_name=sheet_name, index=False)
