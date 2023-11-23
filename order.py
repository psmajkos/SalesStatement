import csv
import pandas as pd
from tkinter import Tk, filedialog

# Create a Tkinter root window (it will be hidden)
root = Tk()
root.withdraw()

# Ask the user to select the input CSV file using a file dialog
csv_file_path = filedialog.askopenfilename(title="Select CSV file", filetypes=[("CSV files", "*.csv")])

# Ask the user to specify the output Excel file using a file dialog
output_excel_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

# Open the selected CSV file
with open(csv_file_path, newline='', encoding='utf-8') as csvfile:
    # Read the CSV file
    csv_reader = csv.DictReader(csvfile)

    type_column_name = 'Type'
    value_to_filter = 'order'

    # Define the conditions for selecting data
    selected_columns = ['OrderDate', 'BuyerName', 'BuyerPhone', 'BuyerAddress', 'BuyerZip', 'BuyerCity', 'BuyerCountryCode', 'DeliveryMethod', 'DeliveryAmount', 'TotalToPayAmount']

    # Initialize an empty list to store selected data
    selected_data_list = []

    for row in csv_reader:
        if row[type_column_name] == value_to_filter:
            selected_data = [row.get(column, '') for column in selected_columns]
            selected_data_list.append(selected_data)

# Create a DataFrame from the selected data
df = pd.DataFrame(selected_data_list, columns=selected_columns)

# Convert 'TotalToPayAmount' column to numeric values
df['TotalToPayAmount'] = pd.to_numeric(df['TotalToPayAmount'], errors='coerce')

# Convert 'OrderDate' to datetime format
df['OrderDate'] = pd.to_datetime(df['OrderDate']).dt.date

# Concatenate selected columns into a single column
df['CombinedAddress'] = df[['BuyerName', 'BuyerAddress', 'BuyerZip', 'BuyerCity', 'BuyerCountryCode', 'BuyerPhone']].agg(', '.join, axis=1)

# Drop the individual columns that were combined
df = df.drop(['BuyerName', 'BuyerAddress', 'BuyerZip', 'BuyerCity', 'BuyerCountryCode', 'BuyerPhone'], axis=1)

# Add a new column 'TotalPaidAmountDivided' by dividing 'TotalToPayAmount' by 1.23
df['TotalPaidAmountDivided'] = df['TotalToPayAmount'] / 1.23

# Add a new column 'Difference' by subtracting 'TotalPaidAmountDivided' from 'TotalToPayAmount'
df['Wartość VAT'] = df['TotalToPayAmount'] - df['TotalPaidAmountDivided']

# # Create a new column 'TotalAmountWithDelivery' by adding 'TotalToPayAmount' and 'DeliveryAmount' when 'DeliveryAmount' is equal to 8.99
# df['TotalAmountWithDelivery'] = df.apply(lambda row: row['TotalToPayAmount'] + row['DeliveryAmount'] if row['DeliveryAmount'] == 8.99 else row['TotalToPayAmount'], axis=1)

# print(df['DeliveryAmount'].unique())

df['kwota otrzymana'] = df['TotalToPayAmount']

# Reorder columns
df = df[['OrderDate', 'CombinedAddress', 'DeliveryMethod', 'TotalToPayAmount', 'TotalPaidAmountDivided', 'Wartość VAT', 'kwota otrzymana', 'DeliveryAmount']]

# Write the DataFrame to an Excel file
df.to_excel(output_excel_file_path, index=False)

print(f"Data has been saved to {output_excel_file_path}")
