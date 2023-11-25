import csv
# import tkinter as Tk
from tkinter import filedialog, Tk
import pandas as pd

def get_input_file_path():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select Input CSV File", filetypes=[("CSV files", "*.csv")])
    return file_path

def process_order_report(input_file_path, output_excel_file_path):
    # Define the columns you want to select
    selected_columns = ['OrderDate', 'BuyerName', 'BuyerPhone', 'BuyerAddress', 'BuyerZip', 'BuyerCity', 'BuyerCountryCode', 'DeliveryMethod', 'DeliveryAmount', 'TotalToPayAmount']

    # Create DataFrames to store selected data
    df_selected = pd.DataFrame(columns=selected_columns)
    df_line_item = process_line_item_data(input_file_path)

    # Open the input CSV file in read mode
    with open(input_file_path, 'r', newline='', encoding='utf-8') as input_csvfile:
        # Create a CSV reader object
        csv_reader = csv.DictReader(input_csvfile)

        # Check if all selected columns are present in the header
        missing_columns = [column for column in selected_columns if column not in csv_reader.fieldnames]
        if missing_columns:
            raise ValueError(f"Columns {missing_columns} not found in the CSV file.")

        # Iterate over rows in the CSV file
        for row in csv_reader:
            # Check if the row is empty
            if not any(row.values()):
                continue  # Skip to the next iteration

            # Process the non-empty row
            try:
                # Convert 'TotalToPayAmount' column to numeric values
                row['TotalToPayAmount'] = pd.to_numeric(row['TotalToPayAmount'], errors='coerce')

                # Convert 'OrderDate' to datetime format
                row['OrderDate'] = pd.to_datetime(row['OrderDate']).date()

                # Concatenate selected columns into a single column
                row['CombinedAddress'] = ', '.join([row['BuyerName'], row['BuyerAddress'], row['BuyerZip'], row['BuyerCity'], row['BuyerCountryCode'], row['BuyerPhone']])

                # Add a new column 'TotalPaidAmountDivided' by dividing 'TotalToPayAmount' by 1.23
                row['TotalPaidAmountDivided'] = row['TotalToPayAmount'] / 1.23

                # Add a new column 'Difference' by subtracting 'TotalPaidAmountDivided' from 'TotalToPayAmount'
                row['Wartość VAT'] = row['TotalToPayAmount'] - row['TotalPaidAmountDivided']

                # Add a new column 'kwota otrzymana'
                row['kwota otrzymana'] = row['TotalToPayAmount']

                # Reorder columns and append selected data to df_selected
                selected_data = [row[column] for column in selected_columns]
                df_selected = df_selected.append(pd.Series(selected_data, index=df_selected.columns), ignore_index=True)
            except Exception as e:
                print(f"Error processing row: {row}")
                print(f"Error details: {e}")
                continue

    # Save DataFrames to separate sheets in the same Excel file
    with pd.ExcelWriter(output_excel_file_path, engine='xlsxwriter') as writer:
        df_selected.to_excel(writer, sheet_name='SelectedData', index=False)
        df_line_item.to_excel(writer, sheet_name='LineItemData', index=False)

    print("Data has been saved to", output_excel_file_path)

def process_line_item_data(input_file_path):
    # Define the columns for lineItem data
    line_item_columns = ['Type', 'OrderId', 'LineItemId', 'Name', 'Quantity']

    # Create a DataFrame to store lineItem data
    df_line_item = pd.DataFrame(columns=line_item_columns)

    # Open the input CSV file in read mode
    with open(input_file_path, 'r', newline='', encoding='utf-8') as input_csvfile:
        # Create a CSV reader object
        csv_reader = csv.DictReader(input_csvfile, fieldnames=line_item_columns)

        # Iterate over rows in the CSV file
        for row in csv_reader:
            # Check if the row is a lineItem
            if row['Type'] == 'lineItem':
                # Append the row to df_line_item
                df_line_item = df_line_item.append(row, ignore_index=True)

    # Set to store unique OrderIds
    unique_order_ids = set()
    # Dictionary to store product names for OrderIds with a count of 1
    order_id_product_names = {}
    # List to store encountered OrderIds in the order they appear
    order_id_sequence = []

    # Iterate over rows in the DataFrame
    for index, row in df_line_item.iterrows():
        # Extract relevant information
        order_id = row['OrderId']
        product_name = row['Name']

        # Add the unique OrderId to the set and the sequence list
        if order_id not in unique_order_ids:
            unique_order_ids.add(order_id)
            order_id_sequence.append(order_id)

            # Check if the product count is 1 and store the product name
            if order_id_product_names.get(order_id, 0) == 0:
                order_id_product_names[order_id] = product_name

    # Create a DataFrame
    data = {'OrderId': order_id_sequence}
    df_order_count = pd.DataFrame(data)

    # Add a column with the count of orders for each unique OrderId
    df_order_count['OrderCount'] = [df_line_item[df_line_item['OrderId'] == unique_order_id].shape[0] for unique_order_id in order_id_sequence]

    # Replace 'OrderCount' with 'Name' when the count is 1
    df_order_count.loc[df_order_count['OrderCount'] == 1, 'OrderCount'] = [order_id_product_names[order_id] for order_id in df_order_count.loc[df_order_count['OrderCount'] == 1, 'OrderId']]

    return df_order_count
    
# Replace 'input_file.csv' and 'output_file.xlsx' with your actual file paths
# input_file_path = 'orderReport-20231101-20231124.csv'
# input_file_path = 'orderReport-20231101-20231120.csv'


input_file_path = get_input_file_path()

output_excel_file_path = 'output_file1.xlsx'


# Call the function
process_order_report(input_file_path, output_excel_file_path)
