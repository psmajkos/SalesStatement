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
    csv_reader = csv.reader(csvfile)

    # Get the header
    header = next(csv_reader)

    try:
        # Get the indices of the 'Type', 'OrderId', 'LineItemId', 'Name', and 'Quantity' columns
        type_index = header.index('Type')
        order_id_index = header.index('OrderId')
        line_item_id_index = header.index('LineItemId')
        name_index = header.index('Name')
        quantity_index = header.index('Quantity')

        # Set to store unique OrderIds
        unique_order_ids = set()
        # Dictionary to store product names for OrderIds with a count of 1
        order_id_product_names = {}
        # List to store encountered OrderIds in the order they appear
        order_id_sequence = []

        # Loop through the rows
        for row in csv_reader:
            # Check if the row has values for 'Type', 'OrderId', 'LineItemId', 'Name', and 'Quantity'
            if row and len(row) > type_index and len(row) > order_id_index and len(row) > line_item_id_index and len(row) > name_index and len(row) > quantity_index:
                # Skip rows that contain header labels
                if row[type_index] == 'lineItem' and row[order_id_index] != 'OrderId' and row[line_item_id_index] != 'LineItemId' and row[name_index] != 'Name' and row[quantity_index] != 'Quantity':
                    order_id = row[order_id_index]
                    product_name = row[name_index]

                    # Add the unique OrderId to the set and the sequence list
                    if order_id not in unique_order_ids:
                        unique_order_ids.add(order_id)
                        order_id_sequence.append(order_id)

                        # Check if the product count is 1 and store the product name
                        if order_id_product_names.get(order_id, 0) == 0:
                            order_id_product_names[order_id] = product_name

        # Create a DataFrame
        data = {'OrderId': order_id_sequence}
        df = pd.DataFrame(data)

        # Add a column with the count of orders for each unique OrderId
        df['OrderCount'] = [sum(1 for row in csv.reader(open(csv_file_path)) if unique_order_id in row) for unique_order_id in order_id_sequence]

        # Replace values based on conditions
        df['OrderCount'] = df.apply(lambda row: order_id_product_names[row['OrderId']] if row['OrderCount'] == 1 else 'mix', axis=1)

        # Save DataFrame to Excel
        df.to_excel(output_excel_file_path, index=False)

        print(f"Data has been saved to {output_excel_file_path}")

    except ValueError as e:
        print(f'Error: {e}')
