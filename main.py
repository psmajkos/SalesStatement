import csv
# from tkinter import filedialog, Tk
import pandas as pd
import easygui
if __name__ == "__main__":

    def get_input_file_path():
        file_path = easygui.fileopenbox(title="Select Input CSV File", filetypes=["*.csv"])

        return file_path

    def process_order_report(input_file_path, output_excel_file_path):
        # Define the number of reserved rows
        reserved_rows = 3

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

                    # Check 'DeliveryMethod' value and clear the cell if 'inpost' is not a substring
                    if 'inpost' not in row['DeliveryMethod'].lower():
                        row['DeliveryMethod'] = None
                        row['DeliveryAmount'] = 0  # Clear 'DeliveryAmount' when 'DeliveryMethod' is empty


                    # Reorder columns and concatenate selected data to df_selected
                    selected_data = [row[column] for column in selected_columns]
                    df_selected = pd.concat([df_selected, pd.DataFrame([selected_data], columns=df_selected.columns)], ignore_index=True)


                except Exception as e:
                    print(f"Error processing row: {row}")
                    print(f"Error details: {e}")
                    continue

        df_selected['Brutto z wysyłką'] = pd.to_numeric(df_selected['DeliveryAmount'], errors='coerce') + pd.to_numeric(df_selected['TotalToPayAmount'], errors='coerce')

        # Handle non-numeric values in 'SumAmount' (replace NaN with 0 or any other value as needed)
        df_selected['Brutto z wysyłką'].fillna(0, inplace=True)

        df_selected['wartość netto'] = pd.to_numeric(df_selected['TotalToPayAmount'], errors='coerce') / 1.23

        df_selected['Stawka VAT'] = "23%"

        df_selected["wartość VAT"] = pd.to_numeric(df_selected['TotalToPayAmount'], errors='coerce') - pd.to_numeric(df_selected['wartość netto'], errors='coerce')

        df_selected["Kwota otrzymana"] = pd.to_numeric(df_selected['TotalToPayAmount'], errors='coerce')

        df_selected['Faktura VAT'] = ''

        df_selected['Firma/os. p.'] = ''

        try:
            existing_selected_data = pd.read_excel(output_excel_file_path, sheet_name='SelectedData')
            print("Existing SelectedData:")
            print(existing_selected_data)
        except FileNotFoundError:
            existing_selected_data = pd.DataFrame()
            print("No existing SelectedData file found.")


        # Concatenate existing data with the new data
        df_selected = pd.concat([existing_selected_data, df_selected], ignore_index=True)

        with pd.ExcelWriter(output_excel_file_path, engine='xlsxwriter') as writer:
            # Add a sheet for additional data or headers
            additional_data_sheet = pd.DataFrame({'Header1': ['Value1'], 'Header2': ['Value2'], 'Header3': ['Value3']})
            additional_data_sheet.to_excel(writer, sheet_name='AdditionalData', index=False)

            # Add a sheet for line item data
            df_line_item.to_excel(writer, sheet_name='LineItemData', index=False, startrow=reserved_rows)

            # Add a sheet for selected data
            df_selected.to_excel(writer, sheet_name='SelectedData', index=False, startrow=reserved_rows)

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
                    # Create a DataFrame from the row and append it to df_line_item
                    df_line_item = pd.concat([df_line_item, pd.DataFrame([row])], ignore_index=True)

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
        
    if __name__ == "__main__":
        input_file_path = get_input_file_path()

        output_excel_file_path = 'Zestawienie_pansen_extra1.xlsx'

        # Call the function
        process_order_report(input_file_path, output_excel_file_path)
