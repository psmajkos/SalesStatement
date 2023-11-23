import csv

# Replace 'input_file.csv', 'output_file.csv', and 'output_file2.csv' with your actual file paths
input_file_path = 'orderReport-20231101-20231120.csv'
output_file_path = 'output_file.csv'
output_file_path2 = 'output_file2.csv'

# Open the input CSV file in read mode
with open(input_file_path, 'r', newline='', encoding='utf-8') as input_csvfile:
    # Create a CSV reader object
    csv_reader = csv.reader(input_csvfile)

    # Open the output CSV files in write mode
    with open(output_file_path, 'w', newline='', encoding='utf-8') as output_csvfile, \
            open(output_file_path2, 'w', newline='', encoding='utf-8') as output_csvfile2:
        # Create CSV writer objects
        csv_writer = csv.writer(output_csvfile)
        csv_writer2 = csv.writer(output_csvfile2)

        # Write header to the output files
        header = next(csv_reader)
        csv_writer.writerow(header)
        csv_writer2.writerow(header)

        # Iterate over rows in the CSV file
        for row in csv_reader:
            # Check if the row is empty
            if not any(row):
                continue  # Skip to the next iteration

            # Process the non-empty row (you can customize this part)
            print(row)

            # Write the row to the new CSV file
            csv_writer.writerow(row)

            # Check if the row is a lineItem
            if row[0] == "lineItem":
                # Write the lineItem row to the second output CSV file
                csv_writer2.writerow(row)

print("Rows saved to", output_file_path)
print("Filtered lineItem rows saved to", output_file_path2)
