import os
import pandas as pd

# Set directory
directory = os.getcwd()

# Find all csv files in the directory
csv_files = [f for f in os.listdir(directory) if f.endswith('.csv')]

# Loop through csv files and convert to xlsx
for csv_file in csv_files:
    # Read the csv file with the correct encoding
    df = pd.read_csv(csv_file, encoding='ISO-8859-1')

    # Get the name of the csv file without the extension
    file_name = os.path.splitext(csv_file)[0]

    # Create a new xlsx file with the same name as the csv file
    xlsx_file = file_name + '.xlsx'

    # Save the dataframe to the xlsx file
    df.to_excel(xlsx_file, index=False)

    # Print confirmation message
    print(f'{csv_file} converted to {xlsx_file}')
