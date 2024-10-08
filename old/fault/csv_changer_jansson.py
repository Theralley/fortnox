import os
import pandas as pd

# Set directory
directory = os.getcwd()

# Find all csv files in the directory
csv_files = [f for f in os.listdir(directory) if f.endswith('.csv') and 'jansson' in f]

# Loop through csv files and convert to xlsx
for csv_file in csv_files:
    # Read the csv file with the correct encoding
    df = pd.read_csv(csv_file, encoding='ISO-8859-1')

    # Get the name of the csv file without the extension
    file_name = os.path.splitext(csv_file)[0]

    # Create a new xlsx file with the same name as the csv file
    xlsx_file = file_name + '.xlsx'

    # Combine app_service columns
    app_service_cols = [col for col in df.columns if 'app_service' in col]
    df['app_service'] = df[app_service_cols].apply(lambda x: ' & '.join([val for val in x.drop_duplicates().astype(str) if val != 'nan']), axis=1)


    # Count number of 'godkänt' values and create new column 'days hire'
    app_status_cols = [col for col in df.columns if 'app_status' in col]
    godkant_count = df[app_status_cols].apply(lambda x: x.str.contains('Godkänd').sum(), axis=1)
    df['days_hire'] = godkant_count

    # Remove individual app_service columns
    df.drop(columns=app_service_cols, inplace=True)

    # Save the dataframe to the xlsx file
    df.to_excel(xlsx_file, index=False)

    # Print confirmation message
    print(f'{csv_file} converted to {xlsx_file}')
