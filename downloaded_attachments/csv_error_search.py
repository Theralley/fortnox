import os
import pandas as pd
from datetime import datetime

# Use the current working directory as the folder path
directory = os.getcwd()

# Get the Excel file names
xlsx_files = [f for f in os.listdir(directory) if f.endswith('.xlsx') and 'janssons_kranar_03_15_23' in f]

# Read the data as a DataFrame
for xlsx_file in xlsx_files:
    file_path = os.path.join(directory, xlsx_file)
    data = pd.read_excel(file_path, engine='openpyxl')

    # Check if the "days_hire" column exists, if not, create it and copy values from any column starting with "app_quantity"
    if 'days_hire' not in data.columns:
        app_quantity_col = None
        for col in data.columns:
            if col.startswith("app_quantity"):
                app_quantity_col = col
                break

        if app_quantity_col is not None:
            data['days_hire'] = data[app_quantity_col]

    # Filter the relevant columns
    relevant_columns = [
        'Namn', 'Telefonummer', 'final_price', 'app_service_1', 'days_hire'
    ]
    filtered_data = data[relevant_columns]

    # Iterate through the rows and generate invoices
    for index, row in filtered_data.iterrows():
        customer_name = row['Namn']
        customer_phone = row['Telefonummer']
        final_price = row['final_price']
        service = row['app_service_1']
        days_hire = row['days_hire']

        # Check if name is NaN and set to 'Nameless' if true
        if pd.isna(customer_name):
            customer_name = 'Nameless'
            data.at[index, 'Namn'] = customer_name

        # Check if the phone number contains non-digit characters
        if customer_phone is not None and not str(customer_phone).isdigit():
            # Move text to name and set phone number to '07000000'
            customer_name += f" {customer_phone}"
            data.at[index, 'Namn'] = customer_name
            customer_phone = "07000000"
            data.at[index, 'Telefonummer'] = customer_phone

        # Format the output
        output = f"Customer: {customer_name} {customer_phone} Final Price: {final_price} Service: {service} Days: {days_hire}"

        # Make sure customer_name and customer_phone are not None
        if customer_name is not None and customer_phone is not None:
            print(output)

    # Save the modified data back to the .xlsx file
    data.to_excel(file_path, index=False, engine='openpyxl')
