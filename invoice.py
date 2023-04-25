# Create the invoices
import requests
import json
import openpyxl
import os

# Set up authentication for invoice API
invoice_client_secret = "4V9XMejsDe"

# Check if invoice token file exists
if not os.path.isfile("invoice_token.txt"):
    # If file does not exist, run invoice_token.py script to generate token
    os.system("python invoice_token.py")

# Read the access token from file
with open("invoice_token.txt", "r") as f:
    invoice_access_token = f.read().strip()

# Set the API endpoint and authentication headers for invoice API
invoice_url = 'https://api.fortnox.se/3/invoices'

invoice_headers = {
    'Authorization': f'Bearer {invoice_access_token}',
    'Client-Secret': invoice_client_secret,
    'Content-Type': 'application/json',
    'Accept': 'application/json'
}

# Set up authentication for customer API
customer_client_secret = "KSnpxgS5na"

# Check if customer token file exists
if not os.path.isfile("customer_token.txt"):
    # If file does not exist, run customer_token.py script to generate token
    os.system("python customer_token.py")

# Read the access token from file
with open("customer_token.txt", "r") as f:
    customer_access_token = f.read().strip()

# Set the API endpoint and authentication headers for customer API
customer_url = 'https://api.fortnox.se/3/customers'

customer_headers = {
    'Authorization': f'Bearer {customer_access_token}',
    'Client-Secret': customer_client_secret,
    'Content-Type': 'application/json',
    'Accept': 'application/json'
}

# Load the workbook
wb = openpyxl.load_workbook('downloaded_attachments/test.xlsx')
ws = wb.active

# Find the column indices for the required columns
app_quantity_index = None
final_price_index = None
article_number_index = None
customer_name_index = None
customer_phone_index = None
customer_number_index = None

for i, cell in enumerate(ws[1]):
    if cell.value == "app_quantity_1":
        app_quantity_index = i + 1
    elif cell.value == "final_price":
        final_price_index = i + 1
    elif cell.value == "ID":
        article_number_index = i + 1
    elif cell.value == "Namn":
        customer_name_index = i + 1
    elif cell.value == "Telefonummer":
        customer_phone_index = i + 1
    elif cell.value == "Personummer":
        customer_number_index = i + 1
    elif cell.value == "formname":
        formname_index = i + 1
    elif cell.value == "Ange hur du vill att vi ska vinterförvara just din båt.":
        description_index = i + 1

# Check if all the required columns are found
if None in [app_quantity_index, final_price_index, article_number_index, customer_name_index, customer_phone_index, customer_number_index]:
    print("Error: Required column not found")
    exit()

# Create the invoices
for row in ws.iter_rows(min_row=2, values_only=True):
    customer_name = row[customer_name_index - 1]
    customer_phone = row[customer_phone_index - 1]
    quantity = row[app_quantity_index - 1]
    article_number =row[article_number_index - 1]
    formname = row[formname_index - 1]
    price = row[final_price_index - 1]
    description = row[description_index - 1]
    customer_number = None

    # Read the customer data from the text file
    filename = f"{formname}.txt"
    with open(filename) as f:
        customer_data = f.readlines()

    # Search for the customer in the customers text file by name and phone number
    with open(f"{formname}.txt", "r") as f:
        for line in f:
            if customer_name in line and str(customer_phone) in line:
                customer_number = line.split("Customer number: ")[1].strip()
                break

    if customer_number is None:
        print(f'Error: Customer {customer_name} with phone number {customer_phone} not found')
        continue

    # Extract the invoice data from the Excel sheet
    invoice_data = {
        'Invoice': {
            'CustomerNumber': customer_number,
            'InvoiceRows': []
        }
    }

    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] != article_number:
            continue

        if row[1] is None or row[3] is None or row[4] is None:
            continue

        # Add the invoice row to the invoice data
        invoice_row = {
            #'ArticleNumber': article_number,
            'Unit': 'st', # add the unit of measurement
            'DeliveredQuantity': quantity, # should be DeliveredQuantity instead of Quantity
            'Price': price,
            'Description' : str(description)
        }
        invoice_data['Invoice']['InvoiceRows'].append(invoice_row)

    # Send a POST request to the Fortnox API to create a new invoice
    response = requests.post(invoice_url, headers=invoice_headers, json=invoice_data)

    if response.status_code == 200 or response.status_code == 201:
        print('Invoice successfully created for '+ customer_name +'!')
        response_data = response.json()
        #print(response_data.keys()) # print the keys present in the dictionary
        try:
            invoice_number = response_data['Invoice']['DocumentNumber'] # use the correct key for the invoice number
        except KeyError:
            print(f"Error: {response.content}")
            continue

    # Check if access token has expired
    if response.status_code == 401:
        # If access token has expired, generate a new token and retry
        with open("invoice_token.txt", "r") as f:
            access_token = f.read().strip()
        invoice_headers['Authorization'] = f'Bearer {access_token}'
        response = requests.post(invoice_url, headers=invoice_headers, json=invoice_data)

        # Update the customer number in the Excel sheet
        for row in ws.iter_rows(min_row=2, values_only=True):
            customer_name = row[customer_name_index - 1]
            customer_phone = str(row[customer_phone_index - 1])  # convert to string
            # Search for the customer in the customer token file by name and phone number
            with open(f"{formname}.txt", "r") as f:
                for line in f:
                    if customer_name in line and customer_phone in line:
                        customer_number = line.split("Customer number: ")[1].strip()
                        row[customer_number_index - 1] = customer_number
                        break
                else:
                    #print(f'Error: Customer {customer_name} with phone number {customer_phone} not found in customer token file')
                    continue

        # Save the updated Excel sheet
        wb.save('test.xlsx')

    else:
        print(f'Error: {response.content}')
