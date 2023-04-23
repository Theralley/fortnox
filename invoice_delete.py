import requests
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

def delete_invoice(invoice_number):
    url = f"{invoice_url}/{invoice_number}/cancel"
    response = requests.put(url, headers=invoice_headers)

    if response.status_code == 200 or response.status_code == 201:
        print(f"Successfully deleted invoice {invoice_number}")
    elif response.status_code == 401 and response.content == b'{"message":"unauthorized"}':
        print("Access token expired, refreshing token...")
        os.system("python invoice_token.py")
        with open("invoice_token.txt", "r") as f:
            invoice_access_token = f.read().strip()
        invoice_headers['Authorization'] = f'Bearer {invoice_access_token}'
        delete_invoice(invoice_number)
    else:
        print(f"Error deleting invoice {invoice_number}: {response.content}")

if __name__ == "__main__":
    start = int(input("Enter the start of the invoice range: "))
    end = int(input("Enter the end of the invoice range: "))

    for invoice_number in range(start, end + 1):
        delete_invoice(invoice_number)
