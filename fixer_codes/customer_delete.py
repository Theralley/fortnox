import requests
import os

# Set up authentication
client_secret = "4V9XMejsDe"

# Check if customer token file exists
if not os.path.isfile("customer_token.txt"):
    # If file does not exist, run customer_token.py script to generate token
    os.system("python customer_token.py")

# Read the access token from file
with open("customer_token.txt", "r") as f:
    access_token = f.read().strip()

# Set the API endpoint and authentication headers
customer_url = 'https://api.fortnox.se/3/customers'

headers = {
    'Authorization': f'Bearer {access_token}',
    'Client-Secret': client_secret,
    'Content-Type': 'application/json',
    'Accept': 'application/json'
}

def delete_customer(customer_number):
    url = f"{customer_url}/{customer_number}"
    response = requests.delete(url, headers=headers)

    if response.status_code == 200 or response.status_code == 204:
        print(f"Successfully deleted customer with number {customer_number}")
    elif response.status_code == 401 and response.content == b'{"message":"unauthorized"}':
        print("Access token expired, refreshing token...")
        os.system("python customer_token.py")
        with open("customer_token.txt", "r") as f:
            access_token = f.read().strip()
        headers['Authorization'] = f'Bearer {access_token}'
        delete_customer(customer_number)
    else:
        print(f"Error deleting customer with number {customer_number}: {response.content}")

if __name__ == "__main__":
    start = int(input("Enter the start of the customer number range: "))
    end = int(input("Enter the end of the customer number range: "))

    for customer_number in range(start, end + 1):
        delete_customer(str(customer_number))
