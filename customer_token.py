import requests
import base64
import webbrowser
import time
from urllib.parse import urlparse, parse_qs

# Set up authentication
client_secret = "4V9XMejsDe"
client_id = "MYslrKdiUz4o"
auth_url = "https://apps.fortnox.se/oauth-v1/auth"

auth_params = {
    "client_id": client_id,
    "redirect_uri": "https://mysite.org/activation",
    "scope": "customer",
    "state": "somestate123",
    "access_type": "offline",
    "response_type": "code",
    "account_type": "service"
}

# Authorize and get authorization code
auth_response = requests.get(auth_url, params=auth_params)
if auth_response.status_code != 200:
    print("Error authorizing with Fortnox API")
    print(auth_response.text)
    exit()

webbrowser.open(auth_response.url)
time.sleep(5)  # Wait for user to log in
redirect_url = input("Enter the URL you were redirected to: ")
query_params = parse_qs(urlparse(redirect_url).query)
auth_code = query_params["code"][0]

auth_code = query_params["code"][0]

token_url = "https://api.fortnox.se/oauth-v1/token"
redirect_uri = "https://mysite.org/activation"

# Set the request headers
headers = {
    'Content-Type': 'application/x-www-form-urlencoded',
    'Authorization': 'Basic {}'.format(base64.b64encode('{}:{}'.format(client_id, client_secret).encode()).decode())
}

# Set the request body
body = {
    'grant_type': 'authorization_code',
    'client_id': client_id,
    'redirect_uri': 'https://mysite.org/activation',
    'code': auth_code
}

# Send a POST request to the token URL using the requests library
token_response = requests.post(token_url, headers=headers, data=body)

# Parse the token response
token_data = token_response.json()
access_token = token_data['access_token']
refresh_token = token_data['refresh_token']
expires_in = token_data['expires_in']

# Save the access token to a file
with open('customer_token.txt', 'w') as f:
    f.write(access_token)
    print("Customer-token successfully printed.")
