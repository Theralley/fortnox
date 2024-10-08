import requests
import base64
import webbrowser
import time
from urllib.parse import urlparse, parse_qs


def refresh_access_token(refresh_token):
    token_url = "https://api.fortnox.se/oauth-v1/token"

    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Authorization': 'Basic {}'.format(base64.b64encode('{}:{}'.format(client_id, client_secret).encode()).decode())
    }

    body = {
        'grant_type': 'refresh_token',
        'client_id': client_id,
        'refresh_token': refresh_token
    }

    token_response = requests.post(token_url, headers=headers, data=body)

    if token_response.status_code == 200:
        token_data = token_response.json()
        new_access_token = token_data['access_token']
        new_refresh_token = token_data['refresh_token']
        return new_access_token, new_refresh_token
    else:
        print("Error refreshing access token.")
        return None, None


# Set up authentication
client_secret = "4V9XMejsDe"
client_id = "MYslrKdiUz4o"
auth_url = "https://apps.fortnox.se/oauth-v1/auth"

auth_params = {
    "client_id": client_id,
    "redirect_uri": "https://mysite.org/activation",
    "scope": "invoice",
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

if 'access_token' in token_data:
    access_token = token_data['access_token']
else:
    print("Error: Access token not found in response")

refresh_token = token_data['refresh_token']
expires_in = token_data['expires_in']

# Save the access token and refresh token to a file
with open('invoice_token.txt', 'w') as f:
    f.write(access_token)
    print("Invoice-token successfully printed.")

with open('invoice_refresh_token.txt', 'w') as f:
    f.write(refresh_token)
    print("Invoice-refresh-token successfully printed.")
