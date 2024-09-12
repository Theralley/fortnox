# Fortnox Invoice Automation System

## Overview

This project automates the process of generating invoices from booking platforms, such as those used for tennis/padel courts, and integrates with the Fortnox API for invoice creation and customer management. The system can extract booking data from platforms, process it, and create corresponding invoices within Fortnox. Additionally, it includes features for email notifications, error handling, and token management.

## Features

- **Customer and Invoice Management**: Automatically create and manage customer data and invoices using Fortnox APIs.
- **Email Integration**: Downloads attachments from emails that may include booking confirmations or invoices.
- **Error Notifications**: Sends error emails when issues arise during processing.
- **Token Management**: Handles OAuth tokens for Fortnox API access, including refresh tokens for automatic renewal.
- **Data Cleanup**: Includes scripts for handling duplicate entries or other data management tasks.
- **Approval Workflow**: Manages an approval process for invoices or bookings (optional).

## Project Structure

- **customer.py**: Manages customer data and integrates with Fortnox to create and update customer records.
- **invoice.py**: Main logic for generating invoices from booking data and submitting them to Fortnox.
- **customer_token.py & invoice_token.py**: Handles API token management for customer and invoice-related operations.
- **error_email.py**: Sends notification emails when errors occur during data processing.
- **mail_downloader.py**: Downloads attachments from emails, potentially booking confirmations, and processes them.
- **double_delete.py**: Script to handle duplicate entry removal in the database or external systems.
- **approval-checker/**: Handles the approval of invoices or booking requests before they are processed.
- **downloaded_attachments/**: Directory where email attachments are stored after being downloaded.
- **fixer_codes/**: Contains code or data to fix issues with entries in invoices or bookings.
- **Jansson_ID.csv & Janssons kranar.txt**: Sample data files used for demonstration or testing.
- **invoice_done/**: Folder to store completed invoice data.
- **done_error/**: Folder to store logs or data related to errors during processing.

## Requirements

To install the necessary dependencies, run:

```bash
pip install -r requirements.txt
```

## Fortnox API Integration

The system uses Fortnox API to create and manage customers and invoices. The following token management files are used:
- `customer_refresh_token.txt`: Stores the refresh token for customer-related API calls.
- `invoice_refresh_token.txt`: Stores the refresh token for invoice-related API calls.

Ensure that these tokens are kept up to date to maintain access to the Fortnox API.

## Usage

1. **Customer Management**: To create or update customers in Fortnox, run the `customer.py` script.
2. **Invoice Generation**: Use the `invoice.py` script to generate invoices based on bookings and send them to Fortnox.
3. **Email Attachments**: Run the `mail_downloader.py` script to automatically download booking confirmations or invoices from emails.
4. **Token Management**: Tokens for the Fortnox API are automatically refreshed using the scripts. Ensure that `customer_refresh_token.txt` and `invoice_refresh_token.txt` are configured correctly.
5. **Error Handling**: In case of errors, `error_email.py` will send an email notification to the designated recipients.

## Configuration

- **email_template.html**: Customize this file to change the format of error or notification emails.
- **Tokens**: Update the API tokens in `customer_token.txt` and `invoice_token.txt` files when necessary.

## Contributions

Hamren Holding & Development AB and Bengtsosn Börs och Finans AB
