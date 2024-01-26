import pandas as pd
import requests
from datetime import datetime
import os
from dotenv import load_dotenv
load_dotenv()


def get_exchange_rate(api_key, base_currency='USD', target_currency='AUD'):

    # Sending a GET request to the API
    url = f'https://v6.exchangerate-api.com/v6/{api_key}/latest/{base_currency}'
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()

        # Checking if the API returns a success response
        if data['result'] == 'success':
            exchange_rate = data['conversion_rates'].get(target_currency)
            return exchange_rate
        else:
            return {500, f"Error from API: {data.get('error')}"}

    # If the API returns an error, we return an error message
    except requests.exceptions.RequestException as e:
        return {500, f"Error sending API GET request: {e}"}


# Function to convert the amount to USD
def convert_to_usd(amount, exchange_rate):
    if exchange_rate is not None:
        return amount / exchange_rate
    else:
        return None


def create_excel_file(data):
    try:
        df = pd.DataFrame(data)
        today = datetime.now().strftime('%Y-%m-%d')
        file_name = f'exchange_rates_{today}.xlsx'
        df.to_excel(file_name, index=False)
        print(f"Excel file '{file_name}' created successfully.")

    except Exception as e:
        return {500, f"Error creating Excel file: {e}"}


def main(*target_currencies):

    # Checking if the API key is valid
    api_key = os.getenv('API_KEY')
    if not api_key:
        print("Please provide your Exchange Rate API key.")
        return
    amount_to_convert = 1

    # Creating a dictionary to store the data
    data = {'Rate Type': [], 'Date': [], 'Currency_From': [],
            'Value_From': [], 'Currency_To': [], 'Value_To': []}

    # Getting the exchange rate for each target currency
    for target_currency in target_currencies:
        exchange_rate = get_exchange_rate(
            api_key, target_currency=target_currency)

        # Adding data to the dictionary
        data['Rate Type'].append('Spot rate')
        data['Date'].append(datetime.now().strftime('%d/%m/%Y'))
        data['Currency_From'].append(target_currency)
        data['Value_From'].append(amount_to_convert)
        data['Currency_To'].append('USD')

        # Checking data validity and printing results
        if exchange_rate is not None:
            print(f"Exchange Rate ({target_currency} to USD): {exchange_rate}")
            converted_amount = convert_to_usd(amount_to_convert, exchange_rate)

            if converted_amount is not None:
                data['Value_To'].append(converted_amount)
            else:
                print(f"Error converting amount to USD")

        # If the API returns an error, we add 'Error' to the dictionary
        else:
            data['Value_To'].append('Error')
            print(f"Error getting exchange rate for {target_currency}")

    # Creating the Excel file
    create_excel_file(data)


if __name__ == "__main__":
    main('AUD', 'CAD', 'CHF', 'MXN', 'EUR',
         'GBP', 'HKD', 'JPY', 'NZD', 'USD', 'CNH')
    # CNH isn't supported by the API, so it will return an error. I added MXN in its place.
    # CNH is left in the list to show how the error is handled.
