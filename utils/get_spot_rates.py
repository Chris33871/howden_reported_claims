import pandas as pd
import requests
from datetime import datetime
import os
from dotenv import load_dotenv

load_dotenv()
API_KEY = os.getenv('API_KEY')


def get_exchange_rate(api_key, base_currency='USD', target_currency='AUD'):
    url = f'https://v6.exchangerate-api.com/v6/{api_key}/latest/{base_currency}'

    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()

        if data['result'] == 'success':
            exchange_rate = data['conversion_rates'].get(target_currency)
            return exchange_rate
        else:
            print(f"Error from API: {data.get('error')}")
            return None

    except requests.exceptions.RequestException as e:
        print(f"Error making API request: {e}")
        return None


def convert_to_usd(amount, exchange_rate):
    if exchange_rate is not None:
        return amount / exchange_rate
    else:
        return None


def create_excel_file(data):
    df = pd.DataFrame(data)
    today = datetime.now().strftime('%Y-%m-%d')
    file_name = f'exchange_rates_{today}.xlsx'
    df.to_excel(file_name, index=False)
    print(f"Excel file '{file_name}' created successfully.")


def main(*target_currencies):
    api_key = API_KEY

    if not api_key:
        print("Please provide your Exchange Rate API key.")
        return

    amount_to_convert = 1

    data = {'Rate Type': [], 'Date': [], 'Currency_From': [],
            'Value_From': [], 'Currency_To': [], 'Value_To': []}

    for target_currency in target_currencies:
        exchange_rate = get_exchange_rate(
            api_key, target_currency=target_currency)

        if exchange_rate is not None:
            print(f"Exchange Rate ({target_currency} to USD): {exchange_rate}")

            converted_amount = convert_to_usd(amount_to_convert, exchange_rate)

            if converted_amount is not None:
                data['Rate Type'].append('Spot rate')
                data['Date'].append(datetime.now().strftime('%d/%m/%Y'))
                data['Currency_From'].append(target_currency)
                data['Value_From'].append(amount_to_convert)
                data['Currency_To'].append('USD')
                data['Value_To'].append(converted_amount)
            else:
                print("Error converting amount to USD")

    create_excel_file(data)


if __name__ == "__main__":
    main('AUD', 'CAD', 'CHF', 'MXN', 'EUR', 'GBP', 'HKD', 'JPY', 'NZD', 'USD')
