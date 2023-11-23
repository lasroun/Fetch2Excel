from requests import Session
from requests.exceptions import ConnectionError, Timeout, TooManyRedirects
import json
import pandas as pd
from datetime import datetime

url = 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/map'
parameters = {
    'start': '1',
    'limit': '100',
    'sort': 'cmc_rank'
}
headers = {
    'Accepts': 'application/json',
    'X-CMC_PRO_API_KEY': '4bdefce5-9afd-4f7c-8ff3-768edecd0bf8',
}

session = Session()
session.headers.update(headers)

try:
    response = session.get(url, params=parameters)
    api_response = json.loads(response.text)

    # Extraire les informations spécifiques
    data = api_response['data']
    cryptos = [{
        "Rang": crypto.get("rank"),
        "Nom": crypto.get("name"),
        "Symbol": crypto.get("symbol")
    } for crypto in data if crypto.get('is_active') == 1]
    # Exporter vers Excel
    df = pd.DataFrame(cryptos)
    df.to_excel("cryptomonnaies.xlsx", index=False)

except (ConnectionError, Timeout, TooManyRedirects) as e:
    print(e)
