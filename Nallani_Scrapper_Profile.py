import json
import csv
import sys
from typing import Any, Dict
import requests
from bs4 import BeautifulSoup
import pandas as pd

def get_stock_profile_data(ticker_symbol: Any) -> Dict[str, Any]:
    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:122.0) Gecko/20100101 Firefox/122.0'}
    url = f'https://finance.yahoo.com/quote/{ticker_symbol}/profile'
    r = requests.get(url, headers=headers)
    soup = BeautifulSoup(r.text, 'html.parser')
    
    stock_profile = {
        'stock_name': soup.find('div', {'class':'D(ib) Mt(-5px) Maw(38%)--tab768 Maw(38%) Mend(10px) Ov(h) smartphone_Maw(85%) smartphone_Mend(0px)'}).find_all('div')[0].text,
        'Address': soup.find('div', {'class':'Mb(25px)'}).find_all('p')[0].text.strip(),
        'Key Executives1': soup.find('table', {'class':'W(100%)'}).find_all('tr')[1].find_all('td')[0].text.strip(),
        'Key Executives2': soup.find('table', {'class':'W(100%)'}).find_all('tr')[2].find_all('td')[0].text.strip(),
        'Key Executives3': soup.find('table', {'class':'W(100%)'}).find_all('tr')[3].find_all('td')[0].text.strip(),
        'Key Executives4': soup.find('table', {'class':'W(100%)'}).find_all('tr')[4].find_all('td')[0].text.strip(),
        'Key Executives5': soup.find('table', {'class':'W(100%)'}).find_all('tr')[5].find_all('td')[0].text.strip(),
        'Key Executives6': soup.find('table', {'class':'W(100%)'}).find_all('tr')[6].find_all('td')[0].text.strip(),
        'Key Executives7': soup.find('table', {'class':'W(100%)'}).find_all('tr')[7].find_all('td')[0].text.strip(),
        'Key Executives8': soup.find('table', {'class':'W(100%)'}).find_all('tr')[8].find_all('td')[0].text.strip(),
        'Key Executives9': soup.find('table', {'class':'W(100%)'}).find_all('tr')[9].find_all('td')[0].text.strip(),
        'Key Executives10': soup.find('table', {'class':'W(100%)'}).find_all('tr')[10].find_all('td')[0].text.strip(),
        'Description': soup.find('section', {'class':'quote-sub-section Mt(30px)'}).find('p').text.strip(),
    }
    return stock_profile

if len(sys.argv) < 2:
    print("Usage: python script.py <ticker_symbol1> <ticker_symbol2> ...")
    sys.exit(1)

ticker_symbols = sys.argv[1:]
stock_profile_data = [get_stock_profile_data(symbol) for symbol in ticker_symbols]

with open('duplicate_stock_profile_data.json', 'w', encoding='utf-8') as f:
    json.dump(stock_profile_data, f)

CSV_FILE_PATH = 'duplicate_stock_profile_data.csv'
with open(CSV_FILE_PATH, 'w', newline='', encoding='utf-8') as csvfile:
    fieldnames = stock_profile_data[0].keys()
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(stock_profile_data)

EXCEL_FILE_PATH = 'duplicate_stock_profile_data.xlsx'
df = pd.DataFrame(stock_profile_data)
df.to_excel(EXCEL_FILE_PATH, index=False)

print('Done!')
