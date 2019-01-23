import xlwings as xw
import pandas as pd
import requests
import io
import re
from bs4 import BeautifulSoup


def etf_scrap():
    
    url_acwi = 'https://www.ishares.com/us/products/239600/ishares-msci-acwi-etf?qt=ACWI'
    req = requests.get(url_acwi)
    soup = BeautifulSoup(req.text, 'html.parser')
    csv_loc = 'https://www.ishares.com/' + soup.find('a', string=re.compile('(?i)detailed holdings'))['href']
    data = requests.get(csv_loc)
    
    cols = '(?i)ticker|name|weight|sector|country'
    csv_from = re.search('ticker', data.text, re.IGNORECASE).start()
    df = pd.read_csv(io.StringIO(data.text[csv_from:])).filter(regex=cols).dropna()
    
    
    wb = xw.Book.caller()
    wb.sheets['data'].range("B2").value = df


if __name__ == '__main__':
    etf_scrap()