import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

def get_exchange_rate():
    # Send a request to the website
    response = requests.get("https://cambioticino.ch")
    soup = BeautifulSoup(response.text, "html.parser")

    # Find the exchange rate
    exchange_rate = (soup.find("span", {"id": "CHFtoEUR"}).text)

    return exchange_rate

def save_to_excel(exchange_rate):
    # Load or create the Excel file
    try:
        df = pd.read_excel("exchange_rates.xlsx")
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Date", "Exchange Rate"])

    # Save the exchange rate to the Excel file
    new_df = pd.DataFrame({"Date": [pd.Timestamp.now()], "Exchange Rate": [exchange_rate]})
    df = pd.concat([df, new_df], ignore_index=True)
    df.to_excel("exchange_rates.xlsx", index=False)

if __name__ == "__main__":
    while True:
        exchange_rate = get_exchange_rate()
        save_to_excel(exchange_rate)    
        time.sleep(180) # Wait 30 minutes
