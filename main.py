import requests
import time
import pandas as pd
import os
import re
from datetime import datetime, timedelta
from bs4 import BeautifulSoup

def getPoolPrice():
    # Send a GET request to get website data and convert to string
    response = requests.get('http://ets.aeso.ca/ets_web/ip/Market/Reports/SMPriceReportServlet')
    websiteText = response.text

    #Parse the HTML for the correct table, find all entries in said table
    soup = BeautifulSoup(websiteText, 'html.parser')
    table = soup.find_all('table')[2]
    rows = table.find_all('tr')

    #For storing updates in
    hourlyPrices = []

    # Iterate over each row
    for row in rows:
        # Get all columns in this row
        cols = row.find_all('td')
        # If not the header row, extract data and append
        if len(cols) >= 2:
            hour = cols[0].text.strip()
            price = cols[1].text.strip()
            hourlyPrices.append((hour, price))

    # Print the hours and their prices
    for hour, price in hourlyPrices:
        print(f'Hour: {hour}, Price: {price}')

    #TODO: Combine this with TNG number, update when needed in sheet

def getTNG():
    


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    getPoolPrice()