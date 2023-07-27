import requests
from bs4 import BeautifulSoup
import openpyxl

#Returns pool prices for the last 24 hours as a 2d array where
#hourlyPrices[X][0] = date/hour & hourlyPrices[X][1] = price
def getPoolPrice():
    # Send a get request to get website data and convert to string
    response = requests.get('http://ets.aeso.ca/ets_web/ip/Market/Reports/SMPriceReportServlet')
    websiteText = response.text

    #Parse the HTML for the correct table, find all entries in correct table
    soup = BeautifulSoup(websiteText, 'html.parser')
    table = soup.find_all('table')[2]
    rows = table.find_all('tr')

    #For storing updates in
    hourlyPrices = []

    # For each row in the table
    for row in rows:
        cols = row.find_all('td')
        # If not the header row, extract data and append
        if len(cols) >= 2:
            hour = cols[0].text.strip()
            price = cols[1].text.strip()
            if(price == '-'):
                hourlyPrices.append((hour, float(-1)))
            else:
                hourlyPrices.append((hour, float(price)))

    return hourlyPrices

#Returns TNG number as a float
def getTNG():
    #Grab the website data, filter it down to just show the H.R. Milner row as array
    response = requests.get('http://ets.aeso.ca/ets_web/ip/Market/Reports/CSDReportServlet')
    soup = BeautifulSoup(response.text, 'html.parser')
    table = soup.find_all('table')[9]
    hrmTable = table.find_all('tr')[14].find_all('td')
    #Return TNG for the hour as a float
    return (float(hrmTable[2].text.strip()))


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #getPoolPrice()
    #getTNG()