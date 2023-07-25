import requests
import time
import pandas as pd
import os
from datetime import datetime, timedelta
from bs4 import BeautifulSoup

def getPoolPrice():
    # Send a GET request to get website data and convert to string
    response = requests.get('http://ets.aeso.ca/ets_web/ip/Market/Reports/SMPriceReportServlet')
    websiteText = response.text

    #Get dates and convert to string
    dateTime = datetime.now()
    todayDate = dateTime.strftime("%m/%d/%Y")
    yesterdayDate = (dateTime - timedelta(days = 1)).strftime("%m/%d/%Y")


    #Find current day in the website's data
    dateIndex= websiteText.find(todayDate)
    #TODO: Get all non-null instances of today's (and previous day's dates, record price into excel (updating as needed))
    while(dateIndex != -1):
        dateIndex = websiteText.find(todayDate, dateIndex+1)

    yestIndex= websiteText.find(yesterdayDate)
    while(yestIndex != -1):
        yestIndex = websiteText.find(yesterdayDate, yestIndex+1)

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    getPoolPrice()