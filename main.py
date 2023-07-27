import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

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

#Gets relevant data and update Excel spreadsheet as needed
def updateSpreadSheet():
    #Grab values from other functions
    prices = getPoolPrice()
    tng = getTNG()
    #TODO: Grab data from Excel spreadsheet, see if it needs to be updated, update accordingly
    wb = load_workbook(filename='gasPrices.xlsx')
    sheet = wb.active

    # Create a dictionary for the new data
    newData = {hour: (price if price != -1 else 'NULL') for hour, price in prices}

    #Read the existing data into a dictionary
    existingData = {row[0].value: row[1].value for row in sheet.iter_rows(min_row=2, max_col=2, values_only=True)}

    # Iterate over the new data and merge it with the existing data
    for hour, price in newData.items():
        if hour in existingData and price != 'NULL':
            existingData[hour] = price

    # Combine new data and existing data
    combinedData = {**newData, **existingData}

    # Write the combined data back to the Excel file
    for i, (hour, price) in enumerate(combinedData.items(), start=2):
        sheet.cell(row=i, column=1, value=hour)
        sheet.cell(row=i, column=2, value=price)
        sheet.cell(row=i, column=3, value=tng)

    #Save the new data
    wb.save("gasPrices.xlsx")

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    updateSpreadSheet()