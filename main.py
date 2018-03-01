#import the api wrapper, date-time, json, and an excel interface module
from alpha_vantage.timeseries import TimeSeries
import datetime
import json
from openpyxl import Workbook
from decimal import Decimal
date = None
workingRow = 2
ts = TimeSeries('N1RDW0MA9WZPJGCK')
introMessage = 'Welcome to StockCell,'
tickerInput =  'Please input your desired ticker   > '
tickerError = 'Not a valid ticker'
dateInput = 'What date would you like data for? > '
dateError = 'Please input date in the YYYY-MM-DD format'
#Commenting out this function because I'm skipping the automatic time functionality
#at least for now

# def calendarnormalize():
#     #sets variables that do the heavy lifting
#     today = datetime.datetime.utcnow().strftime('%Y-%m-%d')
#     weekdaycheck = 0#datetime.datetime.today().weekday()
#     utcformonday = datetime.datetime.utcnow().__str__()
#     backoneday = today[:9] + str(((int(today[9]))-1))
#     backtwodays = today[:9] + str(((int(today[9]))-2))
#     backthreedays = utcformonday[:9] + str(((int(today[9]))-3))
#     #if-else to use the variables above
#     if weekdaycheck == 0:
#         print(utcformonday)
#     elif weekdaycheck < 5:
#         return today
#     elif weekdaycheck ==5:
#         return backoneday
#     elif weekdaycheck ==6:
#         return backtwodays
#     else:
#         return
# print(calendarnormalize())
def api_call(ticker):
    while True:
        try:
            jsondict, meta = ts.get_daily(ticker)
            return jsondict
            break
        except ValueError:
            print(errorMessage)
def date_entry(): # TODO: Regex for date entry validation
    while True:
        try:
            global date
            date = (input(dateInput))
            return date
        except ValueError:
            print(dateError)
def data_reader(data, date):
    open_value = data[date]['1. open']
    high_value = data[date]['2. high']
    low_value = data[date]['3. low']
    close_value = data[date]['4. close']
    vol_value = data[date]['5. volume']
    value_dict = (open_value, high_value, low_value, close_value, vol_value)
    return value_dict
def rowfill(fivevalues):
    global workingRow
    ws['C'+ str(workingRow)] = fivevalues[0]
    ws['D'+ str(workingRow)] = fivevalues[1]
    ws['E'+ str(workingRow)] = fivevalues[2]
    ws['F'+ str(workingRow)] = fivevalues[3]
    ws['G'+ str(workingRow)] = fivevalues[4]
    workingRow += workingRow

print(introMessage)
fivevalues = data_reader(api_call(input(tickerInput)), date_entry())
wb = Workbook()
ws = wb.active
ws['A1'] = 'Date'
ws['B1'] = 'Ticker'
ws['C1'] = 'Open'
ws['D1'] = 'High'
ws['E1'] = 'Low'
ws['F1'] = 'Close'
ws['G1'] = 'Volume'
rowfill(fivevalues)
rowfill(fivevalues)
wb.save("tester.xlsx")
