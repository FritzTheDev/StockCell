from alpha_vantage.timeseries import TimeSeries
import datetime
import json
import openpyxl
ts = TimeSeries('N1RDW0MA9WZPJGCK')
def calendarnormalize():
    today = datetime.date.today().strftime('%Y-%m-%d')
    weekday = datetime.datetime.today().weekday()
    backoneday = today
    if weekday < 5:
        return today
    elif weekday ==5:
print(calendarNormalize())
