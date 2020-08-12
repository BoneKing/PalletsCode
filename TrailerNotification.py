#Created By Andy Mahoney
#Clarkson University 
#Under the GPL2.0 License
#install requirements - Python 3.8, pip install pandas xlrd plyer
import pandas as pd 
import xlrd
import time
import datetime
from plyer import notification

today = datetime.date.today() #todays date
withinAlert = today + datetime.timedelta(days = 30) #alert within 30 days of today
#print(today)
TrailerStatus = pd.read_excel("Trailer status.xlsx") #read in the excel file
TrailerStatus["Inspection Due"] = pd.to_datetime(TrailerStatus["Inspection Due"]) #parce all data in column "Inspection Due" to date format
iterator = 0 #iterator to go through inspection due dates
for i in TrailerStatus["Inspection Due"]: 
    dateToCheck = TrailerStatus["Inspection Due"][iterator] #get date
    print(TrailerStatus["Inspection Due"][iterator])
    if dateToCheck < withinAlert: #if the date to check is within range
        print("ALERT")
        TrailerNum = int(TrailerStatus["Trailer #"][iterator]) #gets trailer number and puts it as an int
        messageBody = 'Trailer #' + str(TrailerNum) + ' on ' + dateToCheck.strftime("%m-%d-%Y") #notification message
        notification.notify(
            title= 'ALERT: Inspection Due',
            message= messageBody,
            app_icon=None,  
            timeout=10,  # seconds
        )
        time.sleep(5) #pause 5 seconds between notifications
    else:
        print("All Good")

    iterator += 1


time.sleep(86400)
