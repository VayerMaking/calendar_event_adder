import pandas as pd

from gcsa.event import Event
from gcsa.google_calendar import GoogleCalendar
from gcsa.recurrence import Recurrence, DAILY, SU, SA, WEEKLY

from beautiful_date import *

excel_file = 'timesheet_tues.xlsx'
timesheet = pd.read_excel(excel_file)
print(timesheet)

for i in timesheet.index:
    # print(timesheet['day'][i], timesheet['time'][i], timesheet['subject'][i])
    
    calendar = GoogleCalendar('martin.vayer@gmail.com', credentials_path='client_secret_251214619756-gli6hcekqmbv4rqqju0mn03lao4o5k87.apps.googleusercontent.com.json')
    # start_time = float(timesheet['time'][i].split('-')[0].replace(':', '.'))
    start_time = timesheet['time'][i].split('-')[0]
    end_time = timesheet['time'][i].split('-')[1]

    day = timesheet['day'][i]
    MONDAY = D.today() + MO
    TUESDAY = D.today() + TU
    WEDNESDAY = D.today() + WE
    THURSDAY = D.today() + TH
    FRIDAY = D.today() + FR
    if day == "MO":
     DAY = MONDAY
    elif day == "TU":
     DAY = TUESDAY
    elif day == "WE":
     DAY = WEDNESDAY
    elif day == "TH":
     DAY = THURSDAY
    elif day == "FR":
     DAY = FRIDAY
    
    #print(int(start_time))
    # print(type(start_time))
    start_time_hour = int(start_time[:2])
    start_time_minutes = int(start_time[3:])
    end_time_hour = int(end_time[:2])
    end_time_minutes = int(end_time[3:])
    # print(str(start_time_hour) + " + " + str(start_time_minutes))

    # print(str(end_time_hour) + " + " + str(end_time_minutes))
    # print(start_time == '13 : 40')
    #print(day == "MO")
    
    event = Event(
         str(timesheet['subject'][i]),
         start=(DAY)[start_time_hour:start_time_minutes],
         end=(DAY)[end_time_hour:end_time_minutes],
         recurrence=[Recurrence.rule(freq=WEEKLY)]
         
     )
    print(event)
    calendar.add_event(event)

for event in calendar:
    print(event)