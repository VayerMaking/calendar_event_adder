import pandas as pd

from gcsa.event import Event
from gcsa.google_calendar import GoogleCalendar
from gcsa.recurrence import Recurrence, DAILY, SU, SA, WEEKLY

from beautiful_date import *

import config

excel_file = config.file_path
timesheet = pd.read_excel(excel_file)

for i in timesheet.index:    
    calendar = GoogleCalendar(config.email, credentials_path=config.credentials_path)
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
    

    start_time_hour = int(start_time[:2])
    start_time_minutes = int(start_time[3:])
    end_time_hour = int(end_time[:2])
    end_time_minutes = int(end_time[3:])
    
    event = Event(
         str(timesheet['subject'][i]),
         start=(DAY)[start_time_hour:start_time_minutes],
         end=(DAY)[end_time_hour:end_time_minutes],
         recurrence=[Recurrence.rule(freq=WEEKLY)]
         
     )

    calendar.add_event(event)

# for event in calendar:
#     print(event)