import os
import win32com.client
import pyautogui
import time
import datetime

# Todo: configure Teams path
# Todo: parse data from calendar outlook
# Todo: look for TA Daily Scrum
# Todo: check daily when TA scrum start (time, date)
# Todo: join teams meeting when (1,2 3 time) minutes left

# teams_path = "C:/Users/erikas.treskovskis/AppData/Local/Microsoft/Teams/current/Teams.exe"

# os.startfile(teams_path)

def calendar_info():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    # outlook.getDefaultFolder(9) gives a list of all the meetings from our Outlook calendar
    # outlook.getDefaultFolder(6) gives all emails in the inbox.

    calendar = outlook.getDefaultFolder(9).Items
    calendar.Sort("[Start]")
    calendar.IncludeRecurrences = "True"
    print(type(calendar))
    print(dir(calendar))
    print(calendar.__dict__)

    today_date = datetime.datetime.today()
    tomorrow_date = datetime.timedelta(days=1) + today_date
    begin_day = today_date.date().strftime("%m/%d/%Y")
    end_day = tomorrow_date.date().strftime("%m/%d/%Y")

    meeting_plan = calendar.Restrict("[Start] >= '" + begin_day + "' AND [END] <= '" + end_day + "'")
    for appointments in meeting_plan:
        print(appointments.Start)
        print(appointments.Subject)
        print(appointments.Duration)

calendar_info()
