import os
import sys
import win32com.client
import pyautogui
import ctypes
import time
import datetime


# Todo: configure Teams path
# Todo: parse data from calendar outlook
# Todo: look for TA Daily Scrum
# Todo: check daily when TA scrum start (time, date)
# Todo: join teams meeting when (1,2 3 time) minutes left

# teams_path = "C:/Users/erikas.treskovskis/AppData/Local/Microsoft/Teams/TeamsMeetingAddin/1.0.20.289.5/x86/Microsoft.Teams.AddinLoader.dll"

# os.startfile(teams_path)


def calendar_info():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    # outlook.getDefaultFolder(9) gives a list of all the meetings from our Outlook calendar
    # outlook.getDefaultFolder(6) gives all emails in the inbox.

    calendar = outlook.getDefaultFolder(9).Items
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')
    # outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace('MAPI')
    # folder = outlook.Folders
    # output = [folder[f].Name for f in range(folder.Count)]
    # output = ['Treskovskis.Erikas@fmc-ag.com', 'erikas.treskovskis@auriga.com',
    #           'Public Folders - erikas.treskovskis@auriga.com', 'Online Archive - Treskovskis.Erikas@fmc-ag.com']
    # print(output)

    today_date = datetime.datetime.today() + datetime.timedelta(days=1)
    tomorrow_date = datetime.timedelta(days=1) + today_date
    begin_day = today_date.date().strftime("%m/%d/%Y")
    end_day = tomorrow_date.date().strftime("%m/%d/%Y")

    meeting_plan = calendar.Restrict("[Start] >= '" + begin_day + "' AND [END] <= '" + end_day + "'")

    for appointments in meeting_plan:
        # print(appointments.Start)
        # print(appointments.Subject)
        # print(appointments.Duration)
        print(appointments.Body)

        # print(appointments.MeetingStatus)
        # Shows when meeting starts
        print(appointments.Start)
        # shows organizer name: Bertasius, Ugnius
        print(appointments.GetOrganizer())
        # opens Outlook Meeting/Appointment page
        # print(appointments.Display())
        # shows occurrence of meetings
        print(appointments.GetRecurrencePattern())
        print(appointments.IsRecurring)

        break


calendar_info()
# msteams:{parsed_url}
# Todo: if dispatch object Outlook.Applicantion is cached then delete it
#  out = win32com.client.gencache.EnsureDispatch("Outlook.Application")
#  print(sys.modules[out.__module__].__file__)