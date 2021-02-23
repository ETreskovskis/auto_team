import os
import sys

import pywintypes
import win32com.client
import win32process
import win32gui
import win32api
import win32con
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


# calendar_info()
# webbrowser
# msteams:{parsed_url}
# Todo: if dispatch object Outlook.Applicantion is cached then delete it
#  out = win32com.client.gencache.EnsureDispatch("Outlook.Application")
#  print(sys.modules[out.__module__].__file__)

# -------------------------------- ENUM invoked windows, get tid, pid, get window name ---------------------------------


def get_tid_and_pid(handle, data: list):
    tid, pid = win32process.GetWindowThreadProcessId(handle)
    print(f"TID: {tid} PID: {pid}")


def enum_windows(callback_func, object_str):
    return win32gui.EnumWindows(callback_func, object_str)

# print(enum_windows())


def enum_processes():
    return win32process.EnumProcesses()


def get_handle_object(pid: int):
    try:
        # return win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, win32con.FALSE, pid)
        return win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, win32con.FALSE, pid)
    except pywintypes.error as er:
        print(er)



# print(get_handle_object(18912))


def get_handle_process_module(handle):
    return win32process.EnumProcessModules(handle)


def get_window_text(handle, object):
    print(f"Window name: {win32gui.GetWindowText(handle)} {handle}")


pid_processes = enum_processes()
# print(pid_processes)
# pywintypes.error: (5, 'OpenProcess', 'Access is denied.')
# Regarding error check : https://docs.microsoft.com/en-us/windows/win32/procthread/process-security-and-access-rights?redirectedfrom=MSDN
# https://stackoverflow.com/questions/8543716/python-pywin32-access-denied


# print(enum_windows(get_window_text, "window name"))

# the window with which the user is currently working
current_window = win32gui.GetForegroundWindow()
print(current_window)
tid, pid = win32process.GetWindowThreadProcessId(current_window)
print(tid, pid)
win_handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, pid)
window_path = win32process.GetModuleFileNameEx(win_handle, 0)
print(win_handle)
print(window_path)
is_visible = win32gui.IsWindowVisible(current_window)
print(bool(is_visible))

# creates message box
# win32gui.MessageBox(current_window, "This is message box", "BOX", win32con.MB_HELP)
# cursor = win32gui.GetCursorInfo()
# print(cursor)


def get_window_info(hwnd, top_windows: list):
    tid, pid = win32process.GetWindowThreadProcessId(hwnd)
    top_windows.append(dict(handler=hwnd, tid=tid, pid=pid, name= win32gui.GetWindowText(hwnd)))


top_windows = []
win32gui.EnumWindows(get_window_info, top_windows)
print(top_windows)