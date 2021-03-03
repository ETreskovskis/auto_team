from __future__ import annotations
import re
import sys
from dataclasses import dataclass
from typing import Optional, List

import pywintypes
import win32com.client
import win32process
import win32gui
import win32api
import win32con
import webbrowser
import time
import datetime


def _for_debugging_purpose(ensure_dispatch):
    """If Dispatch object Outlook.Application is cached then delete the file. This may happen when
    win32com.client.gencache.EnsureDispatch("Outlook.Application") was called
    """
    # ensure_dispatch = win32com.client.gencache.EnsureDispatch("Outlook.Application")

    print(sys.modules[ensure_dispatch.__module__].__file__)


@dataclass(init=False)
class DataStorage:
    pass


class OutlookApi:

    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.folders = self.enumerate_outlook_folders()
        self.fail_flag = False

    def enumerate_outlook_folders(self) -> DataStorage:
        """ADD DOCS"""
        folders = DataStorage()

        for num in range(50):

            try:
                folder = self.outlook.GetDefaultFolder(num)
                setattr(folders, folder.Name, num)
            except pywintypes.com_error:
                pass

        return folders

    def sort_calendar_meeting_object(self):
        """ADD DOCS"""

        calendar = self.outlook.getDefaultFolder(self.folders.Calendar).Items
        calendar.IncludeRecurrences = True
        calendar.Sort("[Start]")

        today_date = datetime.datetime.today()
        tomorrow_date = datetime.timedelta(days=1) + today_date
        begin_day = today_date.date().strftime("%m/%d/%Y")
        end_day = tomorrow_date.date().strftime("%m/%d/%Y")

        # return Items collection of MeetingItem
        # https://docs.microsoft.com/en-us/office/vba/api/outlook.items.restrict
        meeting_plan = calendar.Restrict("[Start] >= '" + begin_day + "' AND [END] <= '" + end_day + "'")

        return meeting_plan

    @staticmethod
    def populate_meeting_events(event_items):
        """ADD DOCS"""

        for appointment in event_items:
            event = DataStorage()
            setattr(event, "Start", appointment.Start)
            setattr(event, "End", appointment.End)
            setattr(event, "Subject", appointment.Subject)
            setattr(event, "Duration", appointment.Duration)
            setattr(event, "GetOrganizer", appointment.GetOrganizer())
            setattr(event, "IsRecurring", appointment.IsRecurring)
            setattr(event, "GetRecurrencePattern", appointment.GetRecurrencePattern())
            setattr(event, "Body", appointment.Body)
            setattr(event, "Display", appointment.Display)

            yield event

    def parse_teams_meet_join_url(self, meeting_event: DataStorage) -> Optional[str]:
        """Parse Teams meet-join url from event Body. If body is absent then open Outlook Meeting Occurrence window"""

        if not meeting_event.Body:
            import warnings
            warnings.warn("Outlook calendar event (meeting email BODY) was not parsed")
            self.fail_flag = True
            return meeting_event.Display()
        general_url_pattern = re.compile(
            pattern="http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+")
        meeting_join = re.compile(pattern="meetup-join")
        results = re.findall(general_url_pattern, string=meeting_event.Body)
        format_result = [url.strip(">") for url in results]
        without_https = [url.strip("https:") for url in format_result]

        meet_url, *_ = [url for url in without_https if re.search(meeting_join, url)]
        return meet_url

    @staticmethod
    def open_teams_meet_via_url(url: str):
        """ADD DOCS"""

        full_url = f"msteams:{url}"
        webbrowser.open(full_url)


class EnumActiveWindows:

    def __init__(self):
        self.enum_windows = list()

    @staticmethod
    def _get_window_info(hwnd, enum_windows: list):
        """Callback function. Gets active window information like handler, PID, TID, name"""

        tid, pid = win32process.GetWindowThreadProcessId(hwnd)
        window_info = DataStorage()
        setattr(window_info, "handler", hwnd)
        setattr(window_info, "tid", tid)
        setattr(window_info, "pid", pid)
        setattr(window_info, "name", win32gui.GetWindowText(hwnd))
        enum_windows.append(window_info)

    @property
    def enumerate_windows(self):
        """ADD DOCS"""

        win32gui.EnumWindows(self._get_window_info, self.enum_windows)
        return self.enum_windows


enum = EnumActiveWindows()
data = enum.enumerate_windows
for d in data:
    print(d.__dict__)


class InvokeEvents:

    # Todo: parse exact names TA Daily Scrum (Teams window) and TA Daily Scrum - Meeting Occurrence (Outlook window)????
    outlook_window_name = "TA Daily Scrum - Meeting Occurrence"
    teams_window_name = "TA Daily Scrum"
    cursor_pos_for_outlook = (735, 186)
    cursor_teams_join_button = (1405, 750)
    cursor_microphone_on_off = (1092, 524)
    cursor_background_filters = (534, 690)

    def retrieve_current_window_handler(self, stored_window: List[DataStorage], flag: bool):
        """ADD DOCS"""

        cursor_pos_for_outlook = (735, 186)
        search_pattern = self.teams_window_name

        if not flag:
            search_pattern = self.outlook_window_name

        for window in stored_window:
            if window.name and search_pattern in window.name:
                win32gui.ShowWindow(window.handler, win32con.SW_SHOWNOACTIVATE)
                win32gui.SetForegroundWindow(window.handler)
                time.sleep(1)
                win32gui.MoveWindow(window.handler, 365, 100, 1200, 800, win32con.FALSE)
                time.sleep(1)

    @staticmethod
    def left_button_click(dx: int, dy: int):
        """Add DOCS"""

        win32api.SetCursorPos((dx, dy))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.5)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
