from __future__ import annotations
import os
import re
import sys
from collections import namedtuple
from dataclasses import dataclass
from typing import Optional, List, Tuple, Callable

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


# Todo: probably remove and use dataclass for storing data
CalendarEvent = namedtuple("CalendarEvent", ["event_start", "subject", "duration", "organizer", "recurrence",
                                             "is_recurring", "body"])


@dataclass(init=False)
class DataStorage:
    pass


class OutlookApi:

    def __init__(self):

        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.folders = self.enumerate_outlook_folders()

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

    @staticmethod
    def parse_teams_meet_join_url(meeting_event: DataStorage) -> Optional[str]:
        """Parse Teams meet-join url from event Body. If body is absent then open Outlook Meeting Occurrence window"""

        if not meeting_event.Body:
            import warnings
            warnings.warn("Outlook calendar event (meeting) was not parsed")
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
