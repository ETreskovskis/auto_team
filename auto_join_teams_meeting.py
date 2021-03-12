from __future__ import annotations
import re
import sys
import warnings
from dataclasses import dataclass, field, asdict
from typing import Optional, List, Tuple

import pywintypes
import win32com.client
import win32process
import win32gui
import win32api
import win32con
import webbrowser
import time
import datetime

# Todo: add func to start Outlook if it closed
# Todo: what if there are two or more accounts and it has different calendars????
def _for_debugging_purpose(ensure_dispatch):
    """If Dispatch object Outlook.Application is cached then delete the file. This may happen when
    win32com.client.gencache.EnsureDispatch("Outlook.Application") was called
    """
    # ensure_dispatch = win32com.client.gencache.EnsureDispatch("Outlook.Application")

    print(sys.modules[ensure_dispatch.__module__].__file__)


@dataclass(init=False, order=True)
class DataStorage:
    pass


class OutlookApi:

    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.folders = self._enumerate_outlook_folders()
        self.fail_flag = False

    def _enumerate_outlook_folders(self) -> DataStorage:
        """ADD DOCS"""
        folders = DataStorage()

        for num in range(50):

            try:
                folder = self.outlook.GetDefaultFolder(num)
                setattr(folders, folder.Name, num)
            except pywintypes.com_error:
                pass

        return folders

    # Todo: idea is to sort meetings by provided date. At the moment it retrieves 'todays' meetings
    def _sort_calendar_meeting_object(self) -> List:
        """Sort today`s existing meetings from Outlook Calendar"""

        calendar = self.outlook.getDefaultFolder(self.folders.Calendar).Items
        calendar.IncludeRecurrences = True
        calendar.Sort("[Start]")

        # Modify date by needs
        today_date = datetime.datetime.today() + datetime.timedelta(days=1)
        tomorrow_date = datetime.timedelta(days=1) + today_date
        begin_day = today_date.date().strftime("%m/%d/%Y")
        end_day = tomorrow_date.date().strftime("%m/%d/%Y")

        # return Items collection of MeetingItem
        # https://docs.microsoft.com/en-us/office/vba/api/outlook.items.restrict
        # https://docs.microsoft.com/en-us/office/vba/api/outlook.meetingitem
        meeting_plan = calendar.Restrict("[Start] >= '" + begin_day + "' AND [END] <= '" + end_day + "'")

        return meeting_plan

    @staticmethod
    def _populate_meeting_events(event_items: List):
        """Iterate through list of MeetingItem and parse the meeting data"""
        url = " http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/OnlineMeetingConfLink"
        for appointment in event_items:
            p = appointment.ItemProperties
            try:
                for i in range(120):
                    print(p.Item(i))
            except pywintypes.com_error:
                pass
            # print(appointment.Location)
            # print(appointment.Class)
            # print(appointment.Session)
            # print(appointment.StartUTC)
            # print(appointment.UserProperties)
            # print(appointment.Attachments)
            # print(appointment.Application)
            # print(appointment.EntryID)
            # print(appointment.GlobalAppointmentID)
            # print(appointment.IsOnlineMeeting)
            # print(appointment.ItemProperties)
            # print(appointment.Links)
            # print(appointment.MeetingStatus)
            # print(appointment.Recipients)
            # print(appointment.ReminderMinutesBeforeStart)
            # print(appointment.GetConversation())
            # print(appointment.MeetingWorkspaceURL)
            # print(appointment.MeetingWorkspaceURL)
            # print(appointment.IsOnlineMeeting)
            # inspector = appointment.GetInspector
            # print(inspector.Top)
            # print(inspector.CommandBars.LargeButtons)
            event = DataStorage()
            setattr(event, "Start", appointment.Start)
            setattr(event, "End", appointment.End)
            setattr(event, "Subject", appointment.Subject)
            setattr(event, "Duration", appointment.Duration)
            setattr(event, "GetOrganizer", appointment.GetOrganizer().__str__())
            setattr(event, "IsRecurring", appointment.IsRecurring)
            setattr(event, "GetRecurrencePattern", appointment.GetRecurrencePattern().__int__())
            setattr(event, "Body", appointment.Body)
            setattr(event, "Display", appointment.Display)

            yield event

    # Todo: create email body and URL validators
    def _parse_teams_meet_join_url(self, meeting_event: DataStorage) -> Optional[str]:
        """Parse Teams meet-join url from event Body. If body is absent then open Outlook Meeting Occurrence window"""

        if not meeting_event.Body:
            warnings.warn("Outlook calendar meeting email BODY (meet event) was not parsed")
            self.fail_flag = True
            return meeting_event.Display()
        general_url_pattern = re.compile(
            pattern="http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+")
        meeting_join = re.compile(pattern="meetup-join")
        results = re.findall(general_url_pattern, string=meeting_event.Body)

        if not results:
            warnings.warn(f"email URL is missing in email body")
            self.fail_flag = True
            return meeting_event.Display()

        format_result = [url.strip(">") for url in results]
        without_https = [url.strip("https:") for url in format_result]

        meet_url, *_ = [url for url in without_https if re.search(meeting_join, url)]
        return meet_url

    @staticmethod
    def _open_teams_meet_via_url(url: str):
        """ADD DOCS"""

        full_url = f"msteams:{url}"
        webbrowser.open(full_url)


    # Todo: just an idea. Put all together
    def main(self):
        """ADD DOCS"""

        all_meetings = self._sort_calendar_meeting_object()
        parsed_meeting_data = ((meeting.Start, meeting)for meeting in self._populate_meeting_events(all_meetings))
        # sort meetings by time
        sorted_meetings = sorted(parsed_meeting_data)
        # Todo: add wait. Wait for provided meeting time before 5min had passed - join the meeting
        # Todo: continue logic: select TA Scrum meeting, parse url, open url
        for _, meeting in sorted_meetings:
            search_result = self._parse_teams_meet_join_url(meeting)
            if not search_result:
                # print(meeting.__dict__)
                return search_result

            self._open_teams_meet_via_url(search_result)
            # Todo: test only one item
            break


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


class InvokeEvents:
    # Todo: initialize meeting names. Create names as class instances not as attributes
    # Todo: IMPORTANT!!! THERE IS two type: "meeting name - Meeting" "meetinga name - Meeting Occurrence"!!!!!
    test_meeting = ""
    # Todo: parse exact names TA Daily Scrum (Teams window) and TA Daily Scrum - Meeting Occurrence (Outlook window)????
    outlook_window_name = "TA Daily Scrum - Meeting Occurrence"
    teams_window_name = "TA Daily Scrum"
    cursor_for_outlook_ribbon_teams = (735, 186)
    cursor_teams_join_button = (1405, 750)
    # Todo: implement later on how to interact with multiple buttons
    cursor_microphone_on_off = (1092, 524)
    cursor_background_filters = (534, 690)

    def retrieve_current_window_handler(self, stored_window: List[DataStorage], pos: Tuple[int, int],
                                        search_pattern: str):
        """ADD DOCS"""

        for window in stored_window:
            if window.name and search_pattern in window.name:
                win32gui.ShowWindow(window.handler, win32con.SW_SHOWNOACTIVATE)
                win32gui.SetForegroundWindow(window.handler)
                time.sleep(1)
                win32gui.MoveWindow(window.handler, 365, 100, 1200, 800, win32con.FALSE)
                time.sleep(1)
                self.left_button_click(*pos)
                break

    @staticmethod
    def left_button_click(dx: int, dy: int):
        """Simulate mouse left button click on provided position"""

        win32api.SetCursorPos((dx, dy))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.5)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.5)

    def simulate_auto_join(self, stored_window: List[DataStorage], flag: bool):
        """Add DOCS"""

        # if Teams meet-join URL was parsed without failure
        if not flag:
            self.retrieve_current_window_handler(stored_window, search_pattern=self.teams_window_name,
                                                 pos=self.cursor_teams_join_button)
            return

        self.retrieve_current_window_handler(stored_window, search_pattern=self.outlook_window_name,
                                             pos=self.cursor_for_outlook_ribbon_teams)
        self.retrieve_current_window_handler(stored_window, search_pattern=self.teams_window_name,
                                             pos=self.cursor_teams_join_button)


outlook = OutlookApi()
result = outlook.main()
# meeting1, meeting2 = list(meetings)
# print(meeting1)
# print(sorted(meetings))

# enum = EnumActiveWindows()
# data = enum.enumerate_windows
# for d in data:
#     print(d.__dict__)