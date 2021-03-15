from __future__ import annotations

import datetime
import re
import sys
import time
import warnings
import webbrowser
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from typing import Optional, List, Tuple, Generator, Any

import pywintypes
import win32api
import win32com.client
import win32con
import win32gui
import win32process


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

    @staticmethod
    def _get_event_item_properties(event) -> List[str]:
        """Introspect each scheduled event properties and retrieve everything
        https://docs.microsoft.com/en-us/office/vba/api/outlook.itemproperties
        https://office365itpros.com/2019/10/29/outlook-properties-mark-online-meetings/
        https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-ascal/aa63e887-2e0c-487f-a1a9-d4466708a31b
        """

        properties = event.ItemProperties
        event_data = list()

        try:
            for num in range(120):
                event_data.append(properties.Item(num).__str__())
        except pywintypes.com_error:
            pass
        return event_data

    # Todo: idea is to sort meetings by provided date. At the moment it retrieves 'todays' meetings
    def _sort_calendar_meeting_object(self) -> List:
        """Sort today`s existing meetings from Outlook Calendar"""

        calendar = self.outlook.getDefaultFolder(self.folders.Calendar).Items
        calendar.IncludeRecurrences = True
        calendar.Sort("[Start]")

        # Modify date by needs
        today_date = datetime.datetime.today()
        tomorrow_date = datetime.timedelta(days=1) + today_date
        begin_day = today_date.date().strftime("%m/%d/%Y")
        end_day = tomorrow_date.date().strftime("%m/%d/%Y")

        # return Items collection of MeetingItem
        # https://docs.microsoft.com/en-us/office/vba/api/outlook.items.restrict
        # https://docs.microsoft.com/en-us/office/vba/api/outlook.meetingitem
        meeting_plan = calendar.Restrict("[Start] >= '" + begin_day + "' AND [END] <= '" + end_day + "'")

        return meeting_plan

    def _populate_meeting_events(self, event_items: List) -> Generator[DataStorage, None, None]:
        """Iterate through list of MeetingItem and parse the meeting data"""
        for appointment in event_items:
            appointment_properties = self._get_event_item_properties(appointment)
            event = DataStorage()
            setattr(event, "Start", appointment.Start)
            setattr(event, "End", appointment.End)
            setattr(event, "Subject", appointment.Subject)
            setattr(event, "Duration", appointment.Duration)
            setattr(event, "Location", appointment.Location)
            setattr(event, "GetOrganizer", appointment.GetOrganizer().__str__())
            setattr(event, "IsRecurring", appointment.IsRecurring)
            setattr(event, "GetRecurrencePattern", appointment.GetRecurrencePattern().__int__())
            setattr(event, "Body", appointment.Body)
            setattr(event, "Display", appointment.Display)
            setattr(event, "Properties", appointment_properties)

            yield event

    def _parse_teams_meet_join_url(self, meeting_event: DataStorage) -> Optional[str]:
        """Parse Teams meet-join url from event Properties. If URL is absent then open Outlook Meeting Occurrence window
        """

        meet_properties = meeting_event.Properties
        meet_url = None

        general_url_pattern = re.compile(
            pattern="http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+")
        meeting_join = re.compile(pattern="meetup-join")
        for items in meet_properties:
            result = re.findall(general_url_pattern, string=items)
            if result:
                format_result = [url.strip(">") for url in result]
                removed_https_prefix = [url.strip("https:") for url in format_result]
                meet_url = [url for url in removed_https_prefix if re.search(meeting_join, url)]

                if meet_url:
                    return meet_url[0]

        if not meet_url:
            warnings.warn("Meeting URL ir missing!")
            self.fail_flag = True
            return meeting_event.Display()

    @staticmethod
    def _open_teams_meet_via_url(url: str) -> bool:
        """Open Teams via URL"""

        try:
            full_url = f"msteams:{url}"
            return webbrowser.open(full_url)
        except Exception as error:
            msg_error, *_ = error.args
            print(msg_error)

    def _meeting_time_and_url_mapper(self, meetings: List) -> List[Tuple[float, str, Any]]:
        """Get meeting time and URL. Map them together."""

        waiting_process = list()
        for meet_start, meeting_object in meetings:
            search_result = self._parse_teams_meet_join_url(meeting_object)
            meeting_time = datetime.datetime(meet_start.year, meet_start.month, meet_start.day, meet_start.hour,
                                             meet_start.minute, meet_start.second)
            waiting_time = meeting_time - datetime.datetime.now()

            waiting_process.append((waiting_time.total_seconds(), search_result, meeting_object))
        return waiting_process

    def _wait_for_meeting(self, meeting_data: Tuple[int, str, Any]) -> bool:
        """Wait for meeting. Join the meeting 3 minutes before start"""

        seconds, url, meet_object = meeting_data
        text = f"Meeting via Teams which start at: {meet_object.Start} - Subject: {meet_object.Subject} " \
               f"- Organizer: {meet_object.GetOrganizer} - Location: {meet_object.Location}"
        print(text)
        time_to_wait = seconds - 3 * 60
        time.sleep(time_to_wait)
        return self._open_teams_meet_via_url(url)

    def main(self):
        """Main method of Outlook calendar logic."""

        all_meetings = self._sort_calendar_meeting_object()
        parsed_meeting_data = ((meeting.Start, meeting) for meeting in self._populate_meeting_events(all_meetings))
        # sort meetings by time
        sorted_meetings = sorted(parsed_meeting_data)
        waiting_process = self._meeting_time_and_url_mapper(sorted_meetings)

        with ThreadPoolExecutor() as executor:
            results = executor.map(self._wait_for_meeting, waiting_process)

            for meet_result in results:
                print(f"Meeting starts in 3min. Window is open: {meet_result}")


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
        setattr(window_info, "class_name", win32gui.GetClassName(hwnd))
        enum_windows.append(window_info)

    @property
    def enumerate_windows(self):
        """ADD DOCS"""

        win32gui.EnumWindows(self._get_window_info, self.enum_windows)
        return self.enum_windows


class InvokeEvents:
    # Todo: initialize meeting names. Create names as class instances not as attributes
    teams_window_name = "TA Daily Scrum"
    micro_teams = "Microsoft Teams"
    cursor_for_outlook_ribbon_teams = (735, 186)
    cursor_teams_join_button = (1405, 750)
    # Todo: implement later on how to interact with multiple buttons
    cursor_microphone_on_off = (1092, 524)
    cursor_background_filters = (534, 690)

    # Todo: refactor method. add flags which buttons should be disabled. remove filtering EnumActiveWindows should handle this
    def retrieve_current_window_handler(self, stored_window: List[DataStorage], pos: Tuple[int, int],
                                        search_pattern: str):
        """Retrieve window handler by search pattern. Set window as foreground window and resize it. Perform button
        click"""

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

    # Todo: refactor method.
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


if __name__ == '__main__':
    from pprint import pprint

    # outlook = OutlookApi()
    # outlook.main()

    # Todo: How to identify correct Teams active window???
    # Todo: check active teams windows before and after session started = investigate difference by pattern
    search = InvokeEvents.micro_teams
    enum = EnumActiveWindows()
    data = enum.enumerate_windows
    before_teams = [(win.name, win.class_name, win.handler) for win in data if search in win.name]
    # before_set = set(before_teams)
    pprint(before_teams)
    # pprint(before_set)

    # one is a channel, second is "joined-call"
    win = [('Microsoft Teams Notification', 'Chrome_WidgetWin_1', 13110044),
           ('TA Daily Scrum | Microsoft Teams', 'Chrome_WidgetWin_1', 1248714),
           ('TA Daily Scrum | Microsoft Teams', 'Chrome_WidgetWin_1', 199116)]

    # outlook = OutlookApi()
    # outlook.main()

    # new_data = enum.enumerate_windows
    # after_teams = [win.name for win in new_data if search in win.name]
    # after_set = set(after_teams)
    # pprint(after_teams)
    # pprint(after_set)
    #
    # difference = after_set.difference(before_set)
    # print(difference)
