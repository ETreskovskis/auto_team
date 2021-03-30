from __future__ import annotations

import _ctypes
import datetime
import re
import sys
import time
import warnings
import webbrowser
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
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
            meeting_event.Display()
            return meet_url

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

            waiting_process.append(
                (waiting_time.total_seconds(), search_result, meeting_object))
        return waiting_process

    def _wait_for_meeting(self, meeting_data: Tuple[int, str, Any]) -> bool:
        """Wait for meeting. Join the meeting 3 minutes before start"""

        seconds, url, meet_object = meeting_data
        if self.fail_flag or not url:
            warnings.warn(
                message=f"Meeting {meet_object.Subject} URL is missing: {url}. Check displayed OutLook window")
            return False

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

        # Todo add more logic of joining the meeting, changing mic and camera HERE?
        # Todo: split steps

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
        # setattr(window_info, "context", win32gui.GetDC(hwnd))
        # print(win32gui.GetStockObject(hwnd))
        enum_windows.append(window_info)

    @property
    def enumerate_windows(self):
        """ADD DOCS"""

        win32gui.EnumWindows(self._get_window_info, self.enum_windows)
        return self.enum_windows


class InvokeEvents:
    # Todo: initialize meeting names. Create names as class instances not as attributes
    """Control Type Identifiers:
    https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-controltype-ids


    Control Pattern Identifiers:
    https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-controlpattern-ids

    Property ID:
    Refer https://docs.microsoft.com/en-us/windows/desktop/WinAuto/uiauto-automation-element-propids
    Refer https://docs.microsoft.com/en-us/windows/desktop/WinAuto/uiauto-control-pattern-propids

    Accessible Role:
    https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.accessiblerole?view=netframework-4.8

    Accessible State:
    https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.accessiblestates?view=netframework-4.8
    """

    # More about Window parameters: https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow

    @staticmethod
    def activate_window(window_handler):
        """Retrieve window handler by search pattern. Set window as foreground window and resize it. Perform button
        click"""

        win32gui.ShowWindow(window_handler, win32con.SW_SHOWNOACTIVATE)
        win32gui.SetForegroundWindow(window_handler)
        win32gui.MoveWindow(window_handler, 365, 100, 1200, 800, win32con.FALSE)

    @staticmethod
    def left_button_click(dx: int, dy: int):
        """Simulate mouse left button click on provided position"""

        win32api.SetCursorPos((dx, dy))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.5)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.5)


@dataclass()
class SearchPattern:
    subject_unknown = 'New Window | Microsoft Teams'
    subject_known: list = field(default_factory=list)

    def add_name(self, subject):
        if subject:
            new_name = "".join([subject, " | Microsoft Teams"])
            self.subject_known.append(new_name)


if __name__ == '__main__':
    from pprint import pprint

    # Todo: experiment
    search_pattern = SearchPattern()
    search_pattern.add_name("TEST window")
    print(search_pattern.subject_known)
    # outlook = OutlookApi()
    # outlook.main()

    # Todo: How to identify correct Teams active window???
    # Todo: check active teams windows before and after session started = investigate difference by pattern
    # enum = EnumActiveWindows()
    # data = enum.enumerate_windows
    # teams_all_windows = [(win.name, win.class_name, win.handler) for win in data if search in win.name]
    # teams_all_windows_handlers = [win.handler for win in data if search in win.name]
    # # teams_all_windows_handlers = [(win.handler, win) for win in data if search in win.name]
    # pprint(teams_all_windows)
    #
    # # one is a channel, second is "joined-call"
    # win = [('Microsoft Teams Notification', 'Chrome_WidgetWin_1', 13110044),
    #        ('TA Daily Scrum | Microsoft Teams', 'Chrome_WidgetWin_1', 1248714),
    #        ('TA Daily Scrum | Microsoft Teams', 'Chrome_WidgetWin_1', 199116)]
    #
    # wins = [('Microsoft Teams Notification', 'Chrome_WidgetWin_1', 788316),
    #         ('New Window | Microsoft Teams', 'Chrome_WidgetWin_1', 11143342),
    #         ('Bertasius Ugnius | Microsoft Teams', 'Chrome_WidgetWin_1', 263748)]
    #
    # search_when_subject_unknown = "New Window | Microsoft Teams"
    # search_when_subject_known = "TA Daily Scrum | Microsoft Teams"
    #
    # import comtypes
    # import comtypes.client
    #
    # uiauto_core = comtypes.client.GetModule("UIAutomationCore.dll")
    # # https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-uiautomationoverview
    # _iui_auto = uiauto_core.IUIAutomation
    # # print(dir(_iui_auto))
    # # Reference for UUID https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ff384838(v=vs.85)
    # uuid = "{ff48dba4-60ef-4201-aa87-54103eef594e}"
    #
    # # Note:!!
    # # Can not load UIAutomationCore.dll.\nYou may need to install Windows Update KB971513.
    # # \nhttps://github.com/yinkaisheng/WindowsUpdateKB971513ForIUIAutomation'
    #
    # iui_automation = comtypes.client.CreateObject(uuid, interface=_iui_auto)
    # control_view_walker = iui_automation.ControlViewWalker
    # raw_view_walker = iui_automation.RawViewWalker
    # root_element = iui_automation.GetRootElement()
    #
    #
    # def iterate_over_elements(walker, root_element, max_dep=0xFFFFFFFF):
    #     child = walker.GetFirstChildElement(root_element)
    #     if not child:
    #         yield None
    #     depth = 0
    #     while max_dep >= depth:
    #         subling = walker.GetNextSiblingElement(child)
    #         if subling:
    #             yield subling
    #             child = subling
    #             depth += 1
    #         else:
    #             break
    #
    #
    # def print_ui_element_info(element):
    #     print(40 * "=")
    #     print(f"Element name: {element.CurrentName}")
    #     print(f"Current Control Type: {element.CurrentControlType}")
    #     print(f"Current Native Window Handle: {element.CurrentNativeWindowHandle}")
    #     print(f"Current Is Control Element: {element.CurrentIsControlElement}")
    #     print(f"Current Is Controller For: {element.CurrentControllerFor}")
    #
    #
    # def _print_bouding_rectangle(pointer_item):
    #     print(f"Left: {pointer_item.CurrentBoundingRectangle.left}")
    #     print(f"Top: {pointer_item.CurrentBoundingRectangle.top}")
    #     print(f"Right: {pointer_item.CurrentBoundingRectangle.right}")
    #     print(f"Bottom: {pointer_item.CurrentBoundingRectangle.bottom}")
    #
    #
    # import re
    #
    # # Todo: get open window.handler (id). DOUBLE check if window is has active flag if not make it visible/display
    # #  otherwise "Camera" ControlType would not be found!!!
    # child_sub = list()
    # for subling in iterate_over_elements(raw_view_walker, root_element):
    #     match = re.search(pattern="Microsoft Teams", string=subling.CurrentName.__str__())
    #     if match and subling.CurrentNativeWindowHandle in teams_all_windows_handlers:
    #         child_sub.append(subling)
    #         # print(subling.CurrentNativeWindowHandle)
    #
    # # for subling in iterate_over_elements(raw_view_walker, child_sub[0]):
    # #     print(subling.CurrentNativeWindowHandle)
    # #     print(subling.CurrentControlType)
    # #     print("=" * 60)
    # #
    # # for subling in iterate_over_elements(raw_view_walker, child_sub[1]):
    # #     print(subling.CurrentNativeWindowHandle)
    # #     print(subling.CurrentControlType)
    # #     print("=" * 60)
    #
    # for subling in child_sub: InvokeEvents.activate_window(subling.CurrentNativeWindowHandle)
    # time.sleep(1)
    #
    # # =========================== Get ControlType Document 50030 ==================================
    # get_document_control = [element for element in
    #                         map(raw_view_walker.GetFirstChildElement, child_sub) if
    #                         element.CurrentControlType == 50030]
    # print(get_document_control)
    # # print(get_document_control[0].CurrentNativeWindowHandle)
    # # print(get_document_control[1].CurrentNativeWindowHandle)
    #
    # # test_this = [child for child in child_sub if
    # #              raw_view_walker.GetFirstChildElement(child).CurrentControlType == 50030]
    # # print(test_this[0].CurrentNativeWindowHandle)
    # # print(test_this[1].CurrentNativeWindowHandle)
    #
    # # Todo: DOUBLE check if window is has active flag if not make it visible/display
    # #  otherwise "Camera" ControlType would not be found!!!
    #
    # # # ITERATE over elements of ControlType Document
    # # # first item is Pane (with toolbar Controltype) second Pane(with all other Control types: Audio, volume...)
    # control_50033 = list()
    # join_button = list()
    # for item in iterate_over_elements(control_view_walker, get_document_control[0]):
    #     if item.CurrentControlType == 50033:
    #         control_50033.append(item)
    #     if "Join" in item.CurrentName:
    #         join_button.append(item)
    #
    # print(control_50033)
    # print(join_button)
    # #
    # # # # GET TOOLBAR 50021 then get camera access
    # # # print("**" * 100)
    # # get_toolbar = [element for element in map(raw_view_walker.GetFirstChildElement, control_50033)
    # #                if element.CurrentControlType == 50021]
    # #
    # # # # Get Camera ControlType
    # # get_camera, *_ = [camera for camera in map(control_view_walker.GetFirstChildElement, get_toolbar) if
    # #                   camera.CurrentControlType == 50002]
    # # print(get_camera.CurrentControlType, get_camera.CurrentName)
    # # _print_bouding_rectangle(get_camera)
    # #
    # # x = (get_camera.CurrentBoundingRectangle.right + get_camera.CurrentBoundingRectangle.left) // 2
    # # y = (get_camera.CurrentBoundingRectangle.bottom + get_camera.CurrentBoundingRectangle.top) // 2
    # # InvokeEvents.left_button_click(x, y)
