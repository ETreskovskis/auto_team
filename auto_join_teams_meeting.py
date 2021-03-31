from __future__ import annotations

import _ctypes
import ctypes

import comtypes
import comtypes.client
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


@dataclass()
class SearchPattern:
    """Holds various search patterns for finding window name, button names etc."""

    microphone_re = re.compile(pattern="(?P<mic>[a-zA-Z]ic\s[a-zA-Z]{2,3})")
    camera_re = re.compile(pattern="(?P<camera>[a-zA-Z]amera\s[a-zA-Z]{2,3})")

    subject_unknown = 'New Window | Microsoft Teams'
    subject_name: str = None
    microsoft_teams = re.compile(pattern="Microsoft Teams")
    join_button_patt = "Join With"
    microphone_control_name = "Microphone"
    video_options = "Video options"
    camera_control_name = "Camera"

    def add_name(self, subject: str):
        if subject:
            new_name = "".join([subject, " | Microsoft Teams"])
            self.subject_name = new_name


class OutlookApi:
    """Main class for Outlook API.

    More information about meetings:
    https://docs.microsoft.com/en-us/office/vba/api/outlook.itemproperties
    https://office365itpros.com/2019/10/29/outlook-properties-mark-online-meetings/
    https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-ascal/aa63e887-2e0c-487f-a1a9-d4466708a31b

    MeetingItem info:
    https://docs.microsoft.com/en-us/office/vba/api/outlook.items.restrict
    https://docs.microsoft.com/en-us/office/vba/api/outlook.meetingitem
    """

    def __init__(self, time_before: int = 3 * 60):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.folders = self._enumerate_outlook_folders()
        self.fail_flag = False
        self.start_before = time_before

    def _enumerate_outlook_folders(self) -> DataStorage:
        """Enumerate Outlook folders"""

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
        """Introspect each scheduled event properties and retrieve everything."""

        properties = event.ItemProperties
        event_data = list()

        try:
            for num in range(120):
                event_data.append(properties.Item(num).__str__())
        except pywintypes.com_error:
            pass
        return event_data

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

    def _meeting_time_and_url_mapper(self, meetings: List) -> List[Tuple[float, str, SearchPattern, Any]]:
        """Get meeting time and URL. Map them together."""

        waiting_process = list()
        for meet_start, meeting_object in meetings:
            possible_win_name = SearchPattern()
            possible_win_name.add_name(meeting_object.Subject)
            url_result = self._parse_teams_meet_join_url(meeting_object)
            meeting_time = datetime.datetime(meet_start.year, meet_start.month, meet_start.day, meet_start.hour,
                                             meet_start.minute, meet_start.second)
            waiting_time = meeting_time - datetime.datetime.now()

            waiting_process.append(
                (waiting_time.total_seconds(), url_result, possible_win_name, meeting_object))
        return waiting_process

    def _wait_for_meeting(self, meeting_data: Tuple[float, str, SearchPattern, Any]) -> bool:
        """Wait for meeting. Join the meeting 3 minutes before start"""

        seconds, url, _, meet_object = meeting_data
        if self.fail_flag or not url:
            warnings.warn(
                message=f"Meeting {meet_object.Subject} URL is missing: {url}. Check displayed OutLook window")
            return False

        text = f"Meeting via Teams which start at: {meet_object.Start} - Subject: {meet_object.Subject} " \
               f"- Organizer: {meet_object.GetOrganizer} - Location: {meet_object.Location}"
        print(text)
        time_to_wait = seconds - self.start_before
        time.sleep(time_to_wait)
        return self._open_teams_meet_via_url(url)

    @staticmethod
    def drop_outdated_meetings(meetings: List[Tuple[float, str, SearchPattern, Any]]):
        """Drop outdated meetings when time is negative"""

        for _enum, meeting in enumerate(meetings):
            _time, *_ = meeting
            if _time < 0:
                meetings.pop(_enum)
        return meetings

    def validate_meetings(self, meetings: List):
        """Validate if there is valid meeting list"""

        if not meetings:
            self.fail_flag = True

    def main(self):
        """Main method of Outlook calendar logic."""

        all_meetings = self._sort_calendar_meeting_object()
        parsed_meeting_data = ((meeting.Start, meeting) for meeting in self._populate_meeting_events(all_meetings))
        # sort meetings by time
        sorted_meetings = sorted(parsed_meeting_data)
        waiting_meetings = self._meeting_time_and_url_mapper(sorted_meetings)
        # Output:  List[Tuple[float, str, SearchPattern, Any]]

        # FOR:::ITERATE OVER waiting_process

        # Remove and drop outdated meetings. Validate if there are any
        # current_meetings = self.drop_outdated_meetings(waiting_meetings)
        # return not self.validate_meetings(current_meetings)

        wait_for_meeting = self._wait_for_meeting
        # if wait_for_meeting:
        #     # Todo: initialize IUIAutomation with parser settings (mic on/off, camera on/off)
        #     # Todo: initialize EnumActiveWindows and enumerate_windows
        #     #     Todo: find window by search pattern meeting.Subject
        #     #     Todo: active window with EnumActiveWindows.activate_window
        #     #     Todo: find window and his sublings
        #     # Todo: Get Document control type
        #     # Todo: Get Pane
        #     pass

        return waiting_meetings
        # with ThreadPoolExecutor() as executor:
        #     results = executor.map(self._wait_for_meeting, waiting_process)
        #
        #     for meet_result in results:
        #         print(f"Meeting starts in 3min. Window is open: {meet_result}")

    def _main(self, meeting: Tuple[float, str, SearchPattern, Any], enum: EnumActiveWindows, iui_auto: IUIAutomation):
        """This would be refactored"""
        # Tuple[time_to_start, URL, SearchPattern, DataStorage(with all attributes)]

        time_to_start, url, search_patt, meet_obj = meeting

        # Enumerate active windows
        enumerated = enum.enumerate_windows
        teams_window = enum.validate_teams_open_window(enumerated, search_patt)
        if not teams_window:
            return False

        # Iterate over Teams Window. Get ControlTypes. IUIAutomation block
        from_root_element = iui_auto.child_siblings_from_root_element(iui_auto.raw_view_walker, iui_auto.root_element,
                                                                      enum_wind=teams_window, search_patt=search_patt)
        # Todo: create DocumentControl class == 50030
        get_document_control_list = [element for element in
                                     map(iui_auto.raw_view_walker.GetFirstChildElement, from_root_element) if
                                     element.CurrentControlType == 50030]

        if not get_document_control_list:
            return False

        document_control, *_ = get_document_control_list
        get_controls_50033_list = iui_auto.region_control_siblings_from_document_control(
            walker=iui_auto.control_view_walker,
            element=document_control,
            search_patt=search_patt)

        # first item is Pane (with toolbar Controltype) second Pane(with all other Control types: Audio, volume...)
        if not get_controls_50033_list or len(get_controls_50033_list) < 2:
            return False

        iui_auto.get_microphone_control_type(iui_auto.control_view_walker, get_controls_50033_list, search_patt)

        if not iui_auto.microphone_control:
            return False

        # Get Toolbar and Camera Controls
        tool_bar = iui_auto.get_toolbar_control_type(iui_auto.raw_view_walker, get_controls_50033_list, search_patt)

        if not tool_bar:
            return False

        if not iui_auto.camera_control:
            return False


class EnumActiveWindows:
    """Enumerate windows. Activate windows.

    More about Window parameters and constants:
    https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow
    """

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
        """Retrieve enumerated active windows"""

        win32gui.EnumWindows(self._get_window_info, self.enum_windows)
        return self.enum_windows

    @staticmethod
    def validate_teams_open_window(enumerated: List[DataStorage], search_patt: SearchPattern) -> List[int]:
        """Find open Teams window. Search is based on meeting.Subject name"""

        teams_window = [window.handler for window in enumerated if search_patt.subject_name in window.name]
        if not teams_window:
            teams_window = [window.handler for window in enumerated if
                            search_patt.subject_unknown in window.name]
        return teams_window

    @staticmethod
    def activate_window(window_handler):
        """Retrieve window handler by search pattern. Set window as foreground window and resize it. Perform button
        click"""

        win32gui.ShowWindow(window_handler, win32con.SW_SHOWNOACTIVATE)
        win32gui.SetForegroundWindow(window_handler)
        win32gui.MoveWindow(window_handler, 365, 100, 1200, 800, win32con.FALSE)


class IUIAutomation:
    """ Reference regarding initializing UIAutomationCore, UUID.
    UIAutomationCore: https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-uiautomationoverview
    UUID: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ff384838(v=vs.85)

    Note:
        Can not load UIAutomationCore.dll.\nYou may need to install Windows Update KB971513.
        https://github.com/yinkaisheng/WindowsUpdateKB971513ForIUIAutomation

    Other references regarding ControlType, Property ID, Accessible Role, Accessible States:
    ControlType id`s:
    https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-controltype-ids

    Control Pattern Identifies:
    https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-controlpattern-ids

    Property ID:
    https://docs.microsoft.com/en-us/windows/desktop/WinAuto/uiauto-automation-element-propids
    https://docs.microsoft.com/en-us/windows/desktop/WinAuto/uiauto-control-pattern-propids

    Accessible Role::
    https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.accessiblerole?view=netframework-4.8

    Accessible State:
    https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.accessiblestates?view=netframework-4.8

    """

    # Todo: get preferred camera and mic states from parser
    def __init__(self, camera: str = None, mic: str = None):
        self.__iui_auto_core = comtypes.client.GetModule("UIAutomationCore.dll").IUIAutomation
        self.__uuid = "{ff48dba4-60ef-4201-aa87-54103eef594e}"
        self.iui_automation = comtypes.client.CreateObject(self.__uuid, interface=self.__iui_auto_core)
        self.control_view_walker = self.iui_automation.ControlViewWalker
        self.raw_view_walker = self.iui_automation.RawViewWalker
        self.root_element = self.iui_automation.GetRootElement()
        self.join_button = None
        self.microphone_control = None
        self.camera_control = None
        self.cam_state = None
        self.mic_state = None
        self.preferred_states = camera, mic

    @staticmethod
    def iterate_over_elements(walker, element, max_iteration=0xFFFFFFFF) -> Generator[Any, None, None]:
        """Iterate over IUIAutomation element with raw_view_walker or control_view_walker"""

        child = walker.GetFirstChildElement(element)
        if not child:
            yield None
        depth = 0
        while max_iteration >= depth:
            sibling = walker.GetNextSiblingElement(child)
            if sibling:
                yield sibling
                child = sibling
                depth += 1
            else:
                break

    @staticmethod
    def get_bounding_rectangle(element: Any) -> Tuple[int, int, int, int]:
        """Get bounding rectangle of element"""

        top = element.CurrentBoundingRectangle.top
        bottom = element.CurrentBoundingRectangle.bottom
        left = element.CurrentBoundingRectangle.left
        right = element.CurrentBoundingRectangle.right
        return top, bottom, left, right

    @staticmethod
    def debug_ui_element(element):
        """For debugging purposes"""

        print(40 * "=")
        print(f"Element name: {element.CurrentName}")
        print(f"Current Control Type: {element.CurrentControlType}")
        print(f"Current Native Window Handle: {element.CurrentNativeWindowHandle}")
        print(f"Current Is Control Element: {element.CurrentIsControlElement}")
        print(f"Current Is Controller For: {element.CurrentControllerFor}")

    @property
    def camera_state(self):
        """Camera current state"""
        if not self.join_button:
            warnings.warn(f"Class instance 'join_button' is {self.join_button!r}")
            self.cam_state = "unknown"
            return self.cam_state
        result = re.search(SearchPattern.camera, self.join_button.CurrentName)
        *_, self.cam_state = result.group("camera").split(" ")
        return self.cam_state

    @property
    def microphone_state(self):
        """Microphone current state"""
        if not self.join_button:
            warnings.warn(f"Class instance 'join_button' is {self.join_button!r}")
            self.mic_state = "unknown"
            return self.mic_state
        result = re.search(SearchPattern.microphone_re, self.join_button.CurrentName)
        *_, self.mic_state = result.group("mic").split(" ")
        return self.mic_state

    def region_control_siblings_from_document_control(self, walker, element, search_patt: SearchPattern):
        """Retrieve two region ControlType: 50033"""
        # Todo: create PaneControlType == 50033 class

        siblings_5033 = list()
        for sibling in self.iterate_over_elements(walker, element):
            if sibling.CurrentControlType == 50033:
                siblings_5033.append(sibling)
            if search_patt.join_button_patt in sibling.CurrentName:
                self.join_button = sibling
        return siblings_5033

    def child_siblings_from_root_element(self, walker, root_element, search_patt: SearchPattern, enum_wind: List):
        """Get child siblings from root element (Desktop)"""

        to_search = search_patt.subject_name if search_patt.subject_name else search_patt.subject_unknown
        child_sibling = list()
        for sibling in self.iterate_over_elements(walker, root_element):
            match = re.search(pattern=to_search, string=sibling.CurrentName.__str__())
            if not match:
                return False
            if match and sibling.CurrentNativeWindowHandle in enum_wind:
                child_sibling.append(sibling)
        return child_sibling

    def get_microphone_control_type(self, walker, elements: List, search_patt: SearchPattern):
        """Get microphone ControlType from Pane ControlType"""

        # Todo: create class Microphone ControlType == 50002

        for control in elements:
            for element in self.iterate_over_elements(walker, control):
                if element.CurrentName == search_patt.microphone_control_name:
                    self.microphone_control = element

    # Todo: refactor methods below --> merge to one!!!!

    @staticmethod
    def get_toolbar_control_type(walker, elements: List, search_patt: SearchPattern):
        """Get Toolbar ControlType from Pane ControlType"""

        # Todo: create class Toolbar ControlType == 50021
        get_toolbar_control = [element for element in map(walker.GetFirstChildElement, elements) if (
                    element.CurrentControlType == 50021 and element.CurrentName == search_patt.video_options)]
        return get_toolbar_control

    def get_camera_control_type(self, walker, elements: List, search_patt: SearchPattern):
        """Get Camera ControlType from ToolBar ControlType"""

        # Todo: create class Camera ControlType == 50002
        self.camera_control, *_ = [element for element in map(walker.GetFirstChildElement, elements) if (
                element.CurrentControlType == 50002 and element.CurrentName == search_patt.camera_control_name)]


class MouseEvents:
    """Invoke mouse events"""

    @staticmethod
    def left_button_click(dx: int, dy: int):
        """Simulate mouse left button click on provided position"""

        win32api.SetCursorPos((dx, dy))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.5)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.5)


if __name__ == '__main__':
    from pprint import pprint

    # Todo: add flags: microphone ON/OFF camera ON/OFF. To determine current state of mic and camera, Parse "join" button

    # Todo: check active teams windows before and after session started = investigate difference by pattern
    mock_search = "| Microsoft Teams"

    outlook = OutlookApi()
    get_data = outlook.main()
    print(get_data)
    # List[Tuple[time_to_start, URL, SearchPattern, DataStorage(with all attributes)]]
    print(OutlookApi.__name__ + "=" * 50)
    for time, url, pattern, data_obj in get_data:
        mock_search = pattern.subject_name
        print(mock_search)
    print("=" * 50)

    # Todo: this code is executed after Teams window is displayed!!!
    # Todo: create CLASS wrapper which takes the input and gives output via ThreadPoolExecutor
    # Todo: if _wait_for_meeting is True continue logic below otherwise stop.

    _enum = EnumActiveWindows()
    data = _enum.enumerate_windows
    # teams_all_windows = [(win.name, win.class_name, win.handler) for win in data if mock_search in win.name]
    teams_all_windows_handlers = [win.handler for win in data if mock_search in win.name]
    # pprint(teams_all_windows)
    pprint(teams_all_windows_handlers)
    print("=" * 50)
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

    uiauto_core = comtypes.client.GetModule("UIAutomationCore.dll")
    # https://docs.microsoft.com/en-us/windows/win32/winauto/uiauto-uiautomationoverview
    _iui_auto = uiauto_core.IUIAutomation
    # Reference for UUID https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ff384838(v=vs.85)
    uuid = "{ff48dba4-60ef-4201-aa87-54103eef594e}"

    # Note:!!
    # Can not load UIAutomationCore.dll.\nYou may need to install Windows Update KB971513.
    # \nhttps://github.com/yinkaisheng/WindowsUpdateKB971513ForIUIAutomation'

    iui_automation = comtypes.client.CreateObject(uuid, interface=_iui_auto)
    control_view_walker = iui_automation.ControlViewWalker
    raw_view_walker = iui_automation.RawViewWalker
    root_element = iui_automation.GetRootElement()


    def iterate_over_elements(walker, root_element, max_dep=0xFFFFFFFF):
        child = walker.GetFirstChildElement(root_element)
        if not child:
            yield None
        depth = 0
        while max_dep >= depth:
            subling = walker.GetNextSiblingElement(child)
            if subling:
                yield subling
                child = subling
                depth += 1
            else:
                break


    def get_bounding_rectangle(element: Any) -> Tuple[int, int, int, int]:
        """Get bounding rectangle of element"""

        top = element.CurrentBoundingRectangle.top
        bottom = element.CurrentBoundingRectangle.bottom
        left = element.CurrentBoundingRectangle.left
        right = element.CurrentBoundingRectangle.right
        return top, bottom, left, right


    def debug_ui_element(element):
        """For debugging purposes"""

        print(40 * "=")
        print(f"Element name: {element.CurrentName}")
        print(f"Current Control Type: {element.CurrentControlType}")
        print(f"Current Native Window Handle: {element.CurrentNativeWindowHandle}")
        print(f"Current Is Control Element: {element.CurrentIsControlElement}")
        print(f"Current Is Controller For: {element.CurrentControllerFor}")


    # Todo: get open window.handler (id). DOUBLE check if window is has active flag if not make it visible/display
    #  otherwise "Camera" ControlType would not be found!!!

    # Todo: FIND window then do this logic below!!!!!
    child_sub = list()
    for subling in iterate_over_elements(raw_view_walker, root_element):
        match = re.search(pattern=mock_search, string=subling.CurrentName.__str__())
        if match and subling.CurrentNativeWindowHandle in teams_all_windows_handlers:
            child_sub.append(subling)
            # print(subling.CurrentNativeWindowHandle)

    print("CHILD SIBLING")
    print(child_sub)

    # for subling in iterate_over_elements(raw_view_walker, child_sub[0]):
    #     print(subling.CurrentNativeWindowHandle)
    #     print(subling.CurrentControlType)
    #     print("=" * 60)
    #
    # for subling in iterate_over_elements(raw_view_walker, child_sub[1]):
    #     print(subling.CurrentNativeWindowHandle)
    #     print(subling.CurrentControlType)
    #     print("=" * 60)

    # for subling in child_sub: InvokeEvents.activate_window(subling.CurrentNativeWindowHandle)
    # time.sleep(1)
    #
    # =========================== Get ControlType Document 50030 ==================================
    get_document_control, *_ = [element for element in
                                map(raw_view_walker.GetFirstChildElement, child_sub) if
                                element.CurrentControlType == 50030]
    print("Document ControlType " + 40 * "=")
    print(get_document_control)
    print(40 * "=")

    # print(get_document_control[0].CurrentNativeWindowHandle)
    # print(get_document_control[1].CurrentNativeWindowHandle)
    #
    # # test_this = [child for child in child_sub if
    # #              raw_view_walker.GetFirstChildElement(child).CurrentControlType == 50030]
    # # print(test_this[0].CurrentNativeWindowHandle)
    # # print(test_this[1].CurrentNativeWindowHandle)

    # # # ITERATE over elements of ControlType Document
    # # # first item is Pane (with toolbar Controltype) second Pane(with all other Control types: Audio, volume...)
    control_50033 = list()
    join_button = None
    for item in iterate_over_elements(control_view_walker, get_document_control):
        if item.CurrentControlType == 50033:
            control_50033.append(item)
        if SearchPattern.join_button_patt in item.CurrentName:
            join_button = item

    print(control_50033)
    # print(join_button, join_button.CurrentName)

    # for region in control_50033:
    #     debug_ui_element(region)
    #     print(get_bounding_rectangle(region))
    #     print(40 * "=")

    # WILL NOT work or works and retrieves only 1 subling!!! need to use iterate_over_elements with control_view_walker
    # get_check_box = [element.CurrentName for element in map(control_view_walker.GetFirstChildElement, [control_50033[1]])]
    # print(get_check_box)

    # mic = list()
    # join_button = None
    for control_ in control_50033:

        for item in iterate_over_elements(control_view_walker, control_):
            print(item.CurrentName, item.CurrentControlType, get_bounding_rectangle(item))

    # # # GET TOOLBAR 50021 then get camera access
    print("**" * 100)
    get_toolbar = [element for element in map(raw_view_walker.GetFirstChildElement, control_50033)
                   if element.CurrentControlType == 50021]
    print(get_toolbar[0].CurrentName, get_toolbar[0].CurrentControlType,
          get_bounding_rectangle(get_toolbar[0]))  # SHOULD BE: CurrentName -> 'Video options'

    # # Get Camera ControlType
    get_camera = [camera_re for camera_re in map(control_view_walker.GetFirstChildElement, get_toolbar) if
                      camera_re.CurrentControlType == 50002]
    print(get_camera[0].CurrentControlType, get_camera[0].CurrentName)
    print(get_camera)

    # _print_bouding_rectangle(get_camera)
    #
    # x = (get_camera.CurrentBoundingRectangle.right + get_camera.CurrentBoundingRectangle.left) // 2
    # y = (get_camera.CurrentBoundingRectangle.bottom + get_camera.CurrentBoundingRectangle.top) // 2
    # InvokeEvents.left_button_click(x, y)
