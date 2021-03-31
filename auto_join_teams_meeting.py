from __future__ import annotations

import argparse
from functools import partial

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

    subject_name: str = None
    subject_unknown = 'New Window | Microsoft Teams'
    microsoft_teams = re.compile(pattern="Microsoft Teams")
    join_button_patt = "Join With"
    microphone_control_name = "Microphone"
    video_options = "Video options"
    camera_control_name = "Camera"

    def add_name(self, subject: str):
        if subject:
            new_name = "".join([subject, " | Microsoft Teams"])
            self.subject_name = new_name


@dataclass(init=False)
class ControlType:
    PaneControlType: int = 50033
    CheckBoxControlType: int = 50002
    ToolBarControlType: int = 50021
    DocumentControlType: int = 50030


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

    def wait_for_meeting(self, meeting_data: Tuple[float, str, SearchPattern, Any]) -> bool:
        """Wait for meeting. Join the meeting 3 minutes before start"""

        seconds, url, _, meet_object = meeting_data
        if self.fail_flag or not url:
            warnings.warn(
                message=f"Meeting {meet_object.Subject} URL is missing: {url}. Check displayed OutLook window")
            return False

        text = f"Meeting via Teams which starts at: {meet_object.Start} - Subject: {meet_object.Subject} " \
               f"- Organizer: {meet_object.GetOrganizer} - Location: {meet_object.Location}"
        print(text)
        time_to_wait = seconds - self.start_before
        time.sleep(time_to_wait)
        return self._open_teams_meet_via_url(url)

    @staticmethod
    def drop_outdated_meetings(meetings: List[Tuple[float, str, SearchPattern, Any]]) -> List[
        Tuple[float, str, SearchPattern, Any]]:
        """Drop outdated meetings when time is negative"""

        for _enum, meeting in enumerate(meetings):
            _time, *_ = meeting
            if _time < 0:
                meetings.pop(_enum)
        return meetings

    def available_meetings(self):
        """Main method of Outlook calendar logic."""

        all_meetings = self._sort_calendar_meeting_object()
        parsed_meeting_data = ((meeting.Start, meeting) for meeting in self._populate_meeting_events(all_meetings))
        # sort meetings by time
        sorted_meetings = sorted(parsed_meeting_data)
        waiting_meetings = self._meeting_time_and_url_mapper(sorted_meetings)

        # Remove and drop outdated meetings.
        current_meetings = self.drop_outdated_meetings(waiting_meetings)
        return current_meetings


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
        # setattr(window_info, "stock", win32gui.GetStockObject(hwnd))
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
        """Retrieve window handler by search pattern. Set window as foreground window."""

        win32gui.ShowWindow(window_handler, win32con.SW_SHOWNOACTIVATE)
        win32gui.SetForegroundWindow(window_handler)


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

    def __init__(self, camera: str, mic: str):
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
        self.preferred_cam_state = camera.lower() if isinstance(camera, str) and camera else ''
        self.preferred_mic_state = mic.lower() if isinstance(mic, str) and mic else ''

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
    def debug_ui_element(element):
        """For debugging purposes"""

        print(40 * "=")
        print(f"Element name: {element.CurrentName}")
        print(f"Current Control Type: {element.CurrentControlType}")
        print(f"Current Native Window Handle: {element.CurrentNativeWindowHandle}")
        print(f"Current Is Control Element: {element.CurrentIsControlElement}")
        print(f"Current Is Controller For: {element.CurrentControllerFor}")

    @property
    def change_camera_state(self) -> bool:
        """Change camera current state ->> preferred state"""

        current_state = self.camera_state
        pref_state = self.preferred_cam_state
        if current_state != pref_state:
            return True
        return False

    @property
    def change_mic_state(self) -> bool:
        """Change microphone current state ->> preferred state"""

        current_state = self.microphone_state
        pref_state = self.preferred_mic_state
        if current_state != pref_state:
            return True
        return False

    @property
    def camera_state(self) -> str:
        """Camera current state"""
        if not self.join_button:
            warnings.warn(f"Class instance 'join_button' is {self.join_button!r}")
            self.cam_state = "unknown"
            return self.cam_state
        result = re.search(SearchPattern.camera_re, self.join_button.CurrentName)
        *_, self.cam_state = result.group("camera").split(" ")
        return self.cam_state.lower()

    @property
    def microphone_state(self):
        """Microphone current state"""
        if not self.join_button:
            warnings.warn(f"Class instance 'join_button' is {self.join_button!r}")
            self.mic_state = "unknown"
            return self.mic_state
        result = re.search(SearchPattern.microphone_re, self.join_button.CurrentName)
        *_, self.mic_state = result.group("mic").split(" ")
        return self.mic_state.lower()

    @property
    def get_camera_x_y(self) -> Tuple[int, int]:
        """Get Camera ControlType x, y to press"""

        camera = self.camera_control
        x = (camera.CurrentBoundingRectangle.right + camera.CurrentBoundingRectangle.left) // 2
        y = (camera.CurrentBoundingRectangle.bottom + camera.CurrentBoundingRectangle.top) // 2
        return x, y

    @property
    def get_mic_x_y(self) -> Tuple[int, int]:
        """Get Microphone ControlType x, y to press"""

        mic = self.microphone_control
        x = (mic.CurrentBoundingRectangle.right + mic.CurrentBoundingRectangle.left) // 2
        y = (mic.CurrentBoundingRectangle.bottom + mic.CurrentBoundingRectangle.top) // 2
        return x, y

    @property
    def get_join_x_y(self) -> Tuple[int, int]:
        """Get Join ControlType x, y to press"""

        join = self.join_button
        x = (join.CurrentBoundingRectangle.right + join.CurrentBoundingRectangle.left) // 2
        y = (join.CurrentBoundingRectangle.bottom + join.CurrentBoundingRectangle.top) // 2
        return x, y

    def region_control_siblings_from_document_control(self, walker, element, search_patt: SearchPattern):
        """Retrieve two Pane ControlType: 50033"""

        siblings_5033 = list()
        for sibling in self.iterate_over_elements(walker, element):
            if sibling.CurrentControlType == ControlType.PaneControlType:
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

        for control in elements:
            for element in self.iterate_over_elements(walker, control):
                if element.CurrentName == search_patt.microphone_control_name and (
                        element.CurrentControlType == ControlType.CheckBoxControlType):
                    self.microphone_control = element

    @staticmethod
    def get_toolbar_control_type(walker, elements: List, search_patt: SearchPattern):
        """Get Toolbar ControlType from Pane ControlType"""

        get_toolbar_control = [element for element in map(walker.GetFirstChildElement, elements) if (
                element.CurrentControlType == ControlType.ToolBarControlType and (
                element.CurrentName == search_patt.video_options))]
        return get_toolbar_control

    def get_camera_control_type(self, walker, elements: List, search_patt: SearchPattern):
        """Get Camera ControlType from ToolBar ControlType"""

        self.camera_control, *_ = [element for element in map(walker.GetFirstChildElement, elements) if (
                element.CurrentControlType == ControlType.CheckBoxControlType and (
                element.CurrentName == search_patt.camera_control_name))]


class MouseEvents:
    """Invoke mouse events"""

    @classmethod
    def left_button_click(cls, dx: int, dy: int):
        """Simulate mouse left button click on provided position"""

        win32api.SetCursorPos((dx, dy))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.5)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.5)


class TeamsRunner:

    def __init__(self):
        pass

    @staticmethod
    def validate_meetings(meeting_data: List):
        """Validate if there is valid meeting list"""

        if not meeting_data:
            return False
        return True

    @staticmethod
    def validate_mic_camera_join_controls(mic, cam, jbutton):
        """Validate if all necessary buttons are parsed"""

        if not mic or not cam or not jbutton:
            return False
        return True

    @staticmethod
    def main(meeting: Tuple[float, str, SearchPattern, Any], enum: EnumActiveWindows, iui_auto: IUIAutomation,
             outlook: OutlookApi, mouse: MouseEvents) -> Tuple[bool, Tuple]:
        """This would be refactored"""
        # Tuple[time_to_start, URL, SearchPattern, DataStorage(with all attributes)]

        wait_for_meeting = outlook.wait_for_meeting(meeting_data=meeting)
        if not wait_for_meeting:
            return False, meeting

        time_to_start, url, search_patt, meet_obj = meeting

        # Enumerate active windows
        enumerated = enum.enumerate_windows
        teams_window = enum.validate_teams_open_window(enumerated, search_patt)
        if not teams_window:
            return False, meeting

        # Activate window. Set window as foreground window.
        teams_window_hwnd = teams_window[-1]
        enum.activate_window(teams_window_hwnd)

        # Iterate over Teams Window. Get ControlTypes. IUIAutomation block.
        from_root_element = iui_auto.child_siblings_from_root_element(iui_auto.raw_view_walker, iui_auto.root_element,
                                                                      enum_wind=teams_window, search_patt=search_patt)
        get_document_control_list = [element for element in
                                     map(iui_auto.raw_view_walker.GetFirstChildElement, from_root_element) if
                                     element.CurrentControlType == ControlType.DocumentControlType]

        if not get_document_control_list:
            return False, meeting

        # Get Pane ControlTypes and get join button
        document_control, *_ = get_document_control_list
        get_controls_50033_list = iui_auto.region_control_siblings_from_document_control(
            walker=iui_auto.control_view_walker,
            element=document_control,
            search_patt=search_patt)

        # first item is Pane (with toolbar Controltype) second Pane(with all other Control types: Audio, volume...)
        if not get_controls_50033_list or len(get_controls_50033_list) < 2:
            return False, meeting

        iui_auto.get_microphone_control_type(iui_auto.control_view_walker, get_controls_50033_list, search_patt)

        # Get Toolbar and Camera Controls
        tool_bar = iui_auto.get_toolbar_control_type(iui_auto.raw_view_walker, get_controls_50033_list, search_patt)
        if not tool_bar:
            return False, meeting

        iui_auto.get_camera_control_type(iui_auto.control_view_walker, tool_bar, search_patt)

        # Verify ControlTypes: camera, microphone, join button are parsed
        if not TeamsRunner.validate_mic_camera_join_controls(mic=iui_auto.microphone_control,
                                                             cam=iui_auto.camera_control,
                                                             jbutton=iui_auto.join_button):
            return False, meeting

        # Enable microphone, camera, join button coordinates
        camera = iui_auto.get_camera_x_y
        mic = iui_auto.get_mic_x_y
        join_button = iui_auto.get_join_x_y

        # Check if Camera and Microphone should be changed their state
        if iui_auto.change_camera_state:
            mouse.left_button_click(*camera)
        if iui_auto.change_mic_state:
            mouse.left_button_click(*mic)

        # Press JOIN button:
        mouse.left_button_click(*join_button)
        return True, meeting

    @classmethod
    def run_meetings(cls, meetings_data: List[Tuple[float, str, SearchPattern, Any]], enum: EnumActiveWindows,
                     iui_auto: IUIAutomation, outlook: OutlookApi, mouse: MouseEvents) -> bool:
        """Validate meetings first and then run them."""

        if not TeamsRunner.validate_meetings(meetings_data):
            return False

        wrapper_main = partial(TeamsRunner.main, enum=enum, iui_auto=iui_auto, outlook=outlook, mouse=mouse)

        with ThreadPoolExecutor() as executor:
            results = executor.map(wrapper_main, meetings_data)

            for mt_result, mt_obj in results:
                print(f"Meeting organized by: {mt_obj[3].GetOrganizer} "
                      f"subject: {mt_obj[3].Subject}. Successful: {mt_result}")
        return True


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Teams AUTO-JOIN. For additional parameter info use --help")
    parser.add_argument("--mic_state", type=str, required=False,
                        help="Provide flag for microphone: 'on' or 'off'. Note: this set up for all upcoming meetings",
                        default="on")
    parser.add_argument("--camera_state", type=str, required=False,
                        help="Provide flag for camera: 'on' or 'off'. Note: this set up for all upcoming meetings",
                        default="on")
    parser.add_argument("--start_before", type=int, required=False,
                        help="Provide time (seconds) to join before actual meeting has started",
                        default=3 * 60)

    arguments = parser.parse_args()

    outlook_class = OutlookApi(time_before=arguments.start_before)
    planned_meetings = outlook_class.available_meetings()
    iui_auto_class = IUIAutomation(camera=arguments.camera_state, mic=arguments.mic_state)
    enum_class = EnumActiveWindows()
    mouse_event = MouseEvents()
    run_meetings = TeamsRunner.run_meetings(planned_meetings, enum=enum_class, iui_auto=iui_auto_class,
                                            outlook=outlook_class, mouse=mouse_event)
    if not run_meetings:
        sys.exit("There are no meetings to start. Quiting.")
    sys.exit(f"Success status: {run_meetings}")
