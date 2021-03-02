import os
import re
import sys
from collections import namedtuple
from dataclasses import dataclass
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
class OutlookDataStorage:
    pass


class OutlookApi:

    def __init__(self):

        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.folders = self.enumerate_outlook_folders()

    def enumerate_outlook_folders(self) -> OutlookDataStorage:
        folders = OutlookDataStorage()

        for num in range(50):

            try:
                folder = self.outlook.GetDefaultFolder(num)
                setattr(folders, folder.Name, num)
            except pywintypes.com_error:
                pass

        return folders
