import logging
import sys
import os
from urlextract import URLExtract
import time
import validators
from urllib.parse import urlparse
from meeting import log_collected_meetings, MAX_LATENESS_FOR_MEETING, MAX_COLLECT_DURATION

logger = logging.getLogger('MEETING')

try:
    import win32com.client
except ImportError as e:
    logger.error('Failed to import win32com.client: {}.'.format(e))
    logger.info('Program will continue. It will not be able to pickup meetings from Outlook.')


def get_platform():
    platform_name = sys.platform

    if platform_name == 'win32':
        return Windows()
    elif platform_name == 'linux':
        return Linux()
    elif platform_name == 'darwin':
        return MacOS()

    raise 'Unknown platform. Detected: {}, supported: Windows and Linux'.format(platform_name)


class MacOS:
    def __init__(self):
        self.platform_name = 'MacOS'

    def find_zoom_binary(self):
        return '/Applications/zoom.us.app/Contents/MacOS/zoom.us'

    def close_zoom_process(self):
        os.system('pkill -9 zoom.us')

    def get_meetings_from_outlook(self):
        logger.info('Getting meetings from Outlook is not supported on MacOS')
        return []


class Linux:

    def __init__(self):
        self.platform_name = 'Linux'

    def find_zoom_binary(self):
        return '/usr/bin/zoom'

    def close_zoom_process(self):
        os.system('pkill -9 zoom')

    def get_meetings_from_outlook(self):
        logger.info('Getting meetings from Outlook is not supported on Linux')
        return []


class Windows:
    def __init__(self):
        self.platform_name = 'Windows'

    def find_zoom_binary(self):
        sub_folders = [f.path for f in os.scandir('C:\\Users') if f.is_dir()]
        for i in sub_folders:
            if os.path.isfile(i + '\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe'):
                return i + '\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe'

    def close_zoom_process(self):
        os.system("taskkill /f /im Zoom.exe")

    def get_meetings_from_outlook(self):

        current_time = round(time.time(), 0)
        extractor = URLExtract()

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        calendar = outlook.GetDefaultFolder(9).Items
        calendar.IncludeRecurrences = True
        calendar.Sort("[Start]")
        meetings = []
        appointment = calendar.GetNext()
        # Do not collect meetings beyond MAX_COLLECT_DURATION because recurring meetings can go to infinity.
        while appointment and appointment.StartInStartTimeZone.timestamp() < (current_time + MAX_COLLECT_DURATION):
            if appointment.StartInStartTimeZone.timestamp() - (MAX_LATENESS_FOR_MEETING + 5) > current_time:
                meeting_time = appointment.StartInStartTimeZone.strftime("%d-%m-%Y %H:%M")
                meeting_link = appointment.Location
                if not validators.url(meeting_link):
                    # Meeting link is not a link, attempt correction
                    if meeting_link.lower() == 'webex':
                        # Correction for Webex
                        urls = extractor.find_urls(appointment.Body)
                        for url in urls:
                            domain = urlparse(url).netloc.split('.')[-2]
                            if domain == 'webex':
                                meeting_link = url
                                break

                meeting_topic = appointment.ConversationTopic
                meetings.append([meeting_time, meeting_link, None, None, meeting_topic])
            appointment = calendar.GetNext()

        log_collected_meetings('Outlook', meetings)
        return meetings

