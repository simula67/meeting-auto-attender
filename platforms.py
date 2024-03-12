import logging
import sys
import os
from urlextract import URLExtract
import time
import validators
from urllib.parse import urlparse
from constants import MAX_LATENESS_FOR_MEETING, MAX_COLLECT_DURATION


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


def get_domain(url):
    network_loc = urlparse(url).netloc.split('.')
    if len(network_loc) < 2:
        return None
    return network_loc[-2]


def search_links_domain(source, search_domain):
    extractor = URLExtract()
    urls = extractor.find_urls(source)
    for url in urls:
        domain = get_domain(url)
        if not domain:
            continue
        if domain == search_domain:
            return url


def search_links_text(source, search_text):
    extractor = URLExtract()
    urls = extractor.find_urls(source)
    for url in urls:
        if search_text in url:
            return url



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
                    if meeting_link.lower() == 'webex' or meeting_link == '':
                        # Correction for Webex
                        meeting_link = search_links_domain(appointment.Body, 'webex')
                    elif 'zoom' in meeting_link.lower():
                        # Correction for Zoom
                        meeting_link = search_links_domain(meeting_link, 'zoom')
                        if not meeting_link:
                            meeting_link = search_links_domain(appointment.Body, 'zoom')
                    elif 'microsoft teams meeting' in meeting_link.lower():
                        meeting_link = search_links_text(appointment.Body, 'https%3A%2F%2Fteams.microsoft.com%2Fl%2Fmeetup-join%2F')
                        if meeting_link is None:
                            meeting_link = search_links_text(appointment.Body,
                                                             'https://teams.microsoft.com/l/meetup-join/')



                meeting_topic = appointment.ConversationTopic
                meetings.append([meeting_time, meeting_link, None, None, meeting_topic])
            appointment = calendar.GetNext()

        return meetings

