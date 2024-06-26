import logging
import time
import webbrowser
from urllib.parse import urlparse
from constants import MAX_WAIT_LINK_OPEN

import pyautogui
from PIL import Image
logging.getLogger('PIL').setLevel(logging.WARNING)

# enabling mouse fail safe
pyautogui.FAILSAFE = True


logger = logging.getLogger('AUTOMATOR')


def locate_on_screen(image, confidence):
    try:
        return pyautogui.locateOnScreen(image, confidence=confidence)
    except Exception as e:
        logger.error('Image ({}) not found. Confidence: {}'.format(image, confidence))
        return None


def get_position_from_image(image, timeout=60, confidence=0.9):
    position = None
    seconds_spent = 0
    while position is None:
        if seconds_spent < timeout:
            position = locate_on_screen(image, confidence=confidence)
            if position is None:
                logger.info('Waiting for image \'{}\' to be available. '
                            'Timeout is {} seconds, time spent is: {} seconds'
                            .format(image, timeout, seconds_spent))
                time.sleep(1)
            else:
                return position
            seconds_spent = seconds_spent + 1
        else:
            logger.info('Timeout exceeded while waiting for image \'{}\' to be available. Timeout was {} seconds'\
                        .format(image, timeout))
            return


def locate_and_click(image, pre_click_delay=0, timeout=60, confidence=0.9):
    position = get_position_from_image(image, timeout, confidence=confidence)
    time.sleep(pre_click_delay)
    pyautogui.click(position)


class Automator:
    def __init__(self, platform):
        self.platform = platform
        self.zoom_automator = ZoomAutomator(platform)
        self.webex_automator = WebExAutomator(platform)
        self.msteams_automator = MSTeamsAutomator(platform)

    def join_meeting(self, meeting_link, meeting_id, meeting_password):
        if meeting_link is not None:
            url_components = urlparse(meeting_link).netloc.split('.')
            if len(url_components) < 2:
                raise Exception('Meeting link invalid: {}'.format(meeting_link))
            else:
                domain = url_components[-2]

            if 'teams.microsoft.com' in meeting_link:
                self.msteams_automator.join_meeting_with_link(meeting_link)
            elif domain.lower() == 'webex':
                self.webex_automator.join_meeting_with_link(meeting_link)
            elif domain.lower() == 'zoom':
                self.zoom_automator.join_meeting_with_link(meeting_link)
            else:
                # Open browser link for unknown
                webbrowser.open(meeting_link)
        else:
            if meeting_id is not None and meeting_password is not None:
                # Only zoom supported for joining with meeting id and password
                self.zoom_automator.join_meeting_with_id(meeting_id, meeting_password)
            else:
                raise Exception('Cannot join meeting because of insufficient meeting parameters: '
                                'meeting link: {}, meeting id: {}, meeting password: {}'
                                .format(meeting_link, meeting_id, meeting_password))


class WebExAutomator:
    def __init__(self, platform):
        self.platform = platform
        self.confidence = 0.9

    def join_meeting_with_link(self, meeting_link):
        webbrowser.open(meeting_link)
        locate_and_click('images/webex_join_meeting.png', confidence=self.confidence, pre_click_delay=5)
        locate_and_click('images/webex_mute.png', timeout=15, confidence=self.confidence)


class MSTeamsAutomator:
    def __init__(self, platform):
        self.platform = platform
        self.confidence = 0.9

    def join_meeting_with_link(self, meeting_link):
        webbrowser.open(meeting_link)
        if get_position_from_image('images/msteams_mute.png', timeout=5, confidence=self.confidence):
            locate_and_click('images/msteams_mute.png', timeout=15, confidence=self.confidence)
            locate_and_click('images/msteams_join_meeting.png', confidence=self.confidence)
        else:
            locate_and_click('images/msteams_mute_laptop.png', timeout=15, confidence=self.confidence)
            locate_and_click('images/msteams_join_meeting_laptop.png', timeout=15, confidence=self.confidence)


class ZoomAutomator:
    def __init__(self, platform):
        self.platform = platform
        self.bin_path = platform.find_zoom_binary()

        # Default confidence is 0.9
        self.confidence = 0.9
        if self.platform.platform_name == 'Linux':
            # The fonts could be different for Linux, so reduce the confidence a bit
            self.confidence = 0.8
        elif self.platform.platform_name == 'MacOS':
            self.confidence = 0.5

    def mute_mic(self):
        # check whether the mic is muted, if not mute it
        pyautogui.moveTo(x=900, y=900, duration=0.25)
        if locate_on_screen('images/mute.png', confidence=self.confidence) is not None:
            mute_button = locate_on_screen('images/mute.png', confidence=self.confidence)
            pyautogui.click(mute_button)

    def join_meeting_with_link(self, meeting_link):
        meeting_link_without_hash_sucess = meeting_link
        if meeting_link.endswith('#success'):
            logger.info('Removing \'#success\' from meeting link')
            meeting_link_without_hash_sucess = meeting_link[:-len('#success')]

        logger.info('Joining the meeting with link: {}'.format(meeting_link_without_hash_sucess))
        # open the given link in web browser
        webbrowser.open(meeting_link_without_hash_sucess)
        start = time.time()
        time.sleep(3)
        meeting_joined = False
        while not meeting_joined:
            logger.info('Checking if Zoom was opened')
            launch_meeting = locate_on_screen('images/launchmeeting.png', confidence=self.confidence)
            leave_meeting = locate_on_screen('images/leave.png', confidence=self.confidence)
            end_meeting = locate_on_screen('images/end.png', confidence=self.confidence)
            if leave_meeting is not None or end_meeting is not None:
                logger.info('Joined meeting')
                meeting_joined = True
            elif launch_meeting is not None:
                logger.info('Clicking on Launch Meeting')
                pyautogui.click(launch_meeting)
            elif (time.time() - start) >= MAX_WAIT_LINK_OPEN:
                logger.info("Link " + meeting_link + " not opened")
            time.sleep(3)
        self.mute_mic()

    def join_meeting_with_id(self, meeting_id, meeting_password):
        cur = round(time.time(), 0)
        time.sleep(3)
        # locating the Zoom app
        while True:
            zoom_app = locate_on_screen('images/final.png', confidence=self.confidence)
            if zoom_app is not None:
                pyautogui.click(zoom_app)
                break
            elif (time.time() - cur) >= 120:
                logger.info("App Not opened")
                break
            # check every 30 secs
            time.sleep(30)

        time.sleep(3)

        # entering the meeting id
        pyautogui.typewrite(meeting_id)

        # disabling video source
        video_off = locate_on_screen('images/videooff.png', confidence=self.confidence)
        pyautogui.click(video_off)

        # clicking the join button
        join_meeting_button = locate_on_screen('images/join.png', confidence=self.confidence)
        join_meeting_button = (
            join_meeting_button[0] + 75, join_meeting_button[1] + 10, join_meeting_button[2], join_meeting_button[3])
        pyautogui.moveTo(pyautogui.center(join_meeting_button))
        pyautogui.click(join_meeting_button)

        time.sleep(3)

        # checking and entering if meeting password is enabled
        if locate_on_screen('images/password.png', confidence=self.confidence) is not None:
            pyautogui.typewrite(meeting_password)
            join_meeting_button = locate_on_screen('images/joinmeeting.png', confidence=0.9)
            pyautogui.click(join_meeting_button)

        time.sleep(5)
        # check whether the meeting has started and join with 'enableaudio'
        while True:
            if locate_on_screen('images/audioenable.png', confidence=self.confidence) is not None:
                join_with_audio = locate_on_screen('images/audioenable.png', confidence=0.9)
                pyautogui.click(join_with_audio)
                break
            elif locate_on_screen('images/leave.png', confidence=self.confidence) is not None:
                leave_button = locate_on_screen('images/leave.png', confidence=0.9)
                pyautogui.click(leave_button)
                break
            elif (time.time() - cur) >= 30 * 60:
                self.platform.close_zoom_process()
                break
            time.sleep(5)
        self.mute_mic()

    def join_meeting(self, meeting_link, meeting_id, meeting_password):
        if meeting_link is not None:
            self.join_meeting_with_link(meeting_link)
        else:
            self.join_meeting_with_id(meeting_id, meeting_password)
