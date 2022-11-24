import logging
import time
import webbrowser
from urllib.parse import urlparse


import pyautogui
from PIL import Image
logging.getLogger('PIL').setLevel(logging.WARNING)

# enabling mouse fail safe
pyautogui.FAILSAFE = True

MAX_WAIT_LINK_OPEN = 300

logger = logging.getLogger('AUTOMATOR')


def get_position_from_image(image, timeout=60, confidence=0.9):
    position = None
    seconds_spent = 0
    while position is None:
        if seconds_spent < timeout:
            position = pyautogui.locateOnScreen(image, confidence=confidence)
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

    def join_meeting(self, meeting_link, meeting_id, meeting_password):
        if meeting_link is not None:
            domain = urlparse(meeting_link).netloc.split('.')[-2]
            if domain.lower() == 'webex':
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
        locate_and_click('images/webex_mute.png', timeout=15,confidence=self.confidence)



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
        if pyautogui.locateOnScreen('images/mute.png', confidence=self.confidence) is not None:
            mute_button = pyautogui.locateOnScreen('images/mute.png', confidence=self.confidence)
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
            launch_meeting = pyautogui.locateOnScreen('images/launchmeeting.png', confidence=self.confidence)
            leave_meeting = pyautogui.locateOnScreen('images/leave.png', confidence=self.confidence)
            if leave_meeting is not None:
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
            zoom_app = pyautogui.locateOnScreen('images/final.png', confidence=self.confidence)
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
        video_off = pyautogui.locateOnScreen('images/videooff.png', confidence=self.confidence)
        pyautogui.click(video_off)

        # clicking the join button
        join_meeting_button = pyautogui.locateOnScreen('images/join.png', confidence=self.confidence)
        join_meeting_button = (
            join_meeting_button[0] + 75, join_meeting_button[1] + 10, join_meeting_button[2], join_meeting_button[3])
        pyautogui.moveTo(pyautogui.center(join_meeting_button))
        pyautogui.click(join_meeting_button)

        time.sleep(3)

        # checking and entering if meeting password is enabled
        if pyautogui.locateOnScreen('images/password.png', confidence=self.confidence) is not None:
            pyautogui.typewrite(meeting_password)
            join_meeting_button = pyautogui.locateOnScreen('images/joinmeeting.png', confidence=0.9)
            pyautogui.click(join_meeting_button)

        time.sleep(5)
        # check whether the meeting has started and join with 'enableaudio'
        while True:
            if pyautogui.locateOnScreen('images/audioenable.png', confidence=self.confidence) is not None:
                join_with_audio = pyautogui.locateOnScreen('images/audioenable.png', confidence=0.9)
                pyautogui.click(join_with_audio)
                break
            elif pyautogui.locateOnScreen('images/leave.png', confidence=self.confidence) is not None:
                leave_button = pyautogui.locateOnScreen('images/leave.png', confidence=0.9)
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
