import logging
import time
import webbrowser

import pyautogui
from PIL import Image

# enabling mouse fail safe
pyautogui.FAILSAFE = True

MAX_WAIT_LINK_OPEN = 300

logger = logging.getLogger('AUTOMATOR')


class ZoomAutomator:

    def __init__(self, platform):
        self.platform = platform
        self.bin_path = platform.find_zoom_binary()

        # Default confidence is 0.9
        self.confidence = 0.9
        if self.platform.platform_name == 'Linux':
            # The fonts could be different for Linux, so reduce the confidence a bit
            self.confidence = 0.8

    def join_meeting_with_link(self, meeting_link):
        logger.info('Joining the meeting with link: {}'.format(meeting_link))
        # open the given link in web browser
        webbrowser.open(meeting_link)
        start = time.time()
        time.sleep(3)
        while True:
            logger.info('Checking if Zoom was opened')
            open_link = pyautogui.locateOnScreen('images/openlink.png', confidence=self.confidence)
            open_zoom = pyautogui.locateOnScreen('images/openzoom.png', confidence=self.confidence)
            if open_link is not None:
                logger.info('Clicking on Launch Meeting')
                pyautogui.click(open_link)
            elif open_zoom is not None:
                logger.info('Joined meeting')
                # pyautogui.click(open_zoom)
            elif (time.time() - start) >= MAX_WAIT_LINK_OPEN:
                logger.info("Link " + meeting_link + " not opened")
            time.sleep(3)

    def join_meeting_with_id(self, meeting_id, meeting_password):
        cur = round(time.time(), 0)
        time.sleep(3)
        # locating the zoom app
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

    def join_meeting(self, meeting_link, meeting_id, meeting_password):
        if meeting_link is not None:
            self.join_meeting_with_link(meeting_link)
        else:
            self.join_meeting_with_id(meeting_id, meeting_password)

        # check whether the mic is muted, if not mute it
        pyautogui.moveTo(x=900, y=900, duration=0.25)
        if pyautogui.locateOnScreen('images/mute.png', confidence=self.confidence) is not None:
            mute_button = pyautogui.locateOnScreen('images/mute.png', confidence=self.confidence)
            pyautogui.click(mute_button)
