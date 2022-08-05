#!/usr/bin/env python3

import logging
import os

import platforms
import automator
import meeting

logging.basicConfig(format='%(asctime)s %(levelname)-8s %(message)s', level=logging.DEBUG, datefmt='%Y-%m-%d %H:%M:%S')
logger = logging.getLogger('MAIN')
#logger.addHandler(logging.StreamHandler(sys.stdout))


if __name__ == '__main__':

    # Change to script directory
    abspath = os.path.abspath(__file__)
    dir_name = os.path.dirname(abspath)
    os.chdir(dir_name)

    # Mention pre-requisites
    logger.info('Please ensure that you have signed into Zoom')

    # Setup
    platform = platforms.get_platform()
    zoom_automator = automator.ZoomAutomator(platform=platform)

    # Run
    meetings = meeting.get_meetings()
    meeting.join_meetings(meetings, zoom_automator)

    # Cleanup
    logger.info("Done")
