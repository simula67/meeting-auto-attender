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

    logger.info("Changing to script directory")
    abspath = os.path.abspath(__file__)
    dir_name = os.path.dirname(abspath)
    os.chdir(dir_name)

    # Mention pre-requisites
    logger.info('Please ensure that you have signed into Zoom')

    # Setup
    logger.info("Setting up platform")
    platform = platforms.get_platform()
    logger.info("Detected platform: {}".format(platform.platform_name))
    automator = automator.Automator(platform=platform)

    # Run
    logger.info("Getting meetings")
    meetings = meeting.get_meetings()
    logger.info("Joining meetings")
    meeting.join_meetings(meetings, automator)

    # Cleanup
    logger.info("Done")
