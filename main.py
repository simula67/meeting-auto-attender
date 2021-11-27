#!/usr/bin/env python3

import datetime
import logging
import sys
import time

import platforms
import automator
import openpyxl

MAX_LATENESS_FOR_MEETING = 600
MEETING_EARLINESS = 60

logging.basicConfig(format='%(asctime)s %(levelname)-8s %(message)s', level=logging.DEBUG, datefmt='%Y-%m-%d %H:%M:%S')
logger = logging.getLogger('MAIN')
#logger.addHandler(logging.StreamHandler(sys.stdout))



def get_meetings():
    # copying data from excel sheet to the program
    zoom_meetings = []
    wb = openpyxl.load_workbook('meetings.xlsx')
    sheet = wb['Sheet1']

    for i in sheet.iter_rows(values_only=True):
        if i[0] is not None:
            zoom_meetings.append(i)
    zoom_meetings.pop(0)
    zoom_meetings.sort()
    return zoom_meetings


def join_meetings(zoom_meetings, zoomautomator):
    for i in range(len(zoom_meetings)):
        current_meeting = zoom_meetings[i]

        # Setting the meeting times
        current_time = round(time.time(), 0)
        meeting_time = datetime.datetime.strptime(current_meeting[0], "%d-%m-%Y %H:%M").timestamp()

        # Join sometime early for later scheduled meeting
        if current_time < meeting_time - MEETING_EARLINESS:
            sleep_duration = meeting_time - current_time - MEETING_EARLINESS
            next_meeting_time = datetime.timedelta(seconds=sleep_duration)
            logger.info('Sleeping till the next meeting, which is in {}.'.format(next_meeting_time))
            time.sleep(sleep_duration)
        # Too much time has passed already
        elif (current_time - meeting_time) > MAX_LATENESS_FOR_MEETING:
            logger.info('Skipped meeting {} since more than {} minutes have passed since this meeting began'
                  .format(i + 1, MAX_LATENESS_FOR_MEETING / 60))
            continue

        zoomautomator.join_meeting(meeting_link=current_meeting[1], meeting_id=current_meeting[2],
                                   meeting_password=current_meeting[3])


if __name__ == '__main__':

    # Mention pre-requisites
    logger.info('Please ensure that you have signed into Zoom')

    # Setup
    platform = platforms.get_platform()
    zoom_automator = automator.ZoomAutomator(platform=platform)

    # Run
    meetings = get_meetings()
    join_meetings(meetings, zoom_automator)

    # Cleanup
    logger.info("Done")
