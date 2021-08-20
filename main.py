import datetime
import os
import time

import platforms
import zoom_automator
import openpyxl

MAX_LATENESS_FOR_MEETING = 300
MEETING_EARLINESS = 60


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
        meeting_time = datetime.datetime.strptime(current_meeting[0], "%d-%m-%Y %H:%M %p").timestamp()

        # Join sometime early for later scheduled meeting
        if current_time < meeting_time - MEETING_EARLINESS:
            print("Next meeting in ", end="")
            print(datetime.timedelta(seconds=(meeting_time - current_time) - MEETING_EARLINESS))
            time.sleep(meeting_time - current_time - MEETING_EARLINESS)
        # Too much time has passed already
        elif (current_time - meeting_time) > MAX_LATENESS_FOR_MEETING:
            print('Skipped meeting {} since more than {} minutes have passed since this meeting began'
                  .format(i + 1, MAX_LATENESS_FOR_MEETING / 60))
            continue

        zoomautomator.join_meeting(meeting_link=current_meeting[1], meeting_id=current_meeting[2],
                                   meeting_password=current_meeting[3])


if __name__ == '__main__':
    # Mention pre-requisites
    print('Please ensure that you have signed into Zoom')

    # Setup
    platform = platforms.get_platform()
    zoomautomator = zoom_automator.ZoomAutomator(platform=platform)

    # Run
    meetings = get_meetings()
    join_meetings(meetings, zoomautomator)

    # Cleanup
    print("Done")
    platform.close_zoom_process()
