#!/usr/bin/env python3
import json
import logging
import openpyxl
import datetime
import time
import win32com.client

MAX_LATENESS_FOR_MEETING = 600
MEETING_EARLINESS = 60

logger = logging.getLogger('MEETING')


def get_meetings_from_excel():
    meetings = []
    wb = openpyxl.load_workbook('meetings.xlsx')
    sheet = wb['Sheet1']

    for i in sheet.iter_rows(values_only=True):
        if i[0] is not None:
            meetings.append(list(i))
    meetings.pop(0)
    logger.info('Collected following meetings from Excel:')
    for meeting in meetings:
        logger.info(meeting)
        # Convert time to timestamp
        meeting[0] = datetime.datetime.strptime(meeting[0], "%d-%m-%Y %H:%M").timestamp()
    return meetings


def get_meetings_from_json():
    with open('meetings.json', ) as f:
        meetings = json.load(f)
        logger.info('Collected following meetings from JSON:')
        for meeting in meetings:
            logger.info(meeting)
            # Convert time to timestamp
            meeting[0] = datetime.datetime.strptime(meeting[0], "%d-%m-%Y %H:%M").timestamp()
        return meetings


def get_meetings_from_outlook():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendar = outlook.GetDefaultFolder(9)
    appointments = sorted(calendar.Items, key=lambda x: x.StartInStartTimeZone.timestamp())
    meetings = []

    for appointment in appointments:
        current_time = round(time.time(), 0)
        if appointment.StartInStartTimeZone.timestamp() - (MAX_LATENESS_FOR_MEETING + 5) > current_time:
            meeting_time = appointment.StartInStartTimeZone.strftime("%d-%m-%Y %H:%M")
            meeting_link = appointment.Location
            meeting_topic = appointment.ConversationTopic
            meetings.append([meeting_time, meeting_link, None, None, None, meeting_topic])

    logger.info('Collected following meetings from Outlook:')
    for meeting in meetings:
        logger.info(meeting)
        # Convert time to timestamp
        meeting[0] = datetime.datetime.strptime(meeting[0], "%d-%m-%Y %H:%M").timestamp()
    return meetings


def get_meetings():
    meetings = get_meetings_from_outlook() + get_meetings_from_excel() + get_meetings_from_json()
    meetings.sort(key=lambda x: x[0])
    return meetings


def join_meetings(meetings, automator):
    '''
    :param meetings: List([timestamp, meeting_link, meeting_id, meeting_password, meeting_topic])
    :param automator: automator object implementing 'join_meeting' method
    :return:
    '''
    for i in range(len(meetings)):
        current_meeting = meetings[i]

        # Setting the meeting times
        current_time = round(time.time(), 0)
        meeting_time = current_meeting[0]

        # Join sometime early for later scheduled meeting
        if current_time < meeting_time - MEETING_EARLINESS:
            sleep_duration = meeting_time - current_time - MEETING_EARLINESS
            next_meeting_time = datetime.timedelta(seconds=sleep_duration)
            logger.info('Sleeping till the next meeting, which is in {}.'.format(next_meeting_time))
            time.sleep(sleep_duration)
        # Too much time has passed already
        elif (current_time - meeting_time) > MAX_LATENESS_FOR_MEETING:
            logger.info('Skipped meeting \"{}\" (meeting {}) since more than {} minutes have passed since this '
                        'meeting began '
                  .format(current_meeting[4], i + 1, MAX_LATENESS_FOR_MEETING / 60))
            continue

        automator.join_meeting(meeting_link=current_meeting[1], meeting_id=current_meeting[2],
                               meeting_password=current_meeting[3])

