# Description

Python program to automatically join the online meetings.
The program picks meetings from meetings.xlsx, meetings.json or directly from Outlook.

Only Zoom is supported for video conferencing at the moment.

This program uses UI automation and is hence subject to errors.
For example, Chrome may prompt user to click additional buttons to open Zoom ('Open these types of links in Zoom app').
Zoom may change its UI, causing this program to stop working.
Use with caution.

# Platforms

Windows, Linux and macOS.

# Input format

The input should be in the given format:

Time : dd-mm-yyyy hh:mm (24-hour format)

Meeting URL: https://us02web.zoom.us/j/85071211231231231 (string)

Meeting ID : 123456123 (string)

Meeting Password : 1234 (string)

Comment: Example meeting (string)

Please refer the example in meetings.xlsx or meetings.json.

These example meetings should get skipped when the program is run since they started a long time ago.
There is no need to remove them from these files, as long as MAX_LATENESS_FOR_MEETING is not changed in meeting.py.

# Modules used

pyautogui - https://pyautogui.readthedocs.io/en/latest/

openpyxl - https://openpyxl.readthedocs.io/en/stable/

PIL - https://pillow.readthedocs.io/en/stable/

PyWin32 - https://github.com/mhammond/pywin32

# Pre-Requirements

1. Zoom app
2. Web browser (chrome, firefox. Make sure it has pop-up enabled to open Zoom app)
3. Python - Download and install from https://www.python.org/downloads/

# Steps to use

1. Open command prompt and type following command (installing modules - Pyautogui, Openpyxl, Pillow) (this is required only for the first time)

pip install -r requirements.txt

2.Optionally, open meetings.xlsx and enter the schedule of the day in the Excel sheet in the correct columns in the correct format.

Time : dd-mm-yyyy hh:mm Meeting ID : 123456123 (string)(not required if meeting link is provided) Meeting Password : 1234 (string)(not required if meeting link is provided) Comment: Example meeting (Optional)

Warning : Please enter as given.

Or, again optionally, you can use meetings.json and enter the meeting details there.

The program will combine the meetings it obtained from all sources (meetings.xlsx, meetings.json and Outlook)

3. Run Zoom and log in with your username and password, if you want to join as a particular user.

4. Make sure to close all other windows and free up the desktop.

5. Run main.py.

6. Do not close the command prompt where the program is running and watch for any errors that show up.

Keep an eye out in case of errors and failures.

# Errors

1. Mouse losing control: quickly move the mouse as far up and left as you can
2. Program stuck in infinite loop: in the command prompt spam CTRL + C, re-run main.py to restart
3. Any other errors will show up in the command prompt window if not it ll close re-run main.py to restart
4. Do not let the computer sleep when there is long intervals between meetings

# Future work

Add support for Microsoft Teams and Google Meet.

# Notes

Zoom has functionality to automatically keep [audio](https://support.zoom.us/hc/en-us/articles/203024649-Muting-your-microphone-when-joining-a-meeting) and [video](https://support.zoom.us/hc/en-us/articles/4404456197133-Turning-video-off-when-joining-a-meeting) disabled when joining a meeting.

It is highly advisable to have those enabled and not rely on functionality provided in this app.
This is because this app uses UI automation, and it will be less reliable than Zoom's own settings.

# Credit

Initial version from https://github.com/Kn0wn-Un/Auto-Zoom has been re-written almost entirely.

