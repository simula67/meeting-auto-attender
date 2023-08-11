# Description

Python program to automatically join the online meetings.
The program picks meetings from meetings.xlsx, meetings.json or directly from Outlook
(Picking up meetings from Outlook is supported only on Windows).

Only Zoom, WebEx and MS Teams are supported for video conferencing at the moment.

This program uses UI automation and is hence subject to errors.
For example, Chrome may prompt user to click additional buttons to open Zoom ('Open these types of links in Zoom app').
Zoom may change its UI, causing this program to stop working.

USE WITH CAUTION.

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

1. Zoom, WebEx, MS Teams apps
2. Web browser (Chrome, Firefox etc. Make sure it has pop-up enabled to open Zoom app)
3. Python - Download and install from https://www.python.org/downloads/

# Steps to use

1. Open a command prompt and type following command (installing modules - Pyautogui, Openpyxl, Pillow, pywin32) 

pip install -r requirements.txt

This is required only for the first time.

Installation of pywin32 may fail on non-Windows systems.
This is only required for picking up meetings from Outlook and therefore, this feature is not supported on non-Windows platforms.
The rest of the program should continue to run even without this module.

2. Optionally, open meetings.xlsx and enter the schedule of the day in the Excel sheet in the correct columns in the correct format.

Time : dd-mm-yyyy hh:mm Meeting ID : 123456123 (string)(not required if meeting link is provided) Meeting Password : 1234 (string)(not required if meeting link is provided) Comment: Example meeting (Optional)

Please follow the format that is specified.

Or, again optionally, you can use meetings.json and enter the meeting details there.

The program will combine the meetings it obtained from all sources (meetings.xlsx, meetings.json and Outlook)

3. Run Zoom, WebEx and MS Teams and log in with your username and password, if you want to join as a particular user.

4. Close all other window and have a clean desktop.

5. Run main.py.

6. Do not close the terminal window where the program is running.

Watch the terminal window for error messages.

# Errors

1. Mouse losing control: quickly move the mouse as far up and left as you can
2. Program stuck in infinite loop: in the command prompt spam CTRL + C, re-run main.py to restart
3. Any other errors will show up in the command prompt window if not it ll close re-run main.py to restart
4. Do not let the computer sleep when there is long intervals between meetings

# Future work

Add support for Google Meet.

# Notes

Zoom has functionality to automatically keep [audio](https://support.zoom.us/hc/en-us/articles/203024649-Muting-your-microphone-when-joining-a-meeting) and [video](https://support.zoom.us/hc/en-us/articles/4404456197133-Turning-video-off-when-joining-a-meeting) disabled when joining a meeting.
WebEx has similar [functionality](https://help.webex.com/en-us/article/npg35it/Webex-App-%7C-Choose-the-default-audio-and-video-for-meetings)

It is highly advisable to have those enabled and not rely on functionality provided in this app.
This is because this app uses UI automation, and it will be less reliable than Zoom's or WebEx's own settings.

# Troubleshooting

## Changes in UI elements between machines
It is possible that the images that are provided in this repository are not matching the UI elements on your machine.
This could be because of fonts, themes, etc. being different on your machine.
If the corresponding UI elements are different on your machine, the program might exit with the following error:

`Timeout exceeded while waiting for image <image> to be available. Timeout was <x> seconds`

If you take screenshots of the UI elements presented in the "images" folder and replace them, that should solve this problem.

Unfortunately, it is not that easy to provide all possible versions of these images as part of this project.
You can also play with the confidence value in the `automator.py` to solve this problem.
However, if the confidence is too low, it can cause mis-clicks, therefore, this method is not recommended.

# Credit

Initial version from https://github.com/Kn0wn-Un/Auto-Zoom has been re-written almost entirely.
