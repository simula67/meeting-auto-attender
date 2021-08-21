# Description

Python program to automatically join the online Zoom meetings based on the given input in the Excel sheet meetings.xlsx

# Platforms

Windows and Linux

# Input format

The input should be in the given format:

Time : dd-mm-yyyy hh:mm AM/PM

Meeting URL: https://us02web.zoom.us/j/85071211231231231 (string)

Meeting ID : 123456123 (string)

Meeting Password : 1234 (string)

Please refer the example in meetings.xlsx

# Modules used

pyautogui - https://pyautogui.readthedocs.io/en/latest/

openpyxl - https://openpyxl.readthedocs.io/en/stable/

PIL - https://pillow.readthedocs.io/en/stable/

# Pre-Requirements

1. Zoom app

2. Web browser (chrome, firefox preferred ,make sure it has pop-up enabled to open zoom app)

3. Python - Download and install from https://www.python.org/downloads/

# Steps to use

1. Open command prompt and type following command (installing modules - pyautogui, openpyxl, Pillow) (this is required only for the first time)

pip install -r requirements.txt

2. Open meetings.xlsx enter the schedule of the day in the Excel sheet in the correct columns in the correct format

Time : dd-mm-yyyy hh:mm AM/PM Meeting ID : 123456123 (string)(not required if meeting link is provided) Meeting Password : 1234 (string)(not required if meeting link is provided)

Warning : Please enter as given.

3. Run Zoom and log in with your username and password, if you want to join as a particular user.

4. Make sure to close all other windows and free up the desktop.

5. Run main.py.

6. Do not close the command prompt where the program is running and watch for any errors that show up.

Keep an eye out in case of errors and failures.

# Errors

1. Mouse losing control : quickly move the mouse as far up and left as you can
2. Program stuck in infinite loop : in the command prompt spam CTRL + C, re-run main.py to restart
3. Any other errors will show up in the command prompt window if not it ll close re-run main.py to restart
4. Do not let the computer sleep when there is long intervals between meetings


# Note

This is not original work. It is a copied and refactored version from https://github.com/Kn0wn-Un/Auto-Zoom

