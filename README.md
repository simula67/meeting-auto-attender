# Description

Python program to automatically join the online zoom meetings based on the given input in the Excel sheet meetings.xlsx

# Input format

The input should be in the given format:

Time : dd-mm-yyyy hh:mm AM/PM

Meeting ID : 123456123 (string)

Meeting Password : 1234 (string)


# Modules used

pyautogui - https://pyautogui.readthedocs.io/en/latest/

openpyxl - https://openpyxl.readthedocs.io/en/stable/

PIL - https://pillow.readthedocs.io/en/stable/

# Pre-Requirements

Windows machine

Zoom app

Web browser (chrome, firefox preferred ,make sure it has pop-up enabled to open zoom app)

Python - Download and install from https://www.python.org/downloads/

# Steps to use

1. Open command prompt and type (installing modules - pyautogui, openpyxl, Pillow) (this is required only for the first time)

pip install -r requirements.txt

2. open meetings.xlsx enter the schedule of the day in the excel sheet in the correct columns in the correct format

Time : dd-mm-yyyy hh:mm AM/PM Meeting ID : 123456123 (string)(not required if meeting link is provided) Meeting Password : 1234 (string)(not required if meeting link is provided)

Warning : please enter as given

3. Make sure to close all windows and free up the desktop

4. run main.py

5. Do not close the command prompt thats where the program is running and any errors show up

Keep an eye out in case of errors and failures

# Errors

1. Mouse losing control : quickly move the mouse as far up and left as you can
2. program stuck in infinite loop : in the command prompt spam CTRL + C, re-run main.py to restart
3. any other errors will show up in the command prompt window if not it ll close re-run main.py to restart
4. do not let the computer sleep when there is long intervals between meetings


# Note

This is not original work. It is a copied version from https://github.com/Kn0wn-Un/Auto-Zoom

