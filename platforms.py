import sys
import os

def get_platform():
    platform_name = sys.platform

    if platform_name == 'win32':
        return Windows()
    elif platform_name == 'linux':
        return Linux()

    raise 'Unknown platform. Detected: {}, supported: Windows and Linux'.format(platform_name)


class Linux:

    def __init__(self):
        self.platform_name = 'Linux'

    def find_zoom_binary(self):
        return '/usr/bin/zoom'

    def close_zoom_process(self):
        os.system('pkill -9 zoom')


class Windows:
    def __init__(self):
        self.platform_name = 'Windows'

    def find_zoom_binary(self):
        sub_folders = [f.path for f in os.scandir('C:\\Users') if f.is_dir()]
        for i in sub_folders:
            if os.path.isfile(i + '\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe'):
                return i + '\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe'

    def close_zoom_process(self):
        os.system("taskkill /f /im Zoom.exe")
