import os

def get_platform():
    platform = Windows()

    if os.name != 'nt':
        raise Exception('Only Windows supported at the moment and this appears not be Windows')

    return platform

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
