import json
from os import mkdir, path
from files import __location__
from tkinter import *
from detect_dark_mode import is_system_dark
from os.path import isdir, join, expanduser


#app_data_path = join(expanduser("~"), "AppData", "Local", "CSV 2 Paper")
app_data_path = join(expanduser("~"), "Developer", "CSV2PAPER", "CSV2Paper", "Settings")
#/Users/Joseph/Developer/CSV2PAPER/CSV2Paper/settings
if not isdir(app_data_path):
    mkdir(app_data_path)

settings_path = join(app_data_path, "settings")
if not isdir(settings_path):
    mkdir(settings_path)

class UserSettings():
    def __init__(self):
        self.check_for_updates_on_start = BooleanVar(value=True)
        self.default_theme = StringVar(value='system')

        self.saved_settings = None
        
        with open(path.join(settings_path, 'settings.json'), 'r') as settings:
            self.saved_settings = json.load(settings)
            
        self.check_for_updates_on_start.set(self.saved_settings['check_for_updates_on_start'])
        self.default_theme.set(self.saved_settings['default_theme'])
            
        if self.default_theme.get() == 'system':
            self.dark_mode_enabled = is_system_dark()
        elif self.default_theme.get() == 'light':
            self.dark_mode_enabled = False
        else:
            self.dark_mode_enabled = True

    def save_to_disk(self):
        settings_dict = {"check_for_updates_on_start": self.check_for_updates_on_start.get(),
        "default_theme": self.default_theme.get()}
        with open(path.join(settings_path, 'settings.json'), 'w') as settings:
            settings.seek(0)
            json.dump(settings_dict, settings)
            settings.truncate()
    
    def update_dark_mode(self):
        if self.default_theme.get() == 'system':
            self.dark_mode_enabled = is_system_dark()
        elif self.default_theme.get() == 'light':
            self.dark_mode_enabled = False
        else:
            self.dark_mode_enabled = True