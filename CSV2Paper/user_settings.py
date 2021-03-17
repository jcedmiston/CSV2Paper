import json
from os import path
from files import __location__
from tkinter import *
from detect_dark_mode import is_system_dark

class UserSettings():
    def __init__(self):
        self.check_for_updates_on_start = BooleanVar(value=True)
        self.default_theme = StringVar(value='system')

        self.saved_settings = None
        with open(path.join(__location__, 'settings', 'settings.json'), 'r') as settings:
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
        with open(path.join(__location__, 'settings', 'settings.json'), 'w') as settings:
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