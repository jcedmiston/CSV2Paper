import json
import subprocess
import threading
import requests
from os import path
from tempfile import gettempdir
from tkinter import *
from tkinter import ttk

from files import __location__
from user_settings import UserSettings
from windows_style_button import WindowsButton
from windows_title_bar_button import WindowsTitleBarButton

class Updater:
    def __init__(self, base: Tk, user_settings: UserSettings, on_start=False):
        self.base = base
        self.user_settings = user_settings
        self.update_popup = Toplevel(base, takefocus=True)
        self.update_popup.grab_set()
        self.headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}
        
        self.window_bg = None
        self.widget_bg = None
        self.fg = None
        self.insert_bg = None
        self.disabled_bg = 'gray80'
        self.select_bg = None
        self.folder_icon_file = None
        self.up_arrow_icon_file = None
        self.down_arrow_icon_file = None
        self.set_colors()

        self.update_popup.update_idletasks()
        x = base.winfo_rootx()
        y = base.winfo_rooty()
        x_offset = base.winfo_screenwidth()  / 2 - self.update_popup.winfo_width() / 2
        y_offset = base. winfo_screenheight() / 3 - self.update_popup.winfo_height() / 2
        geom = "+%d+%d" % (x_offset,y_offset)
        self.update_popup.wm_geometry(geom)
        self.update_popup.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.update_popup.wm_title("CSV 2 Paper")
        
        self.update_popup.resizable(False, False)
        self.update_popup.rowconfigure(0,weight=1)
        self.update_popup.columnconfigure(0,weight=1)
        self.update_popup.configure(bg=self.window_bg)

        s = ttk.Style()
        s.theme_use('alt')
        s.configure('blue.Horizontal.TProgressbar', troughcolor  = 'gray35', troughrelief = 'flat', background = '#2f92ff')
        self.progress_indeterminate = ttk.Progressbar(self.update_popup, style = 'blue.Horizontal.TProgressbar', orient="horizontal",length=225, mode="indeterminate")
        self.progress_indeterminate["maximum"] = 100
        self.progress_indeterminate.grid(row=2, column=0, columnspan=2, pady=5, padx=5, sticky='ew')
        self.progress_indeterminate.start(20)

        self.current_version = None
        with open(path.join(__location__, 'resources','version.json'), 'r') as current_version_source:
            self.current_version = json.load(current_version_source)
        
        self.status_label_text = StringVar(value='Checking for updates...')
        self.status_label = Label(self.update_popup, bg=self.window_bg, fg=self.fg, bd=0, textvariable=self.status_label_text, justify=LEFT)
        self.status_label.grid(row=1, column=0, columnspan=2, pady=(5,0), padx=5, sticky='nswe')
        self.current_version_num_label = Label(self.update_popup, bg=self.window_bg, fg=self.fg, bd=0, text='Current Version: '+self.current_version['exec_version'], justify=CENTER)
        self.new_version_num = StringVar()
        self.new_version_num_label = Label(self.update_popup, bg=self.window_bg, fg=self.fg, bd=0, textvariable=self.new_version_num, justify=CENTER)

        self.install_now = BooleanVar()
        self.ok = BooleanVar()
        
        self.updates_checkbox = Checkbutton(self.update_popup, relief=FLAT, offrelief=FLAT, overrelief=FLAT, bg=self.window_bg, fg=self.fg, activebackground=self.window_bg, activeforeground=self.fg, selectcolor=self.select_bg, text='Check for updates on start', variable=self.user_settings.check_for_updates_on_start, onvalue=True, offvalue=False)
        self.install_now_button = WindowsButton(self.update_popup, darkmode=user_settings.dark_mode_enabled, highlight=True, text="Install Now", command=lambda: self.install_now.set(True))
        self.maybe_later_button = WindowsButton(self.update_popup, darkmode=user_settings.dark_mode_enabled, text="Maybe Later", command=lambda: self.install_now.set(False))
        self.acknowledge_button = WindowsButton(self.update_popup, darkmode=user_settings.dark_mode_enabled, highlight=True, text="Ok", command=lambda: self.ok.set(True))

        self.update_available = False
        self.update_installer_url = None
        self.updated_version_num = None
        self.ready_to_install = False
        self.update_popup.update()

        self.cancel_install = threading.Event()
        self.check_thread = threading.Thread(target=self.check_for_updates)

        self.check_thread.start()
        while self.check_thread.is_alive():
            self.update_popup.update()

        if self.update_available:
            self.progress_indeterminate.grid_remove()
            self.status_label_text.set(value='There is a new version available!')
            self.current_version_num_label.grid(row=2, column=0, padx=5, sticky='nswe')
            self.new_version_num.set(value = "New Version: "+self.updated_version_num)
            self.new_version_num_label.grid(row=2, column=1, padx=15, sticky='nswe')
            self.updates_checkbox.grid(row=3, column=0, columnspan=2, pady=5, padx=5, sticky='we')
            self.install_now_button.grid(row=4, column=0, pady=5, padx=5, sticky='we')
            self.maybe_later_button.grid(row=4, column=1, pady=5, padx=5, sticky='we')
            self.update_popup.update()
            self.update_popup.wait_variable(self.install_now)
            self.current_version_num_label.destroy()
            self.new_version_num_label.destroy()
            self.install_now_button.destroy()
            self.maybe_later_button.destroy()
            self.status_label_text.set(value='Downloading update...')
            self.progress_indeterminate.grid()

            if self.install_now.get():
                self.download_thread = threading.Thread(target=self.get_update_installer, args=(self.update_installer_url, self.cancel_install))
                self.download_thread.start()
                while self.download_thread.is_alive():
                    self.update_popup.update()
                if self.ready_to_install:
                    self.check_thread.join()
                    self.download_thread.join()
                    self.base.destroy()
                else:
                    self.on_closing()
            else:
                self.on_closing()
        else:
            if on_start:
                self.on_closing()
                return
            self.status_label_text.set(value='This version is up to date.')
            self.current_version_num_label.grid(row=2, column=0, padx=5, sticky='nswe')
            self.acknowledge_button.grid(row=3, column=0, padx=5, pady=5, sticky='nswe')
            self.update_popup.wait_variable(self.ok)
            self.on_closing()

    def check_for_updates(self):
        try:
            update_info_download = requests.get('https://jcedmiston.github.io/CSV2Paper/update.json', headers=self.headers)
            assert update_info_download.status_code == 200
            update = update_info_download.json()
        except AssertionError:
            self.update_available = False
        if self.current_version['exec_version'] != update['exec_version']:
            self.update_available = True
            self.updated_version_num = update['exec_version']
            self.update_installer_url = update['installer_link']

    def get_update_installer(self, url, stopped):
        filename = url.split("/")[-1]
        installer_download = requests.get(url, headers=self.headers)
        try:
            assert installer_download.status_code == 200
        except AssertionError: 
            return
        if not stopped.is_set(): 
            installer = installer_download.content
        else: return
        if not stopped.is_set():
            with open(path.join(gettempdir(), filename), 'wb') as installer_file:
                installer_file.write(installer)
        else: return
        if not stopped.is_set(): 
            subprocess.Popen([path.join(gettempdir(),filename), "/SILENT"])
            self.ready_to_install = True
        else: return
    
    def on_closing(self):
        self.cancel_install.set()
        self.base.deiconify()
        self.update_popup.grab_release()
        self.update_popup.destroy()

    def move_window(self, event):
        self.update_popup.geometry('+{0}+{1}'.format(event.x_root, event.y_root))

    def get_pos(self, event):
        xwin = self.update_popup.winfo_x()
        ywin = self.update_popup.winfo_y()
        startx = event.x_root
        starty = event.y_root

        ywin = ywin - starty
        xwin = xwin - startx

        def move_window(event):
            self.update_popup.geometry('+{0}+{1}'.format(event.x_root + xwin, event.y_root + ywin))
        
        startx = event.x_root
        starty = event.y_root

        self.title_bar.bind('<B1-Motion>', move_window)

    def set_colors(self):
        if self.user_settings.dark_mode_enabled:
            self.window_bg = 'gray15'
            self.widget_bg = 'gray35'
            self.fg = 'white'
            self.insert_bg = 'white'
            self.disabled_bg = 'gray20'
            self.select_bg = 'gray30'
        else:
            self.window_bg = 'SystemButtonFace'
            self.widget_bg = 'SystemWindow'
            self.fg = 'SystemWindowText'
            self.insert_bg = 'SystemWindowText'
            self.disabled_bg = 'gray80'
            self.select_bg = 'SystemWindow'
