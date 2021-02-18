import asyncio
import json
import subprocess
import threading
from os import path
from tempfile import gettempdir
from tkinter import *
from tkinter import ttk

import aiohttp

from files import __location__
from windows_style_button import WindowsButton
from windows_title_bar_button import WindowsTitleBarButton


async def get_update_installer(url):
    filename = url.split("/")[-1]
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers={'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}) as update_download:
            assert update_download.status == 200
            installer = await asyncio.wait_for(update_download.read(), 180)
    with open(path.join(gettempdir(), filename), 'wb') as installer_file:
        installer_file.write(installer)
        print(gettempdir())
        print(filename)
    subprocess.Popen([path.join(gettempdir(),filename), "/SILENT"])

async def get_update_info():
    async with aiohttp.ClientSession() as session:
        async with session.get("https://jcedmiston.me/s/update.json", headers={'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}) as update_info:
            assert update_info.status == 200
            return await asyncio.wait_for(update_info.json(), 60)

class Updater:
    def __init__(self, base):
        self.base = base
        self.update_popup = Toplevel(base, takefocus=True)
        self.update_popup.grab_set()
        
        self.update_popup.update_idletasks()
        x = base.winfo_rootx()
        y = base.winfo_rooty()
        x_offset = base.winfo_screenwidth()  / 2 - self.update_popup.winfo_width() / 2
        y_offset = base. winfo_screenheight() / 3 - self.update_popup.winfo_height() / 2
        geom = "+%d+%d" % (x_offset,y_offset)
        self.update_popup.wm_geometry(geom)
        self.update_popup.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        #self.update_popup.wm_title("Checking for updates...")
        self.update_popup.overrideredirect(True)
        self.update_popup.resizable(False, False)
        self.update_popup.rowconfigure(0,weight=1)
        self.update_popup.columnconfigure(0,weight=1)
        self.update_popup.configure(bg='gray15')

        self.title_bar = Frame(self.update_popup, bg="gray20")
        self.close_button = WindowsTitleBarButton(self.title_bar, width=45, height=28, command=self.on_closing)
        self.close_button.pack(side=RIGHT)
        self.win_title = Label(self.title_bar, bg='gray20', fg='white', padx=4, bd=0, text="CSV 2 Paper")
        self.win_title.pack(side=LEFT)
        self.title_bar.grid(row=0, column=0, columnspan=100, sticky='nsew')
        

        # bind title bar motion to the move window function
        #self.title_bar.bind('<B1-Motion>', self.move_window)
        self.title_bar.bind('<Button-1>', self.get_pos)

		#self.running_description = StringVar(value="Mapping data to fields...")
		#self.running_description_label = Label(self.update_popup, textvariable=self.running_description, justify=LEFT)
		#self.running_description_label.grid(row=1, column=0, pady=(10,0), padx=5, sticky=W)
        s = ttk.Style()
        s.theme_use('alt')
        s.configure('blue.Horizontal.TProgressbar', troughcolor  = 'gray35', troughrelief = 'flat', background = '#2f92ff')
        self.progress_indeterminate = ttk.Progressbar(self.update_popup, style = 'blue.Horizontal.TProgressbar', orient="horizontal",length=300, mode="indeterminate")
        self.progress_indeterminate["maximum"] = 100
        self.progress_indeterminate.grid(row=1, column=0, columnspan=2, pady=5, padx=5, sticky='ew')
        self.progress_indeterminate.start(20)

        self.current_version = None
        with open(path.join(__location__, 'resources','version.json'), 'r') as current_version_source:
            self.current_version = json.load(current_version_source)
        
        #self.label_group = Frame(self.update_popup, bg='white', highlightcolor='white', highlightbackground='white', bd=0, highlightthickness=0)

        self.new_version_avail_text = Label(self.update_popup, bg='gray15', fg='white', bd=0, text="There is a new version available!", justify=CENTER)
        self.current_version_num_label = Label(self.update_popup, bg='gray15', fg='white', bd=0, text="Current Version: "+self.current_version['exec_version'], justify=CENTER)
        self.new_version_num = StringVar()
        self.new_version_num_label = Label(self.update_popup, bg='gray15', fg='white', bd=0, textvariable=self.new_version_num, justify=CENTER)
        
        #self.new_version_avail_text.grid(row=0, column=0, columnspan=2, pady=(5,0), padx=5, sticky='nswe')
        #self.current_version_num_label.grid(row=1, column=0, padx=5, sticky='nswe')
        #self.new_version_num_label.grid(row=1, column=1, padx=15, sticky='nswe')

        self.install_now = BooleanVar()
        #self.install_now_button = Button(self.update_popup, relief=GROOVE, activebackground='#e5f1fb', text="Install Now", command=lambda: self.install_now.set(True))
        #self.maybe_later_button = Button(self.update_popup, relief=GROOVE, text="Maybe Later", command=lambda: self.install_now.set(False))
        
        self.install_now_button = WindowsButton(self.update_popup, darkmode=True, highlight=True, text="Install Now", command=lambda: self.install_now.set(True))
        self.maybe_later_button = WindowsButton(self.update_popup, darkmode=True, text="Maybe Later", command=lambda: self.install_now.set(False))

        self.update_available = False
        self.update_installer_url = None
        self.updated_version_num = None
        self.update_popup.update()

        try:
            self.check_thread = threading.Thread(target=self.check_for_updates)
            self.check_thread.start()
            while self.check_thread.is_alive():
                self.update_popup.update()
        except (AssertionError, asyncio.TimeoutError):
            self.on_closing()

        if self.update_available:
            self.progress_indeterminate.destroy()
            self.update_popup.wm_title("New Version")
            self.new_version_avail_text.grid(row=1, column=0, columnspan=2, pady=(5,0), padx=5, sticky='nswe')
            self.current_version_num_label.grid(row=2, column=0, padx=5, sticky='nswe')
            self.new_version_num.set(value = "New Version: "+self.updated_version_num)
            self.new_version_num_label.grid(row=2, column=1, padx=15, sticky='nswe')
            #self.label_group.grid(row=0, column=0, pady=5, columnspan=3, padx=5, sticky='nswe')
            self.install_now_button.grid(row=3, column=0, pady=5, padx=5, sticky='we')
            #self.install_now_button.focus()
            self.maybe_later_button.grid(row=3, column=1, pady=5, padx=5, sticky='we')
            self.update_popup.update()
            self.update_popup.wait_variable(self.install_now)

            if self.install_now.get():
                try:
                    self.download_thread = threading.Thread(target=self.download_installer)
                    self.download_thread.start()
                    while self.download_thread.is_alive():
                        self.update_popup.update()
                    self.base.destroy()
                except (AssertionError, asyncio.TimeoutError):
                    self.on_closing()
            else:
                self.on_closing()
        else:
            self.on_closing()

    def check_for_updates(self):
        try:
            update = asyncio.run(get_update_info())
            if self.current_version['exec_version'] != update['exec_version']:
                self.update_available = True
                self.updated_version_num = update['exec_version']
                self.update_installer_url = update['installer_link']
        except AssertionError:
            self.base.deiconify()
            self.update_popup.destroy()

    def download_installer(self):
        try:
            asyncio.run(get_update_installer(self.update_installer_url))                
        except AssertionError:
            self.base.deiconify()
            self.update_popup.destroy()
    
    def on_closing(self):
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