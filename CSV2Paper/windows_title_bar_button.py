from tkinter import *
from tkinter import Button, Frame
from os.path import join

from files import FilePaths, __location__

class WindowsTitleBarButton(Button):
    def __init__(self, master=None, cnf={}, text=None, state='normal', close=True, command=None, image=None, width=None, height=None, **kw):
        self.is_close = close
        if self.is_close:
            self.activebackground='#f0727D'
        else:
            self.activebackground='#666666'
        self.close_icon = PhotoImage(file = join(__location__, 'resources', 'close-18dp', '1x', 'baseline_close_white_18dp.png'))
        super().__init__(master, cnf, bd=0, takefocus=False, relief=FLAT, overrelief=FLAT, width=width, height=height, activebackground=self.activebackground, bg='gray20', fg='white', state=state, command=command, highlightthickness=0, image=self.close_icon, **kw)

        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
    
    def on_enter(self, e):
        if self.is_close:
            self.configure(background = 'red')
        else:
            self.configure(background = '#444444')
    
    def on_leave(self, e):
        self.configure(background='gray20')