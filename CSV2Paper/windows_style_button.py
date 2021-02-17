from tkinter import *
from tkinter import Button, Frame


class WindowsButton(Frame, Button):
    def __init__(self, master=None, cnf={}, text=None, state='normal', command=None, image=None, dark_mode=True, width=None, height=None, **kw):
        self.dark_mode = dark_mode
        if dark_mode:
            super().__init__(master, cnf, borderwidth=1, bg='gray15', **kw)
            self.dynamicFrame = Frame(self, borderwidth=1, background="#666666")
            self.button = Button(self.dynamicFrame, bd=0, relief=FLAT, overrelief=FLAT, width=width, height=height, activebackground='#666666', bg='#444444', fg='white', text=text, state=state, command=command, highlightthickness=0, image=image)
        else:
            super().__init__(master, cnf, borderwidth=1, **kw)
            self.dynamicFrame = Frame(self, borderwidth=1, background="#adadad")
            self.button = Button(self.dynamicFrame, bd=0, relief=FLAT, overrelief=FLAT, width=width, height=height, activebackground='#d0e3f5', text=text, state=state, command=command, highlightthickness=0, image=image)
        

        self.dynamicFrame.pack(expand=True,fill='both')
        self.button.pack(expand=True,fill='both')
        self.button.bind("<Enter>", self.on_enter)
        self.button.bind("<Leave>", self.on_leave)
        self.button.bind('<FocusIn>', self.focus_in)
        self.button.bind('<FocusOut>', self.focus_out)
    
    def on_enter(self, e):
        if self.dark_mode:
            if self.button['state'] == 'normal':
                self.button.configure(background = '#666666')
                self.dynamicFrame.configure(borderwidth = 1)
                self.configure(borderwidth = 1)
                self.dynamicFrame.configure(background = '#7b7b7b')
        else:
            if self.button['state'] == 'normal':
                self.button.configure(background = '#e7f1fa')
                self.dynamicFrame.configure(borderwidth = 1)
                self.configure(borderwidth = 1)
                self.dynamicFrame.configure(background = '#3078d0')
    
    def on_leave(self, e):
        if self.dark_mode:
            if self.button['state'] == 'normal':
                self.button.configure(background = '#444444')
                focus = self.master.focus_get() 
                if self.button == focus:
                    self.dynamicFrame.configure(borderwidth = 2)
                    self.configure(borderwidth = 0)
                else:
                    self.dynamicFrame.configure(borderwidth = 1)
                    self.configure(borderwidth = 1)
                    self.dynamicFrame.configure(background = '#adadad')
        else:
            if self.button['state'] == 'normal':
                self.button.configure(background = 'SystemButtonFace')
                focus = self.master.focus_get() 
                if self.button == focus:
                    self.dynamicFrame.configure(borderwidth = 2)
                    self.configure(borderwidth = 0)
                else:
                    self.dynamicFrame.configure(borderwidth = 1)
                    self.configure(borderwidth = 1)
                    self.dynamicFrame.configure(background = '#adadad')

    def configure(self, cnf=None, state=None, **kw):
        if state:
            if state == 'normal':
                pass
            self.button.configure(state=state)
        else:
            return super().configure(cnf, **kw)

    def focus_set(self):
        self.button.focus()
        self.dynamicFrame.configure(borderwidth = 2)
        self.dynamicFrame.configure(background = '#3078d0')
        self.configure(borderwidth = 0)
    focus = focus_set

    def focus_in(self, e):
        self.button.configure(default='active')
        self.dynamicFrame.configure(borderwidth = 2)
        self.dynamicFrame.configure(background = '#3078d0')
        self.configure(borderwidth = 0)
    
    def focus_out(self, e):
        self.button.configure(default='normal')
        self.dynamicFrame.configure(borderwidth = 1)
        self.dynamicFrame.configure(background = '#adadad')
        self.configure(borderwidth = 1)
