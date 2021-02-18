from tkinter import *
from tkinter import Button, Frame

class WindowsButton(Frame):
    def __init__(self, master, cnf=None, anchor=None, command=None, compound=LEFT, 
                default='normal', justify=CENTER, state='normal', takefocus=True, name=None, height=None, 
                image_filename=None, subx=None, suby=None, text=None, textvariable=None, underline=-1, 
                width=None, wraplength=None, padx=None, pady=None, darkmode=False,
                highlight=False):
        
        self._no_border = 0
        self._normal_border = 1
        self._expanded_border = 2
        self._surroundings_dark = 'gray15'
        self._activebackground_light = '#d0e3f5'
        self._activebackground_dark = '#808080'
        self._highlight_light = '#e7f1fa'
        self._highlight_dark = '#6a6a6a'
        self._bg_light = 'SystemButtonFace'
        self._bg_dark = '#444444'
        self._fg_dark = 'white'  
        self._highlight_border_light = '#3078d0'
        self._highlight_border_dark = '#7b7b7b'
        self._border_light = '#adadad'
        self._border_dark = '#6a6a6a'
        
        self._subx = subx
        self._suby = suby
        self.image_filename = image_filename
        if image_filename:
            self.image = self.open_image(image_filename, subx, suby)
        else:
            self.image = None
            
        self.command = command
        self.height = height
        self.width = width
        self.overrelief = FLAT 
        self.relief = FLAT
        self.pady = pady
        self.padx = padx

        self.dark_mode = darkmode
        self.is_highlighted = highlight

        self.surroundings = None
        self.activebackground = None
        self.bg = None
        self.highlighted = None
        self.fg = None
        self.highlighted_border = None
        self.border_color = None
        self.border_color = None
        self.highlighted_border = None
        self.highlighted = None
        self.activebackground = None
        self.bg = None

        self.set_colors()

        super().__init__(master=master, cnf=cnf, bd=self._normal_border, bg=self.surroundings, name=name, padx=self.padx, 
                        pady=self.pady, relief=self.relief)

        self._dynamicFrame = Frame(self, borderwidth=self._normal_border, background=self.border_color, relief=self.relief)
        self._button = Button(self._dynamicFrame, cnf=cnf, bg=self.bg, highlightthickness=self._no_border, bd=self._no_border, activebackground=self.activebackground, 
                            fg=self.fg, overrelief=self.overrelief, relief=self.relief, activeforeground=self.fg,
                            anchor=anchor, takefocus=takefocus,command=self.command, compound=compound, default=default, 
                            height=self.height, image=self.image, justify=justify, name=name, state=state, text=text, 
                            textvariable=textvariable, underline=underline, width=self.width, wraplength=wraplength)
        
        self._dynamicFrame.pack(expand=True,fill='both')
        self._button.pack(expand=True,fill='both')

        if self.is_highlighted:
            self.focus_set(highlight = True)
            self._button.bind('<Return>', command)

        self._button.bind("<Enter>", self.on_enter)
        self._button.bind("<Leave>", self.on_leave)
        self._button.bind('<FocusIn>', self.focus_set)
        self._button.bind('<FocusOut>', self.focus_out)

    def on_enter(self, e):
        if self._button['state'] == 'normal':
            self._button.configure(background = self.highlighted)
            self._dynamicFrame.configure(borderwidth = self._normal_border)
            self.configure(borderwidth = self._normal_border)
            self._dynamicFrame.configure(background = self.highlighted_border)
    
    def on_leave(self, e):
        if self._button['state'] == 'normal':
            self._button.configure(background = self.bg)
            focus = self.master.focus_get() 
            if self._button == focus:
                self._dynamicFrame.configure(borderwidth = self._expanded_border)
                self.configure(borderwidth = self._no_border)
            else:
                self._dynamicFrame.configure(borderwidth = self._normal_border)
                self.configure(borderwidth = self._normal_border)
                self._dynamicFrame.configure(background = self.border_color)

    def configure(self, cnf=None, state=None, **kw):
        if state:
            if state == 'normal':
                pass
            self._button.configure(state=state)
        else:
            return super().configure(cnf, **kw)

    def focus_set(self, e=None, highlight=False):
        if not highlight:
            self._button.focus()
        else:
            self._button.bind('<Return>', self.command)
        self._dynamicFrame.configure(borderwidth = self._expanded_border)
        self._dynamicFrame.configure(background = self.highlighted_border)
        self.configure(borderwidth = self._no_border)
    focus = focus_set
    
    def focus_out(self, e):
        self._button.configure(default='normal')
        self._dynamicFrame.configure(borderwidth = self._normal_border)
        self._dynamicFrame.configure(background = self.border_color)
        self.configure(borderwidth = self._normal_border)
        self._button.unbind('<Return>')

    def set_colors(self):
        if self.dark_mode:
            self.surroundings = self._surroundings_dark
            self.activebackground = self._activebackground_dark
            self.bg = self._bg_dark
            self.highlighted = self._highlight_dark
            self.fg = self._fg_dark
            self.highlighted_border = self._highlight_border_dark
            self.border_color = self._border_dark
        else:
            self.surroundings = 'SystemButtonFace'
            self.activebackground = self._activebackground_light
            self.bg = self._bg_light
            self.highlighted = self._highlight_light
            self.fg = 'SystemWindowText'
            self.highlighted_border = self._highlight_border_light
            self.border_color = self._border_light

        if self.is_highlighted and self.dark_mode:
            self.border_color = self._highlight_border_light
            self.highlighted_border = self._highlight_border_light
            self.highlighted = '#1e4074'
            self.activebackground = '#07214d'
            self.bg = '#345695'

    def change_mode(self, updated_image_filename=None):
        self.dark_mode = not self.dark_mode
        self.set_colors()
        self.configure(bg=self.surroundings)
        self._dynamicFrame.configure(background=self.border_color)
        if updated_image_filename:
            self.image_filename = updated_image_filename
            self.image = self.open_image(updated_image_filename, self._subx, self._suby)
            self._button.configure(bg=self.bg, activebackground=self.activebackground, fg=self.fg, activeforeground=self.fg, image=self.image)
        else:
            self._button.configure(bg=self.bg, activebackground=self.activebackground, fg=self.fg, activeforeground=self.fg)
        self.update()

    def open_image(self, image_filename, subx, suby):
        return PhotoImage(file = image_filename).subsample(subx, suby)
