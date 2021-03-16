from tkinter import *
from user_settings import UserSettings
from main_window import MainWindow

if __name__ == '__main__':
	base = Tk()
	user_settings = UserSettings()
	MainWindow(base, user_settings)
	base.mainloop()