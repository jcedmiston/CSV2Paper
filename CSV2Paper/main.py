from tkinter import Tk

from main_window import MainWindow
from user_settings import UserSettings


def main():
	base = Tk()
	user_settings = UserSettings()
	MainWindow(base, user_settings)
	base.mainloop()

if __name__ == '__main__':
	main()
