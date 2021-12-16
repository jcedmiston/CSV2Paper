from os.path import join

from files import __location__
from user_settings import UserSettings


class BaseWindow:
	def __init__(self, user_settings: UserSettings):
		self.user_settings = user_settings

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
		
	def set_colors(self):
		if self.user_settings.dark_mode_enabled:
			self.window_bg = 'gray15'
			self.widget_bg = 'gray35'
			self.fg = 'white'
			self.insert_bg = 'white'
			self.disabled_bg = 'gray20'
			self.select_bg = 'gray30'
			self.folder_icon_file = join(__location__, 'resources', 'folder_open', '2x', 'sharp_folder_open_white_48dp.png')
			self.up_arrow_icon_file = join(__location__, 'resources', 'cheveron_up', '2x', 'sharp_chevron_up_white_48dp.png')
			self.down_arrow_icon_file = join(__location__, 'resources', 'cheveron_down', '2x', 'sharp_chevron_down_white_48dp.png')
		else:
			self.window_bg = 'SystemButtonFace'
			self.widget_bg = 'SystemWindow'
			self.fg = 'SystemWindowText'
			self.insert_bg = 'SystemWindowText'
			self.disabled_bg = 'gray80'
			self.select_bg = 'SystemWindow'
			self.folder_icon_file = join(__location__, 'resources', 'folder_open', '2x', 'sharp_folder_open_black_48dp.png')
			self.up_arrow_icon_file = join(__location__, 'resources', 'cheveron_up', '2x', 'sharp_chevron_up_black_48dp.png')
			self.down_arrow_icon_file = join(__location__, 'resources', 'cheveron_down', '2x', 'sharp_chevron_down_black_48dp.png')
