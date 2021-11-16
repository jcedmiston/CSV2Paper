import csv
import json
import sys
from os.path import abspath, isfile, join, basename
from tkinter import *
from tkinter import filedialog, messagebox

from convert import Convert
from files import FilePaths, __location__
from mailmerge_tracking import MailMergeTracking
from update_checker import Updater
from user_settings import UserSettings
from windows_style_button import WindowsButton


class MainWindow:
	def __init__(self, base: Tk, user_settings: UserSettings):
		self.user_settings = user_settings
		
		self.files = FilePaths()

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

		self.folder_icon_file_dark = join(__location__, 'resources', 'folder_open', '2x', 'sharp_folder_open_black_48dp.png')
		self.up_arrow_icon_file_dark = join(__location__, 'resources', 'cheveron_up', '2x', 'sharp_chevron_up_black_48dp.png')
		self.down_arrow_icon_file_dark = join(__location__, 'resources', 'cheveron_down', '2x', 'sharp_chevron_down_black_48dp.png')
		self.folder_icon_file_light = join(__location__, 'resources', 'folder_open', '2x', 'sharp_folder_open_white_48dp.png')
		self.up_arrow_icon_file_light = join(__location__, 'resources', 'cheveron_up', '2x', 'sharp_chevron_up_white_48dp.png')
		self.down_arrow_icon_file_light = join(__location__, 'resources', 'cheveron_down', '2x', 'sharp_chevron_down_white_48dp.png')
		self.matching_icon_filenames = {self.folder_icon_file_light: self.folder_icon_file_dark, self.folder_icon_file_dark: self.folder_icon_file_light,
										self.down_arrow_icon_file_light: self.down_arrow_icon_file_dark, self.down_arrow_icon_file_dark: self.down_arrow_icon_file_light,
										self.up_arrow_icon_file_light: self.up_arrow_icon_file_dark, self.up_arrow_icon_file_dark: self.up_arrow_icon_file_light}

		# setup main window attributes
		self.base = base
		if self.user_settings.check_for_updates_on_start.get():
			self.base.withdraw()
		self.base.title("CSV 2 Paper")
		self.base.iconbitmap(default=join(__location__, 'resources', 'icons', '32x32.ico'))
		self.base.columnconfigure(1,weight=1)    #confiugures to stretch with a scaler of 1.
		self.base.rowconfigure(5,weight=1)
		self.base.columnconfigure(2,weight=1)
		self.base.protocol("WM_DELETE_WINDOW", self.on_closing)
		
		# setup menubar
		self.menu_bar = Menu(base)
		
		self.file_menu = Menu(self.menu_bar, tearoff=0)
		self.file_menu.add_command(label="Open Configuration Template", command=self.open_setup_template)
		self.file_menu.add_command(label="Save Current Configuration", state="disabled", command=self.save_setup_template)

		self.switch_theme = Menu(self.menu_bar, tearoff=0)
		self.switch_theme.add_radiobutton(label="Use System Theme", value='system', variable=self.user_settings.default_theme, command=self.set_mode)
		self.switch_theme.add_radiobutton(label="Dark", value="dark", variable=self.user_settings.default_theme, command=self.set_mode)
		self.switch_theme.add_radiobutton(label="Light", value="light", variable=self.user_settings.default_theme, command=self.set_mode)
		
		self.options_menu = Menu(self.menu_bar, tearoff=0)
		self.options_menu.add_checkbutton(label='Check for updates on start', variable=self.user_settings.check_for_updates_on_start, onvalue=True, offvalue=False)
		self.options_menu.add_command(label="Check for updates...", command=self.update)
		self.options_menu.add_cascade(label="Switch Theme", menu=self.switch_theme)

		self.help_menu = Menu(self.menu_bar, tearoff=0)
		self.help_menu.add_command(label="About...", command=self.about_popup)
		
		self.menu_bar.add_cascade(label="File", menu=self.file_menu)
		self.menu_bar.add_cascade(label="Options", menu=self.options_menu)
		self.menu_bar.add_cascade(label="Help", menu=self.help_menu)
		self.base.config(menu=self.menu_bar)

		self.template = StringVar(value="Word Template")
		self.template_entry = Entry(textvariable=self.template, relief=FLAT)
		self.template_entry.configure(validate="focusout", validatecommand=self.template_file_text)
		self.template_file_selector = WindowsButton(base, darkmode=self.user_settings.dark_mode_enabled, image_filename=self.folder_icon_file, subx=6, suby=6, command=self.template_file_opener)
		self.template_entry.grid(row=1,column=1,columnspan=2,sticky='we',padx=(5, 30),pady=(0,0))
		self.template_file_selector.grid(row=1,column=1,columnspan=2,sticky=E,padx=(0, 5),pady=5)


		self.csv = StringVar(value="CSV")
		self.csv_entry = Entry(state='disabled', textvariable=self.csv, relief=FLAT)
		self.csv_entry.configure(validate="focusout", validatecommand=self.csv_file_text)
		self.csv_file_selector = WindowsButton(base, darkmode=self.user_settings.dark_mode_enabled, image_filename=self.folder_icon_file, subx=6, suby=6, state='disabled', command=self.csv_file_opener)
		self.csv_entry.grid(row=2,column=1,columnspan=2,sticky='we',padx=(5, 30),pady=5)
		self.csv_file_selector.grid(row=2,column=1,columnspan=2,sticky=E,padx=(0, 5),pady=5)


		self.folder = StringVar(value="Output Folder")
		self.folder_entry = Entry(state='disabled', textvariable=self.folder, relief=FLAT)
		self.folder_entry.configure(validate="focusout", validatecommand=self.directory_selctor_text)
		self.folder_selector = WindowsButton(base, darkmode=self.user_settings.dark_mode_enabled, image_filename=self.folder_icon_file, subx=6, suby=6, state='disabled', command=self.directory_selector)
		self.folder_entry.grid(row=3,column=1,columnspan=2,sticky='we',padx=(5, 30),pady=5)
		self.folder_selector.grid(row=3,column=1,columnspan=2,sticky=E,padx=(0, 5),pady=5)


		self.file_output_info_group = Frame(base, bg=self.window_bg)
		self.file_output_info_group.grid(row=4,column=1, columnspan=2,padx=5,pady=5, sticky='nsew')
		self.file_output_info_group.columnconfigure(0,weight=1)
		self.file_output_info_group.rowconfigure(0,weight=1)

		self.filename = StringVar(value="Output File")
		self.filename_entry = Entry(self.file_output_info_group, state='disabled', textvariable=self.filename, relief=FLAT)

		self.output_as_word = BooleanVar(value=True)
		self.docx_checkbox = Checkbutton(self.file_output_info_group, state='disabled', relief=FLAT, offrelief=FLAT, overrelief=FLAT, text='Word', variable=self.output_as_word, onvalue=True, offvalue=False, command=self.check_runnable)
		
		self.output_as_pdf = BooleanVar(value=True)
		self.pdf_checkbox = Checkbutton(self.file_output_info_group, state='disabled', relief=FLAT, offrelief=FLAT, overrelief=FLAT, text='PDF', variable=self.output_as_pdf, onvalue=True, offvalue=False, command=self.check_runnable)

		self.filename_entry.grid(row=0,column=0,sticky='we',padx=(0, 30))
		self.docx_checkbox.grid(row=0,column=1,sticky='we',padx=(5, 30))
		self.pdf_checkbox.grid(row=0,column=3,sticky='we',padx=(5, 30))

		self.left_merge_fields_group = Frame(base, bg=self.window_bg)
		self.left_merge_fields_group.grid(row=5,column=1, padx=5,pady=5, sticky='nsew')
		self.left_merge_fields_group.columnconfigure(0,weight=1)
		self.left_merge_fields_group.rowconfigure(1,weight=1)

		self.merge_field_label = Label(self.left_merge_fields_group, bg=self.window_bg, fg=self.fg, text="Template Fields", justify=CENTER, padx=10, pady=5)
		self.merge_field_label.grid(row=0,column=0, sticky='nsew')

		self.scroll_merge_fields_y = Scrollbar(self.left_merge_fields_group)
		self.scroll_merge_fields_x = Scrollbar(self.left_merge_fields_group)

		self.merge_fields_listbox = Listbox(self.left_merge_fields_group, bd=0, highlightthickness=0, height=20, width=30, relief=FLAT, yscrollcommand=self.scroll_merge_fields_y.set, xscrollcommand=self.scroll_merge_fields_x.set, exportselection=0)
		self.merge_fields_listbox.bind("<<ListboxSelect>>", self.on_select)
		self.merge_fields_listbox.grid(row=1,column=0, sticky='nsew')

		self.scroll_merge_fields_y.config(command = self.merge_fields_listbox.yview)
		self.scroll_merge_fields_x.config(command = self.merge_fields_listbox.xview)

		self.right_headers_group = Frame(base)
		self.right_headers_group.grid(row=5,column=2, padx=5,pady=5, sticky='nsew')
		self.right_headers_group.columnconfigure(0,weight=1)
		self.right_headers_group.rowconfigure(1,weight=1)

		self.headers_label = Label(self.right_headers_group, text="Data Headers", justify=CENTER, padx=10, pady=5)
		self.headers_label.grid(row=0,column=0, sticky='nsew')

		self.scroll_headers_y = Scrollbar(self.right_headers_group)
		self.scroll_headers_x = Scrollbar(self.right_headers_group)

		self.headers_listbox = Listbox(self.right_headers_group, bd=0, highlightthickness=0, height=20, width=30, relief=FLAT, yscrollcommand=self.scroll_headers_y.set, xscrollcommand=self.scroll_headers_x.set, exportselection=0)
		self.headers_listbox.bind("<<ListboxSelect>>", self.on_select)
		self.headers_listbox.grid(row=1,column=0, sticky='nsew')

		self.scroll_headers_y.config(command = self.headers_listbox.yview)
		self.scroll_headers_x.config(command = self.headers_listbox.xview)

		self.edit_header_buttons = Frame(self.right_headers_group)
		self.edit_header_buttons.grid(row=1,column=2, rowspan=2, padx=5,pady=5, sticky='nsew')

		self.move_header_up_button = WindowsButton(self.edit_header_buttons, darkmode=self.user_settings.dark_mode_enabled, image_filename=self.up_arrow_icon_file, subx=4, suby=4, command=self.move_up)
		self.move_header_up_button.grid(row=0,column=0, sticky='ew')

		self.move_header_down_button = WindowsButton(self.edit_header_buttons, darkmode=self.user_settings.dark_mode_enabled, image_filename=self.down_arrow_icon_file, subx=4, suby=4, command=self.move_down)
		self.move_header_down_button.grid(row=1,column=0, sticky='ew')

		self.break_type_options = ['Page Break', 'Column Break', 'Text Wrapping Break', 'Continuous Section', 'Even Page Section', 'Next Column Section', 'Next Page Section', 'Odd Page Section']
		self.break_type_options_map = {
			'Page Break': 'page_break',
			'Column Break': 'column_break',
			'Text Wrapping Break': 'textWrapping_break',
			'Continuous Section': 'continuous_section',
			'Even Page Section': 'evenPage_section',
			'Next Column Section': 'nextColumn_section',
			'Next Page Section': 'nextPage_section',
			'Odd Page Section': 'oddPage_section'
		}
		self.break_type = StringVar()
		self.break_type.set(self.break_type_options[0])
		
		#self.break_type_select = OptionMenu(base, self.break_type, *self.break_type_options)
		#self.break_type_select["menu"].config(relief=FLAT)
		#self.break_type_select["menu"].config(activeborderwidth=0)
		#self.break_type_select["menu"].config(bd=0)
		#self.break_type_select.grid(row=6,column=1, padx=5, pady=5)
		
		self.test_run = WindowsButton(base, darkmode=self.user_settings.dark_mode_enabled, text ='Test Run', state='disabled', command = self.run_limited_op)
		self.test_run.grid(row=6,column=1, padx=5, pady=5)

		self.run = WindowsButton(base, darkmode=self.user_settings.dark_mode_enabled, text ='Run', state='disabled', command = self.run_op)
		self.run.grid(row=6,column=2, padx=5, pady=5)
		
		self.set_mode(first_run=True)
		if self.user_settings.check_for_updates_on_start.get():
			self.update(True)
		if len(sys.argv) > 1:
			template_file = sys.argv[1]
			self.open_setup_template(template_file)

	def open_setup_template(self, template_file=None):
		if not template_file: 
			template_file = filedialog.askopenfilename(filetypes=[("CSV 2 Paper Configuration File", ".c2p")])

		with open(template_file, 'r') as saved_setup:
			saved_setup_data = json.load(saved_setup)
		
		self.file_menu.entryconfig("Save Current Configuration", state="normal")
		self.csv_entry.configure(state='normal')
		self.csv_file_selector.configure(state='normal')
		self.folder_entry.configure(state='normal')
		self.folder_selector.configure(state='normal')
		self.filename_entry.configure(state='normal')
		self.pdf_checkbox.configure(state='normal')
		self.docx_checkbox.configure(state='normal')
		self.test_run.configure(state='normal')
		self.run.configure(state='normal')
		
		if self.check_for_file(saved_setup_data["template_file"], self.template_file_opener) is not None:
			self.files.template = saved_setup_data["template_file"]
			self.template_entry.delete(0, END)
			self.template_entry.insert(0, self.files.template)

		if self.check_for_file(saved_setup_data["csv_file"], self.csv_file_opener) is not None:
			self.files.csv_file = saved_setup_data["csv_file"]
			self.csv_entry.delete(0, END)
			self.csv_entry.insert(0, self.files.csv_file)

		self.files.folder = saved_setup_data["folder"]
		self.folder_entry.delete(0, END)
		self.folder_entry.insert(0, self.files.folder)

		self.files.filename = saved_setup_data["filename"]
		self.filename_entry.delete(0, END)
		self.filename_entry.insert(0, self.files.filename)
		self.output_as_pdf.set(saved_setup_data["output_as_pdf"])
		self.output_as_word.set(saved_setup_data["output_as_word"])

		self.merge_fields_listbox.delete(0,END)
		self.headers_listbox.delete(0,END)

		with MailMergeTracking(self.files.template) as document:
			fields = document.get_merge_fields()
			fields = sorted(fields)
		
		with open(self.files.csv_file, encoding='utf8', newline='') as csv_file:
			csv_list = csv.reader(csv_file)
			headers = next(csv_list)
			headers = sorted(headers)
		
		data_valid = True
		if not fields == sorted(saved_setup_data["fields"]):
			data_valid = False
		saved_headers = sorted(list(saved_setup_data["matched_fields"].values()))
		if not all(elem in headers for elem in saved_headers):
			data_valid = False
	
		if data_valid:
			for template_value, csv_value in saved_setup_data["matched_fields"].items():
				self.merge_fields_listbox.insert(END, template_value)
				self.headers_listbox.insert(END, csv_value)
			try:
				for csv_value in saved_setup_data["headers"][len(saved_setup_data["fields"]):]:
					self.headers_listbox.insert(END, csv_value)
			except IndexError:
				pass
		
	
	def check_for_file(self, file, callback):
		if isfile(file):
			return file
		if messagebox.askyesno("CSV 2 Paper", 'The file "'+ basename(file) +'" no longer exists or has been moved. Would you like to relocate it?'):
			callback()
			return None

	def save_setup_template(self):
		save_data = {}
		save_data["template_file"] = self.files.template
		save_data["csv_file"] = self.files.csv_file
		save_data["folder"] = self.files.folder
		save_data["filename"] = self.files.filename
		save_data["output_as_pdf"] = self.output_as_pdf.get()
		save_data["output_as_word"] = self.output_as_word.get()
		save_data["matched_fields"] = self.map_fields()
		save_data["fields"] = self.merge_fields_listbox.get(0,END)
		save_data["headers"] = self.headers_listbox.get(0,END)

		save_as_filename = filedialog.asksaveasfilename(defaultextension=".c2p", filetypes=[("CSV 2 Paper Configuration File", ".c2p")])
		with open(save_as_filename, "w", encoding="utf-8") as save_file:
			save_file.seek(0)
			save_file.write(json.dumps(save_data))
			save_file.close()

	def template_file_opener(self):
		template_file = filedialog.askopenfilename(filetypes=[("Word Document", ".docx")])
		self.files.template = template_file
		self.template_entry.delete(0,END)
		self.template_entry.insert(0,template_file)
		with MailMergeTracking(self.files.template) as document:
			fields = document.get_merge_fields()
			fields = sorted(fields)
			self.merge_fields_listbox.delete(0,END)
			for field in fields:
				self.merge_fields_listbox.insert(END, field)

		self.csv_entry.configure(state='normal')
		self.csv_file_selector.configure(state='normal')

	def template_file_text(self):
		self.files.template = abspath(self.template_entry.get())
		with MailMergeTracking(self.files.template) as document:
			fields = document.get_merge_fields()
			fields = sorted(fields)
			self.merge_fields_listbox.delete(0,END)
			for field in fields:
				self.merge_fields_listbox.insert(END, field)

	def csv_file_opener(self):
		csv_file = filedialog.askopenfilename(filetypes=[("Comma-Seperated Values (CSV)", ".csv")])
		self.files.csv_file = csv_file
		self.csv_entry.delete(0, END)
		self.csv_entry.insert(0, csv_file)
		with open(self.files.csv_file, encoding='utf8', newline='') as csv_file:
			csv_list = csv.reader(csv_file)
			headers = next(csv_list)
			
			#data = next(csv_list) # new feature - show data preview
			#output = []
			#for i in range(len(headers)):
			#	output.append(headers[i]+" "+data[i])
			#headers = sorted(output) # end new feature
			
			headers = sorted(headers)
			self.headers_listbox.delete(0,END)
			for header in headers:
				self.headers_listbox.insert(END, header)

		self.folder_entry.configure(state='normal')
		self.folder_selector.configure(state='normal')
	
	def csv_file_text(self):
		self.files.csv_file = self.csv_entry.get()
		with open(self.files.csv_file, encoding='utf8', newline='') as csv_file:
			csv_list = csv.reader(csv_file)
			headers = next(csv_list)
			
			#data = next(csv_list) # new feature - show data preview
			#output = []
			#for i in range(len(headers)):
			#	output.append(headers[i]+" "+data[i])
			#headers = sorted(output) # end new feature
			
			headers = sorted(headers)
			self.headers_listbox.delete(0,END)
			for header in headers:
				self.headers_listbox.insert(END, header)

	def directory_selector(self):
		folder_selected = filedialog.askdirectory()
		self.files.folder = folder_selected
		self.folder_entry.delete(0,END)
		self.folder_entry.insert(0,folder_selected)
		
		self.file_menu.entryconfig("Save Current Configuration", state="normal")
		self.filename_entry.configure(state='normal')
		self.run.configure(state='normal')
		self.test_run.configure(state='normal')
		self.pdf_checkbox.configure(state='normal')
		self.docx_checkbox.configure(state='normal')

	def directory_selctor_text(self):
		self.files.folder = self.folder_entry.get()

	def on_select(self, sender):
		listbox = sender.widget
		idxs = listbox.curselection()
		if not idxs:
			return
		for pos in idxs:
			self.merge_fields_listbox.selection_clear(0,END)
			self.merge_fields_listbox.selection_set(pos)
			self.merge_fields_listbox.see(pos)
			self.merge_fields_listbox.activate(pos)

			self.headers_listbox.selection_clear(0,END)
			self.headers_listbox.selection_set(pos)
			self.headers_listbox.activate(pos)

	def move_up(self):
		idxs = self.headers_listbox.curselection()
		if not idxs:
			return
		for pos in idxs:
			if pos==0:
				self.merge_fields_listbox.selection_clear(0,END)
				self.merge_fields_listbox.selection_set(pos)
				self.merge_fields_listbox.see(pos)
				continue
			text=self.headers_listbox.get(pos)
			self.headers_listbox.delete(pos)
			self.headers_listbox.insert(pos-1, text)
			self.headers_listbox.selection_set(pos-1)
			self.headers_listbox.activate(pos-1)
			self.merge_fields_listbox.selection_clear(0,END)
			self.merge_fields_listbox.selection_set(pos-1)

	def move_down(self):
		idxs = self.headers_listbox.curselection()
		if not idxs:
			return
		for pos in idxs:
			# Are we at the bottom of the list?
			if pos == self.headers_listbox.size()-1:
				self.merge_fields_listbox.selection_clear(0,END)
				self.merge_fields_listbox.selection_set(pos)
				self.merge_fields_listbox.see(pos)
				continue
			text=self.headers_listbox.get(pos)
			self.headers_listbox.delete(pos)
			self.headers_listbox.insert(pos+1, text)
			self.headers_listbox.selection_set(pos+1)
			self.headers_listbox.activate(pos+1)
			self.merge_fields_listbox.selection_clear(0,END)
			self.merge_fields_listbox.selection_set(pos+1)
	
	def check_runnable(self):
		if not self.output_as_word.get() and not self.output_as_pdf.get():
			self.run.configure(state='disabled')
		if self.output_as_word.get() or self.output_as_pdf.get():
			self.run.configure(state='normal')

	def update(self, on_start=False):
		Updater(self.base, self.user_settings, on_start=on_start)

	def about_popup(self):
		about_win = Toplevel(takefocus=True)
		about_win.focus_force()
		about_win.grab_set()
		about_win.wm_title("About CSV 2 Paper")
		about_win.resizable(0, 0)
		about_win.columnconfigure(0,weight=1)
		about_text = """Material Design Icon Pack made by\nGoogle (flaticon.com/authors/google")\nRetrieved from Flaticon (flaticon.com)\nLicense under Creative Commons 3.0 BY\n(creativecommons.org/licenses/by/3.0/)"""
		about_label = Label(about_win, text=about_text, justify=CENTER, padx=10, pady=5)
		about_label.grid(row=0, column=0, sticky="nsew")
		close_button = Button(about_win, text="Close", command=about_win.destroy, justify=CENTER)
		close_button.grid(row=1, column=0, pady=5, padx=5)
		about_win.update_idletasks() 
		x = self.base.winfo_rootx()
		y = self.base.winfo_rooty()
		x_offset = self.base.winfo_width() / 2 - about_win.winfo_width() / 2
		y_offset = self.base.winfo_height() / 4 - about_win.winfo_height() / 2
		geom = "+%d+%d" % (x+x_offset,y+y_offset)
		about_win.wm_geometry(geom)

	def map_fields(self):
		headers = self.headers_listbox.get(0,END)
		fields = self.merge_fields_listbox.get(0,END)
		mapped_fields = {}
		for i, field in enumerate(fields):
			mappped_fields[fields[i]] = headers[i]
		return map

	def run_limited_op(self):
		self.run_op(4)

	def run_op(self, limit=None):
		mapped_fields = self.map_fields()
		self.files.template = self.template_entry.get()
		self.files.csv_file = self.csv_entry.get()
		self.files.folder = self.folder_entry.get()
		Convert(self.base, mapped_fields, self.files, self.output_as_word.get(), self.output_as_word.get(), self.user_settings, limit)

	def on_closing(self):
		self.user_settings.save_to_disk()
		self.base.destroy()

	def set_mode(self, first_run=False):
		current_mode = self.user_settings.dark_mode_enabled
		self.user_settings.update_dark_mode()
		if current_mode == self.user_settings.dark_mode_enabled and not first_run:
			return
		def update_elements(base):
			for child in base.winfo_children():
				if isinstance(child, WindowsButton) and not first_run:
					if child.image_filename:
						child.change_mode(self.matching_icon_filenames[child.image_filename])
					else:
						child.change_mode()
				elif isinstance(child, Frame):
					child.configure(bg=self.window_bg)
					update_elements(child)
				elif isinstance(child, Label):
					child.configure(bg=self.window_bg, fg=self.fg)
				elif isinstance(child, Entry):
					child.configure(bg=self.widget_bg, fg=self.fg, insertbackground=self.insert_bg, disabledbackground=self.disabled_bg)
				elif isinstance(child, Checkbutton):
					child.configure(bg=self.window_bg, fg=self.fg, 
									activebackground=self.window_bg, activeforeground=self.fg, 
									selectcolor=self.select_bg)
				elif isinstance(child, Listbox):
					child.configure(bg=self.widget_bg, fg=self.fg)
		self.set_colors()
		self.base.configure(bg=self.window_bg)
		update_elements(self.base)
		self.base.update()
	
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
