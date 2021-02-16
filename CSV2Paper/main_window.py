import csv
from os.path import abspath, join
from tkinter import *
from tkinter import filedialog

from drag_drop_listbox import DragDropListbox
from files import FilePaths, __location__
from mailmerge_tracking import MailMergeTracking
from update_checker import Updater
from windows_style_button import WindowsButton


class MainWindow:
	def __init__(self, base):
		self.files = FilePaths()

		self.base = base
		self.base.title("CSV 2 Paper")
		self.base.columnconfigure(1,weight=1)    #confiugures to stretch with a scaler of 1.
		self.base.rowconfigure(5,weight=1)
		self.base.columnconfigure(2,weight=1)
		self.base.protocol("WM_DELETE_WINDOW", self.on_closing)

		self.menu_bar = Menu(base)
		
		self.help_menu = Menu(self.menu_bar, tearoff=0)
		self.help_menu.add_command(label="About...", command=self.about_popup)
		self.menu_bar.add_cascade(label="Help", menu=self.help_menu)
		
		self.base.config(menu=self.menu_bar)


		self.template = StringVar(value="Word Template")
		self.template_entry = Entry(textvariable=self.template)
		self.template_entry.configure(validate="focusout", validatecommand = lambda:self.template_file_text())
		self.template_file_selector = WindowsButton(base, text ='+', command = lambda:self.template_file_opener())
		self.template_entry.grid(row=1,column=1,columnspan=2,sticky='we',padx=(5, 30),pady=(0,0))
		self.template_file_selector.grid(row=1,column=1,columnspan=2,sticky=E,padx=(0, 5),pady=5)


		self.csv = StringVar(value="CSV")
		self.csv_entry = Entry(state='disabled', textvariable=self.csv)
		self.csv_entry.configure(validate="focusout", validatecommand = lambda:self.csv_file_text())
		self.csv_file_selector = WindowsButton(base, text ='+', state='disabled', command = lambda:self.csv_file_opener())
		self.csv_entry.grid(row=2,column=1,columnspan=2,sticky='we',padx=(5, 30),pady=5)
		self.csv_file_selector.grid(row=2,column=1,columnspan=2,sticky=E,padx=(0, 5),pady=5)


		self.folder = StringVar(value="Output Folder")
		self.folder_entry = Entry(state='disabled', textvariable=self.folder)
		self.folder_entry.configure(validate="focusout", validatecommand = lambda:self.directory_selctor_text())
		self.folder_selector = WindowsButton(base, text ='+', state='disabled', command = lambda:self.directory_selector())
		self.folder_entry.grid(row=3,column=1,columnspan=2,sticky='we',padx=(5, 30),pady=5)
		self.folder_selector.grid(row=3,column=1,columnspan=2,sticky=E,padx=(0, 5),pady=5)


		self.file_output_info_group = Frame(base)
		self.file_output_info_group.grid(row=4,column=1, columnspan=2,padx=5,pady=5, sticky='nsew')
		self.file_output_info_group.columnconfigure(0,weight=1)
		self.file_output_info_group.rowconfigure(0,weight=1)

		self.filename = StringVar(value="Output File")
		self.filename_entry = Entry(self.file_output_info_group, state='disabled', textvariable=self.filename)

		self.output_as_word = BooleanVar(value=True)
		self.docx_checkbox = Checkbutton(self.file_output_info_group, state='disabled', text='Word', variable=self.output_as_word, onvalue=True, offvalue=False, command=self.check_runnable)
		
		self.output_as_pdf = BooleanVar(value=True)
		self.pdf_checkbox = Checkbutton(self.file_output_info_group, state='disabled', text='PDF', variable=self.output_as_pdf, onvalue=True, offvalue=False, command=self.check_runnable)

		self.filename_entry.grid(row=0,column=0,sticky='we',padx=(0, 30))
		self.docx_checkbox.grid(row=0,column=1,sticky='we',padx=(5, 30))
		self.pdf_checkbox.grid(row=0,column=3,sticky='we',padx=(5, 30))


		self.left_merge_fields_group = Frame(base)
		self.left_merge_fields_group.grid(row=5,column=1, padx=5,pady=5, sticky='nsew')
		self.left_merge_fields_group.columnconfigure(0,weight=1)
		self.left_merge_fields_group.rowconfigure(1,weight=1)

		self.merge_field_label = Label(self.left_merge_fields_group, text="Template Fields", justify=CENTER, padx=10, pady=5)
		self.merge_field_label.grid(row=0,column=0, sticky='nsew')

		self.scroll_merge_fields_y = Scrollbar(self.left_merge_fields_group, orient=VERTICAL)
		self.scroll_merge_fields_x = Scrollbar(self.left_merge_fields_group, orient=HORIZONTAL)

		self.merge_fields_listbox = Listbox(self.left_merge_fields_group, height=20, width=30, yscrollcommand=self.scroll_merge_fields_y.set, xscrollcommand=self.scroll_merge_fields_x.set, exportselection=0)
		self.merge_fields_listbox.bind("<<ListboxSelect>>", self.on_select)
		self.merge_fields_listbox.grid(row=1,column=0, sticky='nsew')

		self.scroll_merge_fields_y.config(command = self.merge_fields_listbox.yview)
		self.scroll_merge_fields_y.grid(row=1,column=1, sticky='nsew')

		self.scroll_merge_fields_x.config(command = self.merge_fields_listbox.xview)
		self.scroll_merge_fields_x.grid(row=2,column=0, sticky='nsew')


		self.right_headers_group = Frame(base)
		self.right_headers_group.grid(row=5,column=2, padx=5,pady=5, sticky='nsew')
		self.right_headers_group.columnconfigure(0,weight=1)
		self.right_headers_group.rowconfigure(1,weight=1)

		self.headers_label = Label(self.right_headers_group, text="Data Headers", justify=CENTER, padx=10, pady=5)
		self.headers_label.grid(row=0,column=0, sticky='nsew')

		self.scroll_headers_y = Scrollbar(self.right_headers_group, orient=VERTICAL)
		self.scroll_headers_x = Scrollbar(self.right_headers_group, orient=HORIZONTAL)

		self.headers_listbox = Listbox(self.right_headers_group, height=20, width=30, yscrollcommand=self.scroll_headers_y.set, xscrollcommand=self.scroll_headers_x.set, exportselection=0, )
		self.headers_listbox.bind("<<ListboxSelect>>", self.on_select)
		self.headers_listbox.grid(row=1,column=0, sticky='nsew')

		self.scroll_headers_y.config(command = self.headers_listbox.yview)
		self.scroll_headers_y.grid(row=1,column=1, sticky='nsew')

		self.scroll_headers_x.config(command = self.headers_listbox.xview)
		self.scroll_headers_x.grid(row=2,column=0, sticky='nsew')

		self.edit_header_buttons = Frame(self.right_headers_group)
		self.edit_header_buttons.grid(row=1,column=2, rowspan=2, padx=5,pady=5, sticky='nsew')

		self.up_arrow_icon = PhotoImage(file = join(__location__, 'resources','up-arrow.png')).subsample(30, 30)
		self.move_header_up_button = Button(self.edit_header_buttons, image=self.up_arrow_icon, command=lambda: self.move_up())
		self.move_header_up_button.grid(row=0,column=0, sticky='ew')

		self.down_arrow_icon = PhotoImage(file = join(__location__, 'resources', 'down-arrow.png')).subsample(30, 30)
		self.move_header_down_button = Button(self.edit_header_buttons, image=self.down_arrow_icon, command=lambda: self.move_down())
		self.move_header_down_button.grid(row=1,column=0, sticky='ew')

		self.run = WindowsButton(base, text ='Run', state='disabled', command = self.run_op)
		self.run.grid(row=6,column=1, columnspan=2,padx=5,pady=5)
		
		self.base.withdraw()
		Updater(self.base)

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
		csv_file = filedialog.askopenfilename(filetypes=[("CSV", ".csv")])
		self.files.csv_file = csv_file
		self.csv_entry.delete(0, END)
		self.csv_entry.insert(0, csv_file)
		with open(self.files.csv_file, encoding='utf8', newline='') as csv_file:
			csv_list = csv.reader(csv_file)
			headers = next(csv_list)
			'''
			data = next(csv_list) # new feature - show data preview
			output = []
			for i in range(len(headers)):
				output.append(headers[i]+" "+data[i])
			headers = sorted(output) # end new feature
			'''
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
			'''
			data = next(csv_list) # new feature - show data preview
			output = []
			for i in range(len(headers)):
				output.append(headers[i]+" "+data[i])
			headers = sorted(output) # end new feature
			'''
			headers = sorted(headers)
			self.headers_listbox.delete(0,END)
			for header in headers:
				self.headers_listbox.insert(END, header)

	def directory_selector(self):
		folder_selected = filedialog.askdirectory()
		self.files.folder = folder_selected
		self.folder_entry.delete(0,END)
		self.folder_entry.insert(0,folder_selected)

		self.filename_entry.configure(state='normal')
		self.run.configure(state='normal')
		self.pdf_checkbox.configure(state='normal')
		self.docx_checkbox.configure(state='normal')

	def directory_selctor_text(self):
		self.files.folder = self.folder_entry.get()

	def on_select(self, sender):
		listbox = sender.widget
		try:
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
		except:
			pass

	def move_up(self):
		try:
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
		except:
			pass

	def move_down(self):
		try:
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
		except:
			pass
	
	def check_runnable(self):
		if not self.output_as_word.get() and not self.output_as_pdf.get():
			self.run.configure(state='disabled')
		if self.output_as_word.get() or self.output_as_pdf.get():
			self.run.configure(state='normal')

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
		map = {}
		for index in range(len(fields)):
			map[fields[index]] = headers[index]
		return map

	def run_op(self):
		map = self.map_fields()
		self.files.template = self.template_entry.get()
		self.files.csv_file = self.csv_entry.get()
		self.files.folder = self.folder_entry.get()
		Run(self.base, map, self.files, self.output_as_word.get(), self.output_as_word.get())

	def on_closing(self):
		self.base.destroy()