import csv
import re
from tkinter import *
from tkinter import filedialog
from os.path import abspath
from mailmerge import MailMerge
from WriteOut import write_out

class FilePaths:
	def __init__(self, responsesFilePath, template, folder, filename):
		self.responsesFilePath = responsesFilePath
		self.template = template
		self.folder = folder
		self.filename = filename

def map_fields(fields_list, headers_list, filename, word, pdf):
	headers = headers_list.get(0,END)
	fields = fields_list.get(0,END)
	map = {}
	print(word)
	print(pdf)
	for index in range(len(fields)):
		map[fields[index]] = headers[index]
	write_out(map, files.responsesFilePath, files.template, files.folder, filename, word, pdf)

class App:
	def __init__(self, base):
		self.base = base
		self.base.title("CSV 2 Paper")
		self.base.columnconfigure(1,weight=1)    #confiugures to stretch with a scaler of 1.
		self.base.rowconfigure(5,weight=1)
		self.base.columnconfigure(2,weight=1)     
		self.base.geometry('450x500')
		
		self.menu_bar = Menu(base)
		
		self.help_menu = Menu(self.menu_bar, tearoff=0)
		self.help_menu.add_command(label="About...", command=self.about_popup)
		self.menu_bar.add_cascade(label="Help", menu=self.help_menu)
		
		self.base.config(menu=self.menu_bar)


		self.template = StringVar(value="Template")
		self.template_entry = Entry(textvariable=self.template)
		self.template_file_selector = Button(base, text ='+', command = lambda:self.template_file_opener())
		self.template_entry.grid(row=1,column=1,columnspan=2,sticky='we',padx=(5, 30),pady=(0,0))
		self.template_file_selector.grid(row=1,column=1,columnspan=2,sticky=E,padx=(0, 5),pady=5)


		self.csv = StringVar(value="CSV")
		self.csv_entry = Entry(state='disabled', textvariable=self.csv)
		self.csv_file_selector = Button(base, text ='+', state='disabled', command = lambda:self.csv_file_opener())
		self.csv_entry.grid(row=2,column=1,columnspan=2,sticky='we',padx=(5, 30),pady=5)
		self.csv_file_selector.grid(row=2,column=1,columnspan=2,sticky=E,padx=(0, 5),pady=5)


		self.folder = StringVar(value="Output Folder")
		self.folder_entry = Entry(state='disabled', textvariable=self.folder)
		self.folder_selector = Button(base, text ='+', state='disabled', command = lambda:self.directory_selector())
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
		self.left_merge_fields_group.rowconfigure(0,weight=1)

		self.scroll_merge_fields_y = Scrollbar(self.left_merge_fields_group, orient=VERTICAL)
		self.scroll_merge_fields_x = Scrollbar(self.left_merge_fields_group, orient=HORIZONTAL)

		self.merge_fields_listbox = Listbox(self.left_merge_fields_group, height=20, width=30, yscrollcommand=self.scroll_merge_fields_y.set, xscrollcommand=self.scroll_merge_fields_x.set)
		self.merge_fields_listbox.grid(row=0,column=0, sticky='nsew')

		self.scroll_merge_fields_y.config(command = self.merge_fields_listbox.yview)
		self.scroll_merge_fields_y.grid(row=0,column=1, sticky='nsew')

		self.scroll_merge_fields_x.config(command = self.merge_fields_listbox.xview)
		self.scroll_merge_fields_x.grid(row=1,column=0, sticky='nsew')


		self.right_side_headers_group = Frame(base)
		self.right_side_headers_group.grid(row=5,column=2, padx=5,pady=5, sticky='nsew')
		self.right_side_headers_group.columnconfigure(0,weight=1)
		self.right_side_headers_group.rowconfigure(0,weight=1)

		self.scroll_headers_y = Scrollbar(self.right_side_headers_group, orient=VERTICAL, width=10)
		self.scroll_headers_x = Scrollbar(self.right_side_headers_group, orient=HORIZONTAL, width=10)


		self.headers_listbox = Listbox(self.right_side_headers_group, height=20, width=30, yscrollcommand=self.scroll_headers_y.set, xscrollcommand=self.scroll_headers_x.set)
		self.headers_listbox.grid(row=0,column=0, sticky='nsew')

		self.scroll_headers_y.config(command = self.headers_listbox.yview)
		self.scroll_headers_y.grid(row=0,column=1, sticky='nsew')

		self.scroll_headers_x.config(command = self.headers_listbox.xview)
		self.scroll_headers_x.grid(row=1,column=0, sticky='nsew')

		self.edit_header_buttons = Frame(self.right_side_headers_group)
		self.edit_header_buttons.grid(row=0,column=2, rowspan=2, padx=5,pady=5, sticky='nsew')

		self.up_arrow_icon = PhotoImage(file = "./up-arrow.png").subsample(30, 30)
		self.move_header_up_button = Button(self.edit_header_buttons, image=self.up_arrow_icon, command=lambda: self.move_up(self.headers_listbox))
		self.move_header_up_button.grid(row=0,column=0, sticky='ew')

		self.down_arrow_icon = PhotoImage(file = "./down-arrow.png").subsample(30, 30)
		self.move_header_down_button = Button(self.edit_header_buttons, image=self.down_arrow_icon, command=lambda: self.move_down(self.headers_listbox))
		self.move_header_down_button.grid(row=1,column=0, sticky='ew')


		self.run = Button(base, text ='Run', state='disabled', command = lambda:map_fields(self.merge_fields_listbox, self.headers_listbox, self.filename.get(), word=self.output_as_word.get(), pdf=self.output_as_pdf.get()))
		self.run.grid(row=6,column=1, columnspan=2,padx=5,pady=5)

	def template_file_opener(self):
		template_file = filedialog.askopenfilename(filetypes=[("Word Document", ".docx")])
		files.template = abspath(template_file)
		self.template_entry.delete(0,END)
		self.template_entry.insert(0,template_file)
		with MailMerge(files.template) as document:
			fields = document.get_merge_fields()
			fields = sorted(fields)
			for field in fields:
				self.merge_fields_listbox.insert(END, field)

		self.csv_entry.configure(state='normal')
		self.csv_file_selector.configure(state='normal')

	def csv_file_opener(self):
		csv_file = filedialog.askopenfilename()
		files.responsesFilePath = abspath(csv_file)
		self.csv_entry.delete(0, END)
		self.csv_entry.insert(0, csv_file)
		with open(csv_file, encoding='utf8', newline='') as auditionsFile:
			auditions = csv.reader(auditionsFile)
			headers = next(auditions)
			headers = sorted(headers)
			self.headers_listbox.delete(0,END)
			for header in headers:
				self.headers_listbox.insert(END, header)

		self.folder_entry.configure(state='normal')
		self.folder_selector.configure(state='normal')
					

	def directory_selector(self):
		folder_selected = filedialog.askdirectory()
		files.folder = abspath(folder_selected)
		self.folder_entry.delete(0,END)
		self.folder_entry.insert(0,folder_selected)

		self.filename_entry.configure(state='normal')
		self.run.configure(state='normal')
		self.pdf_checkbox.configure(state='normal')
		self.docx_checkbox.configure(state='normal')

	def move_up(self, list_box):
		try:
			idxs = list_box.curselection()
			if not idxs:
				return
			for pos in idxs:
				if pos==0:
					continue
				text=list_box.get(pos)
				list_box.delete(pos)
				list_box.insert(pos-1, text)
				list_box.selection_set(pos-1)
		except:
			pass

	def move_down(self, list_box):
		try:
			idxs = list_box.curselection()
			if not idxs:
				return
			for pos in idxs:
				# Are we at the bottom of the list?
				if pos == list_box.size()-1: 
					continue
				text=list_box.get(pos)
				list_box.delete(pos)
				list_box.insert(pos+1, text)
				list_box.selection_set(pos + 1)
		except:
			pass
	
	def check_runnable(self):
		if not self.output_as_word.get() and not self.output_as_pdf.get():
			self.run.configure(state='disabled')
		if self.output_as_word.get() or self.output_as_pdf.get():
			self.run.configure(state='normal')

	def about_popup(self):
			about_win = Toplevel()
			about_win.wm_title("About CSV 2 Paper")
			about_win.resizable(0, 0)
			about_win.columnconfigure(0,weight=1)
			about_text = """Material Design Icon Pack made by\nGoogle (flaticon.com/authors/google")\nRetrieved from Flaticon (flaticon.com)\nLicense under Creative Commons 3.0 BY\n(creativecommons.org/licenses/by/3.0/)"""
			about_label = Label(about_win, text=about_text, justify=CENTER, padx=10, pady=5)
			about_label.grid(row=0, column=0, sticky="nsew")
			close_button = Button(about_win, text="Close", command=about_win.destroy, justify=CENTER)
			close_button.grid(row=1, column=0, pady=5, padx=5)

files = FilePaths("/", "/", "/", "Output File")

if __name__ == '__main__':
	base = Tk()
	app = App(base)
	base.mainloop()