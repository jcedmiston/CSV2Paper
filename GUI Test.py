from tkinter import *
from tkinter import filedialog
import csv
from os import error as osError
from os import rename
from os.path import join, dirname, basename, isdir, isfile, abspath
from datetime import datetime
from mailmerge import MailMerge

class filePaths:
    def __init__(self, responsesFilePath, template, folder):
        self.responsesFilePath = responsesFilePath
        self.template = template
        self.folder = folder

def writeOut(responsesFilePath, template, folder):
    with open(responsesFilePath, encoding='utf8', newline='') as auditionsFile:
        auditions = csv.reader(auditionsFile)

        next(auditions)
        for audition in auditions:
            document = MailMerge(template)
            document.merge(
                auditon_number=audition[1],
                childs_name=audition[7],
                date_of_birth=datetime.strptime(audition[8][:15], "%a %b %d %Y").strftime("%m/%d/%Y"),
                age=audition[9],
                ballet_school=audition[10],
                training_years=audition[11],
                hrs_per_week=audition[12],
                training_level=audition[13],
                other_training=audition[14],
                danced_before=audition[16],
                previous_role_1=audition[17],
                previous_role_2=audition[18],
                previous_role_3=audition[19],
                other_leads=audition[15],
                height=audition[20],
                weight=audition[21],
                bust=audition[22],
                waist=audition[23],
                hips=audition[24],
                girth=audition[25],
                inseam=audition[26],
                guardian_name=audition[2],
                phone_number=audition[4],
                parents_email=audition[3],
                street_address=audition[6],
                conflict_dates=audition[27])
            
            filename = str(audition[1]).strip()+"-"+str(audition[7]).strip()+".docx"
            filepath = abspath(join(folder, filename))

            document.write(filepath)
            document.close()

def aboutPopup():
    aboutWin = Toplevel()
    aboutWin.wm_title("About CSV 2 Paper")
    labelBonus = Label(aboutWin, text="Input")
    labelBonus.grid(row=0, column=0)
    B1 = Button(aboutWin, text="Okay", command=aboutWin.destroy)
    B1.grid(row=1, column=0)

base = Tk()
base.title("CSV 2 Paper")
menubar = Menu(base)
# Create a canvas
base.columnconfigure(1,weight=1)    #confiugures column 0 to stretch with a scaler of 1.
base.rowconfigure(4,weight=1)       #confiugures row 0 to stretch with a scaler of 1.
base.columnconfigure(2,weight=1)    #confiugures column 0 to stretch with a scaler of 1.
base.rowconfigure(4,weight=1)       #confiugures row 0 to stretch with a scaler of 1.
base.geometry('450x500')
files = filePaths("/", "/", "/")

# Function for opening the file
def template_file_opener():
    template_file = filedialog.askopenfilename()
    files.template = abspath(template_file)
    template_entry.delete(0,END)
    template_entry.insert(0,template_file)
    with MailMerge(files.template) as document:
        fields = document.get_merge_fields()
        for field in fields:
            fieldsbox.insert(END, field)

    csv_entry.configure(state='normal')
    csv_entry.insert(0, 'CSV')
    csv_file_selector.configure(state='normal')

def csv_file_opener():
    csv_file = filedialog.askopenfilename()
    files.responsesFilePath = abspath(csv_file)
    csv_entry.delete(0,END)
    csv_entry.insert(0,csv_file)
    with open(csv_file, encoding='utf8', newline='') as auditionsFile:
        auditions = csv.reader(auditionsFile)
        headers = next(auditions)
        headersbox.delete(0,END)
        for header in headers:
            headersbox.insert(END, header)

    folder_entry.configure(state='normal')
    folder_entry.insert(0, 'Output Folder')
    folder_selector.configure(state='normal')
            

def directory_selector():
    folder_selected = filedialog.askdirectory()
    files.folder = abspath(folder_selected)
    folder_entry.delete(0,END)
    folder_entry.insert(0,folder_selected)

def field_labels(file):
    with MailMerge(files.template) as document:
        fields = document.get_merge_fields()
        fieldsbox.delete(0,END)
        for field in fields:
            fieldsbox.insert(END, field)

def moveup():
    try:
        idxs = headersbox.curselection()
        if not idxs:
            return
        for pos in idxs:
            if pos==0:
                continue
            text=headersbox.get(pos)
            headersbox.delete(pos)
            headersbox.insert(pos-1, text)
            headersbox.selection_set(pos-1)
    except:
        pass

def movedown():
    try:
        idxs = headersbox.curselection()
        if not idxs:
            return
        for pos in idxs:
            # Are we at the bottom of the list?
            if pos == headersbox.size()-1: 
                continue
            text=headersbox.get(pos)
            headersbox.delete(pos)
            headersbox.insert(pos+1, text)
            headersbox.selection_set(pos + 1)
    except:
        pass

helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="About...", command=aboutPopup())
menubar.add_cascade(label="Help", menu=helpmenu)
base.config(menu=menubar)

upArrow = PhotoImage(file = "./up-arrow.png").subsample(12, 12)
downArrow = PhotoImage(file = "./down-arrow.png").subsample(12, 12)

template_entry = Entry()
template_entry.insert(0, 'Template')

csv_entry = Entry(state='disabled')
folder_entry = Entry(state='disabled')

template_file_selector = Button(base, text ='+', command = lambda:template_file_opener())
csv_file_selector = Button(base, text ='+', state='disabled', command = lambda:csv_file_opener())
folder_selector = Button(base, text ='+', state='disabled', command = lambda:directory_selector())
run = Button(base, text ='Run', command = lambda:writeOut(responsesFilePath=files.responsesFilePath,
                                                            template=files.template,
                                                            folder=files.folder))

#template_entry_label = Label(base, text = "Template", font = ("Arial",10), justify=LEFT)
#template_entry_label.grid(row=0,column=1,sticky='w',padx=5)
template_entry.grid(row=1,column=1,columnspan=2,sticky='we',padx=(5, 30),pady=(0,0))
template_file_selector.grid(row=1,column=1,columnspan=2,sticky=E,padx=(0, 5),pady=5)

csv_entry.grid(row=2,column=1,columnspan=2,sticky='we',padx=(5, 30),pady=5)
csv_file_selector.grid(row=2,column=1,columnspan=2,sticky=E,padx=(0, 5),pady=5)

folder_entry.grid(row=3,column=1,columnspan=2,sticky='we',padx=(5, 30),pady=5)
folder_selector.grid(row=3,column=1,columnspan=2,sticky=E,padx=(0, 5),pady=5)

run.grid(row=5,column=1, columnspan=2,padx=5,pady=5)

left = Frame(base)
left.grid(row=4,column=1, padx=5,pady=5, sticky='nsew')
left.columnconfigure(0,weight=1)
left.rowconfigure(0,weight=1)

right = Frame(base)
right.columnconfigure(0,weight=1)
right.rowconfigure(0,weight=1)

right.grid(row=4,column=2, padx=5,pady=5, sticky='nsew')
fieldsbox = Listbox(left, height=20, width=30)
fieldsbox.grid(row=0,column=0, sticky='nsew')


scroll_fields_y = Scrollbar(left, orient=VERTICAL)
scroll_fields_x = Scrollbar(left, orient=HORIZONTAL)

fieldsbox.config(yscrollcommand=scroll_fields_y.set)
fieldsbox.config(xscrollcommand=scroll_fields_x.set)

scroll_fields_y.config(command = fieldsbox.yview)
scroll_fields_x.config(command = fieldsbox.xview)


scroll_fields_y.grid(row=0,column=1, sticky='nsew')
scroll_fields_x.grid(row=1,column=0, sticky='nsew')

headersbox = Listbox(right, height=20, width=30)
scroll_headers_y = Scrollbar(right, orient=VERTICAL, width=10)
scroll_headers_x = Scrollbar(right, orient=HORIZONTAL, width=10)
headersbox.grid(row=0,column=0, sticky='nsew')

buttons = Frame(right)
buttons.grid(row=0,column=2, rowspan=2, padx=5,pady=5, sticky='nsew')

moveUpButton = Button(buttons, image=upArrow, command=lambda: moveup())
moveDownButton = Button(buttons, image=downArrow, command=lambda: movedown())

scroll_headers_x.grid(row=1,column=0, sticky='nsew')
scroll_headers_y.grid(row=0,column=1, sticky='nsew')
moveUpButton.grid(row=0,column=0, sticky='ew')
moveDownButton.grid(row=1,column=0, sticky='nsew')
headersbox.config(yscrollcommand=scroll_headers_y.set)
headersbox.config(xscrollcommand=scroll_headers_x.set)

scroll_headers_y.config(command = headersbox.yview)

scroll_headers_x.config(command = headersbox.xview)

mainloop()