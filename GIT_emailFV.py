# -*- coding: utf-8 -*-
import os
import csv
import encodings
import smtplib
import calendar
import fnmatch
import shutil
import pandas as pd
from tkinter import *
from tkinter import filedialog, messagebox
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import date
from PIL import Image

root = Tk()
root.geometry('850x500') # sets size of root window
root.title("Sending e-mail")
Label(root, text='Please use below syntax:\n ;firm name; invoice no.; route; loading date / unloading date; number of CMR files; postal address; other information \n', fg='#19334d').pack(padx=0)

sensitiveData = pd.read_excel('private.xlsx')

iconsPath = sensitiveData.iloc[0][1]
attachmentsDirectory = sensitiveData.iloc[1][1]
fileToOpen = sensitiveData.iloc[6][1]
contentToSend = sensitiveData.iloc[7][1]


'''Images for icons:'''

openButtonIcon = PhotoImage(file = iconsPath + 'openBig.gif')
saveButtonIcon = PhotoImage(file = iconsPath + 'saveBig.gif')
sendButtonIcon = PhotoImage(file = iconsPath + 'sendBig.gif')
exportButtonIcon = PhotoImage(file = iconsPath + 'exportBig.gif')
saveAsMenuIcon = PhotoImage(file = iconsPath + 'save_as.gif')
exitMenuIcon = PhotoImage(file = iconsPath + 'exit.gif')
undoMenuIcon = PhotoImage(file = iconsPath + 'undo.gif')
redoMenuIcon = PhotoImage(file = iconsPath + 'redo.gif')
cutMenuIcon = PhotoImage(file = iconsPath + 'cut.gif')
copyMenuIcon = PhotoImage(file = iconsPath + 'copy.gif')
pasteMenuIcon = PhotoImage(file = iconsPath + 'paste.gif')
findMenuIcon = PhotoImage(file = iconsPath + 'find.gif')
newMenuIcon = PhotoImage(file = iconsPath + 'new.gif')
sendMenuIcon = PhotoImage(file = iconsPath + 'send.gif')
exportMenuIcon = PhotoImage(file = iconsPath + 'exportToResize.gif').subsample(3,3) # resizing
saveMenuIcon = PhotoImage(file = iconsPath + 'saveToResize.gif').subsample(3,3) # resizing

# defining "Edit" functions
def undo():
	textPad.event_generate("<<Undo>>")
def redo():
	textPad.event_generate("<<Redo>>")
def cut():
	textPad.event_generate("<<Cut>>")
def copy():
	textPad.event_generate("<<Copy>>")
def paste():
	textPad.event_generate("<<Paste>>")

# defining functions that will be needed for "exportBig()" and "":
def askopenfile():
	filedialog.askopenfile(mode='r')
def asksaveasFile():
	filedialog.asksaveasfile()

def exportBig():
	'''
	Info:
	Something like '---------------------------------' separates content
	to send from content that has been send last time.
	Below variable "separatingIndex" is set of rows that equal to '---------------------------------'.
	Thanks to it program can distinguish last new content that must be send.
	'''
	root.clipboard_clear()
	with open(fileToOpen, encoding='utf-8') as file:
		fileContent = file.readlines()
		separatingIndex = [x for x in range(len(fileContent)) if '---------------------------------' in fileContent[x]]
	fileToExport = filedialog.asksaveasfile(initialfile='export.csv', mode='w', defaultextension='.csv')
	if fileToExport is None:
		return
	newContent = textPad.get(float(separatingIndex[-1]+2), END)
	fileToExport.write(newContent)
	fileToExport.close()

def exit_editor(event=None):
	if messagebox.askokcancel("Quit?", "Do you really want to exit?", icon = 'warning'):
		root.destroy()

def new_file(self):
	root.title("Untitled")
	textPad.delete(1.0,END)

def openBig():
	fileContent = open(fileToOpen)
	for line in fileContent:
		textPad.insert(END, line)

def save_as_function():
	fileToSave = filedialog.asksaveasfile(mode='w', defaultextension='.txt')
	if fileToSave is None:
		return
	wholeContent = textPad.get(0.0, END)
	fileToSave.write(wholeContent)
	fileToSave.close()

def saveBig():
	openedFile = open(fileToOpen, 'w')
	fromFirstLineContent = textPad.get(1.0, END)
	openedFile.write(fromFirstLineContent)
	MsgBox = messagebox.askquestion('Warning','Are you sure you want to overwrite the file?', icon = 'warning')
	if MsgBox == 'Yes':
		try:
			f.close()
		except:
			save_as_function()
	if MsgBox == 'No':
		messagebox.showinfo('No','You will now return to the application screen')

# Function that is needed for "save_menu_bar"
def write_to_file(fileToWrite):
	fileToWrite = None
	try:
		fromFirstLineContent = content_text.get(1.0, END)
		with open(file_name, 'w') as file:
			file.write(fromFirstLineContent)
	except IOError:
		tkinter.messagebox.showwarning("Save", "Could not save the file.")

def save_menu_bar(event=None):
	fileToWrite = None
	if not file_name:
		save_as_function()
	else:
		write_to_file(fileToWrite)
	return "break"


def sendBig():
	'''
	Below is a big function which sends e-mail with proper title, body and attachments.
	'''
	class Subject:
		weekDictionary = {	1: 'wtorek',
							2: 'środę',
							3: 'czwartek',
							4: 'piątek',
							5: 'niedzielę',
							6: 'niedzielę',
							7: 'niedzielę',}
		def __init__(self, weekday):
			self.weekDay = weekday

		def createSubject(self):
			return f"Fakturki na {Subject.weekDictionary.get(self.weekday)}"

	subject = Subject(date.today().isoweekday())
	msg = MIMEMultipart()
	msg['Subject'] = subject.createSubject()

	class Body:
		def __init__(self, contentToSend, attachmentsDirectory):
			self.contentToSend = contentToSend
			self.attachmentsDirectory = attachmentsDirectory

		def createBody(self):
			mailBody = '<font face="verdana, monospace">' + "Cześć tatuś, oto fakturki :)" + \
					  '</font>' +  "<br>"
			with open(self.contentToSend, encoding='utf-8') as file:
				content = csv.reader(file, delimiter =';')

				for row in content:
					if row[0] != 'sent':
						mailBody += '<font face="verdana, monospace">' + "<br>" + \
						"   - " + '<b>' + row[1] + ' '+ '</b>' + row[2] + ' ' + \
						row[3] + ' ' + row[4] + '<b>' + '<font color="MidnightBlue">' + \
						row[5] + ' ' + '</font>' + '<font color="MediumVioletRed">' + \
						row [6] + ' ' +  '<font color="OrangeRed">' + row[7] + ' ' + \
						'</font>' + '</font>' + '</b> <br>'

						# row[2] is the invoice number
						InvoiceNumber = '*' + str.lstrip(row[2]) + '*' # removes spaces to the left of the row[2]
						for doc in os.listdir(self.attachmentsDirectory):
							if fnmatch.fnmatch(doc, InvoiceNumber):
								attachmentPath = os.path.join(self.attachmentsDirectory, doc)
								openedAttachmentPath = open(attachmentPath, 'rb')
								part = MIMEBase('application', 'octet-stream')
								part.set_payload(openedAttachmentPath.read())
								openedAttachmentPath.close()
								encoders.encode_base64(part)
								part.add_header('Content-Disposition',
												'attachment',
												filename = os.path.basename(doc))
								msg.attach(part)

			message = mailBody
			msg.attach(MIMEText(message, 'html'))


	body = Body(contentToSend,
				attachmentsDirectory)
	body.createBody()

	class Sending(Body):
		msg['From'] = sensitiveData.iloc[2][1] # sender's email
		msg['To'] = sensitiveData.iloc[8][1] # receiver's email
		def __init__(self, password, file_to_remove, *args, **kwargs):
			super().__init__(*args, **kwargs)
			self.password = password # password for sender's email
			self.file_to_remove = file_to_remove # abstract file which is created and removed automatically

		def send_email(self):
			server = smtplib.SMTP_SSL(sensitiveData.iloc[3][1], 465) #host, port
			server.ehlo()
			server.login(msg['From'], self.password)
			server.sendmail(msg['From'], msg['To'], msg.as_string())
			server.quit()
			shutil.copyfile(self.contentToSend, self.file_to_remove)

	data_to_send = Sending(sensitiveData.iloc[4][1],
						   'file_to_remove.csv',
						   contentToSend,
						   None)
	data_to_send.send_email()

	class Removing:
		def __init__(self, fv_dest_path, contentToSend, file_to_remove):
			self.fv_dest_path = fv_dest_path # path for archival directory
			self.contentToSend = contentToSend # path for CSV file where data to sent has been imported
			self.file_to_remove = file_to_remove # abstract file which is created and removed automatically

		def remove_and_close(self):
			csvfile = open(self.contentToSend, 'r', encoding='utf-8')
			csv_file = csv.reader(csvfile, delimiter=';')
			ofile = open(self.file_to_remove, 'w', encoding='utf-8', newline='')
			writer = csv.writer(ofile,  delimiter=';')
			for row in csv_file:
				row[0]='sent'
				writer.writerow(row)
			csvfile.close()
			ofile.close()
			os.remove(self.contentToSend)
			os.rename(self.file_to_remove, self.contentToSend)

	removing = Removing(sensitiveData.iloc[5][1],
						contentToSend,
						'file_to_remove.csv')
	removing.remove_and_close()

	class Moving(Removing):
		def __init__(self, attachmentsDirectory, *args, **kwargs):
			self.attachmentsDirectory = attachmentsDirectory
			super().__init__(*args, **kwargs)

		def move(self):
			with open(self.contentToSend, encoding='utf-8') as csvfile:
				csv_file = csv.reader(csvfile, delimiter =';')
				for item in csv_file:
					InvoiceNumber = '*' + str.lstrip(item[2]) + '*'
					for file in os.listdir(self.attachmentsDirectory):
						if fnmatch.fnmatch(file, InvoiceNumber):
							if '.DS_Store' in self.fv_dest_path:
									os.remove(self.fv_dest_path + '.DS_Store')
							shutil.move(self.attachmentsDirectory + file, self.fv_dest_path)

	moving = Moving(attachmentsDirectory,
					sensitiveData.iloc[5][1],
					contentToSend,
					None)
	moving.move()

	class Preparing():
		def __init__(self, file_to_edit):
			self.file_to_edit = file_to_edit
		def prepare(self):
			with open(self.file_to_edit, 'a') as txtfile:
				txtfile.write('---------------------------------')
	preparing = Preparing(fileToOpen)
	preparing.prepare()


	os.remove(contentToSend) #removes file to save memory

def on_find(self):
		t2 = Toplevel(root)
		t2.title('Find')
		t2.geometry('400x100+200+250')
		t2.transient(root)
		Label(t2, text="Find All:").grid(row=0, column=0, sticky='e')
		v=StringVar()
		e = Entry(t2, width=25, textvariable=v)
		e.grid(row=0, column=1, padx=2, pady=2, sticky='we')
		e.focus_set()
		c=IntVar()
		Checkbutton(t2, text='Ignore Case', variable=c).grid(row=1, column=1, sticky='e', padx=2, pady=2)
		Button(t2, text="Find All", underline=0,  command=lambda: search_for(v.get(),c.get(), textPad, t2,e)).grid(row=0, column=2, sticky='e'+'w', padx=2, pady=2)
		#t2.bind('<Return>', lambda: search_for(v.get(), c.get(), textPad, t2,e))
		def close_search():
				textPad.tag_remove('match', '1.0', END)
				t2.destroy()
		t2.protocol('WM_DELETE_WINDOW', close_search)

def search_for(needle,cssnstv, textPad, t2,e) :
		textPad.tag_remove('match', '1.0', END)
		count =0
		if needle:
				pos = '1.0'
				while True:
					pos = textPad.search(needle, pos, nocase=cssnstv, stopindex=END)
					if not pos: break
					lastpos = '%s+%dc' % (pos, len(needle))
					textPad.tag_add('match', pos, lastpos)
					count += 1
					pos = lastpos
				textPad.tag_config('match', foreground='white', background='deepskyblue3')
		e.focus_set()
		t2.title('%d matches found' %count)



menubar = Menu(root)
# File menu
filemenu = Menu(menubar, tearoff=0 )
filemenu.add_command(label = "New (clear screen)", accelerator = 'Cmd+N', compound = LEFT, image = newMenuIcon, underline = 0, command = new_file)
filemenu.add_command(label = "Save", accelerator = 'Cmd+S',compound = LEFT, image = saveMenuIcon,underline = 0, command=saveBig)
filemenu.add_command(label = "Save as",accelerator = 'Shift+Ctrl+S', compound = LEFT, image = saveAsMenuIcon,underline = 0, command=save_as_function)
filemenu.add_command(label = "Export to CSV",accelerator = 'Ctrl+E', compound = LEFT, image = exportMenuIcon,underline = 0, command=exportBig)
filemenu.add_separator()
filemenu.add_command(label = "Exit", accelerator = 'Alt+F4', compound = LEFT, image = exitMenuIcon,underline = 0, command = exit_editor)
menubar.add_cascade(label = "File", menu = filemenu) # all file menu choices will be placed here

# Edit menu
editmenu = Menu(menubar, tearoff = 0)
editmenu.add_command(label="Undo",compound=LEFT,  image=undoMenuIcon, accelerator='Cmd+Z', command = undo)
editmenu.add_command(label="Redo",compound=LEFT,  image=redoMenuIcon, accelerator='Cmd+Y', command = redo)
editmenu.add_separator()
editmenu.add_command(label="Cut", compound=LEFT, image=cutMenuIcon, accelerator='Cmd+X', command = cut)
editmenu.add_command(label="Copy", compound=LEFT, image=copyMenuIcon,  accelerator='Cmd+C', command = copy)
editmenu.add_command(label="Paste",compound=LEFT, image=pasteMenuIcon, accelerator='Cmd+V', command = paste)
editmenu.add_separator()
editmenu.add_command(label="Find", compound=LEFT, image=findMenuIcon, underline= 0, accelerator='Cmd+F', command=on_find)
editmenu.add_separator()
menubar.add_cascade(label = "Edit ", menu=editmenu)

root.config(menu = menubar) # this line displats menu on the top of the root window

lnlabel = Label(root,  width=2,  bg = 'white') # is responsible for white strip on the left
lnlabel.pack(side=LEFT, anchor='nw', fill=Y)

# adding Text Widget & ScrollBar widget
textPad = Text(root, undo=True)
textPad.pack(expand=YES, fill=BOTH)
scroll=Scrollbar(textPad)
textPad.configure(yscrollcommand=scroll.set)
scroll.config(command=textPad.yview)
scroll.pack(side=RIGHT,fill=Y)

# binding events
root.bind('<Command-f>', on_find)
root.bind('<Command-F>', on_find)
root.bind('<Command-N>', new_file)
root.bind('<Command-n>', new_file)
root.bind('<Command-S>', saveBig)
root.bind('<Command-s>', saveBig)


# toolbar
shortcutbar = Frame(root, height=25, bg='gainsboro')
icons = ['openBig', 'saveBig', 'exportBig', 'sendBig']
for i, icon in enumerate(icons):
	tbicon = PhotoImage(file='pictures/'+icon+'.gif')
	cmd = eval(icon)
	toolbar = Button(shortcutbar, image=tbicon, command=cmd, activeforeground='dodgerblue3')
	toolbar.image = tbicon
	toolbar.pack(padx=70, pady=5, side=LEFT)
shortcutbar.pack(expand=NO, fill=X)

root.mainloop()
