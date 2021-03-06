# -*- coding: utf-8 -*-
'''
The main core of this program are 3 classes:
* EditMenu
* FileMenu
* BottomToolbar.
EditMenu includes functions responsible for undo, redo, cut, copy, paste
and find operations, then FileMenu includes new file, exit and save as ones.
Some features from File menu (open, save, export to CSV, send via email)
are common with bottom toolbar, hence they have been set in BottomToolbar class.
'''

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

# Variables for sensitive data
sensitiveData = pd.read_excel('private.xlsx')
attachmentsDirectory = sensitiveData.iloc[1][1]
senderMailAddress = sensitiveData.iloc[2][1]
senderMailPassword = sensitiveData.iloc[4][1]
hostofSenderMail = sensitiveData.iloc[3][1]
receiverMailAddress = sensitiveData.iloc[8][1]
fileToOpen = sensitiveData.iloc[6][1]
contentToSend = sensitiveData.iloc[7][1]
archiveForAttachmentsPath = sensitiveData.iloc[5][1]

# Images for icons
openButtonIcon = PhotoImage(file = 'pictures/' + 'openBig.gif')
saveButtonIcon = PhotoImage(file = 'pictures/' + 'saveBig.gif')
sendButtonIcon = PhotoImage(file = 'pictures/' + 'sendBig.gif')
exportButtonIcon = PhotoImage(file = 'pictures/' + 'exportBig.gif')
saveAsMenuIcon = PhotoImage(file = 'pictures/' + 'save_as.gif')
exitMenuIcon = PhotoImage(file = 'pictures/' + 'exit.gif')
undoMenuIcon = PhotoImage(file = 'pictures/' + 'undo.gif')
redoMenuIcon = PhotoImage(file = 'pictures/' + 'redo.gif')
cutMenuIcon = PhotoImage(file = 'pictures/' + 'cut.gif')
copyMenuIcon = PhotoImage(file = 'pictures/' + 'copy.gif')
pasteMenuIcon = PhotoImage(file = 'pictures/' + 'paste.gif')
findMenuIcon = PhotoImage(file = 'pictures/' + 'find.gif')
newMenuIcon = PhotoImage(file = 'pictures/' + 'new.gif')
sendMenuIcon = PhotoImage(file = 'pictures/' + 'send.gif')
exportMenuIcon = PhotoImage(file = 'pictures/' + 'exportToResize.gif').subsample(3,3) # resizing
saveMenuIcon = PhotoImage(file = 'pictures/' + 'saveToResize.gif').subsample(3,3) # resizing

menubar = Menu(root)
root.config(menu = menubar) # this line displays menu on the top of the root window

lnlabel = Label(root,  width=2,  bg = 'white') # is responsible for white strip on the left
lnlabel.pack(side=LEFT, anchor='nw', fill=Y)

# adding Text Widget & ScrollBar widget
textPad = Text(root, undo=True)
textPad.pack(expand=YES, fill=BOTH)
scroll=Scrollbar(textPad)
textPad.configure(yscrollcommand=scroll.set)
scroll.config(command=textPad.yview)
scroll.pack(side=RIGHT,fill=Y)

class EditMenu:
    def undo(self):
        textPad.event_generate("<<Undo>>")

    def redo(self):
        textPad.event_generate("<<Redo>>")

    def cut(self):
        textPad.event_generate("<<Cut>>")

    def copy(self):
        textPad.event_generate("<<Copy>>")

    def paste(self):
        textPad.event_generate("<<Paste>>")

    def search_for(needle,cssnstv, textPad, t2,e):
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
                textPad.tag_config('match', foreground='white',
                                   background='deepskyblue3')
            e.focus_set()
            t2.title('%d matches found' %count)

    def on_find(self):
        t2 = Toplevel(root)
        t2.title('Find')
        t2.geometry('400x100+200+250')
        t2.transient(root)
        Label(t2, text="Find All:").grid(row=0, column=0, sticky='e')
        v = StringVar()
        e = Entry(t2, width=25, textvariable=v)
        e.grid(row=0, column=1, padx=2, pady=2, sticky='we')
        e.focus_set()
        c = IntVar()
        Checkbutton(t2, text='Ignore Case', variable=c).grid(row=1,
                              column=1, sticky='e', padx=2, pady=2)
        Button(t2, text="Find All", underline=0,
               command=lambda: EditMenu.search_for(v.get(),c.get(), textPad, t2,e)).grid(
               row=0, column=2, sticky='e'+'w', padx=2, pady=2)

        def close_search():
            textPad.tag_remove('match', '1.0', END)
            t2.destroy()

        t2.protocol('WM_DELETE_WINDOW', close_search)

editMenuObject = EditMenu()

class FileMenu:
    def new_file(self):
        root.title("Untitled")
        textPad.delete(1.0,END)

    def exit_editor(event=None):
        if messagebox.askokcancel("Quit?", "Do you really want to exit?",
                                  icon = 'warning'):
            root.destroy()

    def save_as_function(self):
        fileToSave = filedialog.asksaveasfile(mode='w', defaultextension='.txt')
        wholeContent = textPad.get(0.0, END)
        fileToSave.write(wholeContent)
        fileToSave.close()

fileMenuObject = FileMenu()

class BottomToolbar:
    def openBig(self):
        fileContent = open(fileToOpen)
        for line in fileContent:
            textPad.insert(END, line)

    def saveBig(self):
        openedFile = open(fileToOpen, 'w')
        fromFirstLineContent = textPad.get(1.0, END)
        openedFile.write(fromFirstLineContent)
        MsgBox = messagebox.askquestion('Warning',
              'Are you sure you want to overwrite the file?', icon = 'warning')
        if MsgBox == 'Yes':
            try:
                f.close()
            except:
                save_as_function()
        if MsgBox == 'No':
            messagebox.showinfo('No','You will now return to the application screen')

    def exportBig(self):
        '''
        Info:
        Something like '---------------------------------' separates content
        to send from content that has been send last time.
        Below variable "separatingIndex" is set of rows
        that equal to '---------------------------------'.
        Thanks to it program can distinguish last new content that must be send.
        '''
        root.clipboard_clear()
        with open(fileToOpen, encoding='utf-8') as file:
            fileContent = file.readlines()
            separatingIndex = [x for x in range(len(fileContent)) if '---------------------------------' in fileContent[x]]
        fileToExport = filedialog.asksaveasfile(initialfile='export.csv',
                                                mode='w', defaultextension='.csv')
        # below: asksaveasfile returns `None` if dialog closed with "cancel".
        if fileToExport is None:
            return
        newContent = textPad.get(float(separatingIndex[-1]+2), END)
        fileToExport.write(newContent)
        fileToExport.close()

    def sendBig(self):
        '''
        Below is a big function which sends e-mail with proper title,
        body and attachments.
        '''
        class Subject:
            weekDictionary = {  1: 'wtorek',
                                2: 'środę',
                                3: 'czwartek',
                                4: 'piątek',
                                5: 'niedzielę',
                                6: 'niedzielę',
                                7: 'niedzielę',}
            def __init__(self, weekday):
                self.weekday = weekday

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
                mailBody = '<font face="verdana, monospace">' + \
                           "Cześć tatuś, oto fakturki :)" + \
                           '</font>' +  "<br>"
                with open(self.contentToSend, encoding='utf-8') as file:
                    content = csv.reader(file, delimiter =';')

                    for row in content:
                        if row[0] != 'sent':
                            mailBody += '<font face="verdana, monospace">' + \
                            "<br>" + "   - " + '<b>' + row[1] + ' '+ '</b>' + \
                            row[2] + ' ' + row[3] + ' ' + row[4] + '<b>' + \
                            '<font color="MidnightBlue">' + row[5] + ' ' + \
                            '</font>' + '<font color="MediumVioletRed">' + \
                            row [6] + ' ' +  '<font color="OrangeRed">' + \
                            row[7] + ' ' + '</font>' + '</font>' + '</b> <br>'

                            # row[2] is the invoice number
                            # below removes spaces to the left of the row[2]
                            InvoiceNumber = '*' + str.lstrip(row[2]) + '*'
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
            msg['From'] = senderMailAddress
            msg['To'] = receiverMailAddress
            def __init__(self, senderMailPassword, fileToAutoRemove,
                         *args, **kwargs):
                super().__init__(*args, **kwargs)
                self.senderMailPassword = senderMailPassword
                # below is abstract file which is created and
                # removed automatically (see Removing class)
                self.fileToAutoRemove = fileToAutoRemove

            def sendMail(self):
                # 465 is a port for sender's mail
                server = smtplib.SMTP_SSL(hostofSenderMail, 465)
                server.ehlo()
                server.login(msg['From'], self.senderMailPassword)
                server.sendmail(msg['From'], msg['To'], msg.as_string())
                server.quit()
                shutil.copyfile(self.contentToSend, self.fileToAutoRemove)

        sendingObject = Sending(senderMailPassword,
                              'fileToAutoRemove.csv',
                               contentToSend,
                               None)
        sendingObject.sendMail()

        class Removing:
            def __init__(self, archiveForAttachmentsPath, contentToSend,
                         fileToAutoRemove):
                self.archiveForAttachmentsPath = archiveForAttachmentsPath
                self.contentToSend = contentToSend
                self.fileToAutoRemove = fileToAutoRemove

            def removeAbstractFile(self):
                openedContentToSend = open(self.contentToSend, 'r',
                                           encoding='utf-8')
                readContent = csv.reader(openedContentToSend, delimiter=';')
                openedFileToAutoRemove = open(self.fileToAutoRemove, 'w',
                                              encoding='utf-8', newline='')
                writer = csv.writer(openedFileToAutoRemove,  delimiter=';')
                for row in readContent:
                    row[0]='sent'
                    writer.writerow(row)
                openedContentToSend.close()
                openedFileToAutoRemove.close()
                os.remove(self.contentToSend)
                os.rename(self.fileToAutoRemove, self.contentToSend)

        removing = Removing(archiveForAttachmentsPath,
                            contentToSend,
                            'fileToAutoRemove.csv')
        removing.removeAbstractFile()

        class Moving(Removing):
            def __init__(self, attachmentsDirectory, *args, **kwargs):
                self.attachmentsDirectory = attachmentsDirectory
                super().__init__(*args, **kwargs)

            def move(self):
                with open(self.contentToSend,
                encoding='utf-8') as openedContentToSend:
                    readContent = csv.reader(openedContentToSend,
                                             delimiter =';')
                    for row in readContent:
                        InvoiceNumber = '*' + str.lstrip(row[2]) + '*'
                        for doc in os.listdir(self.attachmentsDirectory):
                            if fnmatch.fnmatch(doc, InvoiceNumber):
                                if '.DS_Store' in self.archiveForAttachmentsPath:
                                        os.remove(self.archiveForAttachmentsPath +
                                                  '.DS_Store')
                                shutil.move(self.attachmentsDirectory + doc,
                                            self.archiveForAttachmentsPath)

        moving = Moving(attachmentsDirectory,
                        archiveForAttachmentsPath,
                        contentToSend,
                        None)
        moving.move()

        class Preparing():
            def __init__(self, fileToOpen):
                self.fileToOpen = fileToOpen

            def prepare(self):
                with open(self.fileToOpen, 'a') as file:
                    file.write('---------------------------------')
        preparing = Preparing(fileToOpen)
        preparing.prepare()

        os.remove(contentToSend) #removes file to save memory

bottomToolbarObject = BottomToolbar()


# File menu
filemenu = Menu(menubar, tearoff=0 )
filemenu.add_command(label = "New (clear screen)", accelerator = 'Cmd+N',
                     compound = LEFT, image = newMenuIcon, underline = 0,
                     command = fileMenuObject.new_file)
filemenu.add_command(label = "Save", accelerator = 'Cmd+S',
                     compound = LEFT, image = saveMenuIcon,underline = 0,
                     command=bottomToolbarObject.saveBig)
filemenu.add_command(label = "Save as",accelerator = 'Shift+Ctrl+S',
                     compound = LEFT, image = saveAsMenuIcon,underline = 0,
                     command=fileMenuObject.save_as_function)
filemenu.add_command(label = "Export to CSV",accelerator = 'Ctrl+E',
                     compound = LEFT, image = exportMenuIcon,underline = 0,
                     command=bottomToolbarObject.exportBig)
filemenu.add_separator()
filemenu.add_command(label = "Exit", accelerator = 'Alt+F4',
                     compound = LEFT, image = exitMenuIcon,underline = 0,
                     command = fileMenuObject.exit_editor)
menubar.add_cascade(label = "File", menu = filemenu)

# Edit menu
editmenu = Menu(menubar, tearoff = 0)
editmenu.add_command(label="Undo",compound=LEFT, image=undoMenuIcon,
                     accelerator='Cmd+Z', command = editMenuObject.undo)
editmenu.add_command(label="Redo",compound=LEFT, image=redoMenuIcon,
                     accelerator='Cmd+Y', command = editMenuObject.redo)
editmenu.add_separator()
editmenu.add_command(label="Cut", compound=LEFT, image=cutMenuIcon,
                     accelerator='Cmd+X', command = editMenuObject.cut)
editmenu.add_command(label="Copy", compound=LEFT, image=copyMenuIcon,
                     accelerator='Cmd+C', command = editMenuObject.copy)
editmenu.add_command(label="Paste",compound=LEFT, image=pasteMenuIcon,
                     accelerator='Cmd+V', command = editMenuObject.paste)
editmenu.add_separator()
editmenu.add_command(label="Find", compound=LEFT, image=findMenuIcon,
                     underline=0, accelerator='Cmd+F',
                     command=editMenuObject.on_find)
editmenu.add_separator()
menubar.add_cascade(label = "Edit ", menu=editmenu)


# Binding events
root.bind('<Command-f>', editMenuObject.on_find) # TODO: fix shorcut for searching
root.bind('<Command-F>', editMenuObject.on_find)
root.bind('<Command-N>', fileMenuObject.new_file)
root.bind('<Command-n>', fileMenuObject.new_file)
root.bind('<Command-S>', bottomToolbarObject.saveBig)
root.bind('<Command-s>', bottomToolbarObject.saveBig)

# toolbar
shortcutbar = Frame(root, height=25, bg='gainsboro')
icons = ['bottomToolbarObject.openBig',
         'bottomToolbarObject.saveBig',
         'bottomToolbarObject.exportBig',
         'bottomToolbarObject.sendBig']
for i, icon in enumerate(icons):
    pathForPictures = 'pictures/'+icon+'.gif'
    tbicon = PhotoImage(file=pathForPictures.replace('bottomToolbarObject.', ''))
    cmd = eval(icon)
    toolbar = Button(shortcutbar, image=tbicon, command=cmd,
              activeforeground='dodgerblue3')
    toolbar.image = tbicon
    toolbar.pack(padx=70, pady=5, side=LEFT)
shortcutbar.pack(expand=NO, fill=X)

root.mainloop()
