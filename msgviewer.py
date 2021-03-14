import tempfile
import tkinter as tk 
import extract_msg 
import os
import sys
from email import generator
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from tkinterhtml import HtmlFrame

class Email(object):
    def __init__(self, msg, filename):
        self.__msg = msg
        self.__filename = filename
        self.__createEmail()

    def __createEmail(self):
        sender = self.__msg.sender
        recepient = self.__msg.to
        subject = self.__msg.subject

        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = sender
        msg['To'] = recepient
        
        html = self.__msg.htmlBody if not self.__msg.htmlBody is None else self.__msg.body
        try:
            html = html.decode('utf-8')
        except UnicodeDecodeError as e:
            html = html.decode('iso-8859-1')

        part = MIMEText(html, 'html')
        msg.attach(part)
        outfile_name = self.__saveToFile(msg)

    def open(self):
        filename = self.__eml_file_name
        try:
            os.system(f'open "{filename}" -a "Microsoft Outlook.app" ')
        except:
            os.system(f'open "{filename}" -a "Mail.app" ')


    def __saveToFile(self,msg):
        old_name = self.__filename
        pre, ext = os.path.splitext(old_name)
        outfile_name = pre + ".eml"
        with open(outfile_name, 'w') as outfile:
            gen = generator.Generator(outfile)
            gen.flatten(msg)
        self.__eml_file_name = outfile_name


class MsgReader(object):
    def __init__(self, argv):
        self.file_name = None
        for a in argv:
            if a.endswith(".msg"):
                self.file_name = a
        
        self.root = tk.Tk() 
        T = HtmlFrame(self.root, horizontal_scrollbar="auto", fontscale=1.2)
        T.grid(row=1, column=0)
          
        l = tk.Label(self.root, text = "Message Viewer")
        l.grid(row=0, column=0) 
        l.config(font =("Courier", 16)) 
        self.msg = extract_msg.Message(self.file_name)

        self.attachments = {}
          
        self.listBox = tk.Listbox(self.root, width=80)
        self.listBox.grid(row=2,column=0)
        for i, a in enumerate(self.msg.attachments):
            self.listBox.insert(i, a.longFilename)
            self.attachments[i] = a

        self.listBox.bind('<Double-Button>', self.__list_double_click)

        b1 = tk.Button(self.root, text = "Reply", )
        b1.grid(row=3, column=0)
        b1.bind('<Button>', self.__reply)
          
        b2 = tk.Button(self.root, text = "Exit", 
                    command = self.root.destroy)  
        b2.grid(row=4, column=0) 

        e1 = tk.Entry(self.root)
        e2 = tk.Entry(self.root)
        e3 = tk.Entry(self.root)
        e4 = tk.Entry(self.root)
        e5 = tk.Entry(self.root)

        html = self.msg.htmlBody if not self.msg.htmlBody is None else msg.body
        try:
            html = html.decode('utf-8')
        except UnicodeDecodeError as e:
            html = html.decode('iso-8859-1')
        T.set_content(html)
        tk.mainloop() 

    def __reply(self, arg):
        email = Email(self.msg, self.file_name)
        email.open()

    def __list_double_click(self, arg):
        i, *y = self.listBox.curselection()
        fn = self.attachments[i].longFilename
        temp_dir = tempfile.TemporaryDirectory().name
        os.mkdir(temp_dir)
        file_path = f"{temp_dir}/{fn}"
        with open(file_path, "wb") as f:
            f.write(self.attachments[i].data)
        os.system(f'open  "{file_path}"')




