import extract_msg 
import os
import sys
from email import generator
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import tempfile

from email.mime.application import MIMEApplication

class Email(object):
    def __init__(self, msg, filename):
        self.__msg = msg
        self.__filename = filename
        self.__createEmail()

    def __createEmail(self):
        sender = self.__msg.sender
        recepient = self.__msg.to
        subject = self.__msg.subject

        msg = MIMEMultipart('mixed')
        msg['Subject'] = subject
        msg['From'] = sender
        msg['To'] = recepient
        for a in self.__msg.attachments:
            if not isinstance(a.data, extract_msg.Message):
                msg.attach(MIMEApplication(a.data, Name=a.longFilename))
            else:
                m = a.data
                for b in m.attachments:
                    msg.attach(MIMEApplication(b.data, Name=b.longFilename))
        
        html = self.__msg.htmlBody if not self.__msg.htmlBody is None else self.__msg.body
        try:
            if not isinstance(html,str):
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
        old_name = os.path.basename(old_name)
        pre, ext = os.path.splitext(old_name)
        outfile_name = pre + ".eml"
        temp_dir = tempfile.TemporaryDirectory().name
        os.mkdir(temp_dir)
        outfile_path = temp_dir + "/" + outfile_name 
        with open(outfile_path, 'w') as outfile:
            gen = generator.Generator(outfile)
            gen.flatten(msg)
        self.__eml_file_name = outfile_path


class OutlookPreview(object):
    def __init__(self, argv):
        file_name = None
        for a in argv:
            if a.endswith(".msg"):
                file_name = a
        
        msg = extract_msg.Message(file_name)
        email = Email(msg, file_name)
        email.open()


OutlookPreview(sys.argv)
