'''
Created on 20-Apr-2010

@author: Sunil
'''
# Imports operations
import os
import sys
from optparse import OptionParser
import smtplib
import parser
import email

# import MIME readers
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.mime.audio import MIMEAudio
from email.mime.application import MIMEApplication
from email.message import Message
import mimetypes
from email.errors import *
from email import encoders


COMMASPACE=', '

class Mailer:
    _host = 'localhost'
    _port = '25'
    _timeout = 10
    _message = None
    _from = None
    _to = []
    _user = None
    _passwd = None
    _dbuglevel = False
    
    def __init__(self,message=None, host=None, port=None):
        
        if host is not None:
            self._host = host
        
        if port is not None:
            self._port = port
             
        if message is None:
            self._message = MIMEMultipart()
        else:
            self.parse_message(message)
            
    def add_to(self,recpt, to_msg = True):
        self._to.append(recpt)
        if self._message["To"] is None:
            self._message["To"] = recpt
        else:
            self._message["To"] += (COMMASPACE+recpt)
        #print(self._message["To"])
    
    def set_from(self,sender, to_msg = True):
        self._from = sender
        self._message["From"] = sender
    
    def send(self):
        if self._from is None:
            print("[*]ERR: Sender Not specified")
            return
        if len(self._to) < 1:
            print("[*]Err: At-least one recipient is required")
            return
    
        try:
            client = smtplib.SMTP(self._host, self._port, self._timeout)
            print("Trying to send::\n" + self._message.as_string(False))
            client.set_debuglevel(True)
            client.sendmail(self._from, self._to, self._message.as_string(False))
            #if self._dbuglevel:
                #print("mail-data:\n"+self._message)
            print("SUCCESS:: Mail sent to server")
        except smtplib.SMTPConnectError as ex:
            print("!!Connection Error:: " + ex)
        except smtplib.SMTPHeloError as ex:
            print("!!Server did not respond properly:: " + ex)
        except smtplib.SMTPSenderRefused as ex:
            print("!!Sender Refused:: " + ex)
        except smtplib.SMTPRecipientsRefused as ex:
            print("!!Recipients Refused:: " + ex)
        except smtplib.SMTPResponseException as ex:
            print("!!Error:: Code:: "+ ex.smtp_code + " :: " + ex.smtp_error)
        except smtplib.SMTPServerDisconnected as ex:
            print("!!Server disconnected:: " + ex)
        finally:
            client.quit()
                 
    def attach(self,file):
        (ctype, encoding) = mimetypes.guess_type(file, strict=False)
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"
            
        (main_type, sub_type) = ctype.split('/',1)
        print("[*] Adding attachment of type: "+ ctype)
        if main_type == "text":
            fp = open(file)
            inner_msg = MIMEText(fp.read(), _subtype=sub_type)
            fp.close()
            
        elif main_type == "image":
            fp = open(file,'rb')
            inner_msg = MIMEImage(fp.read(), _subtype=sub_type)
            fp.close()
            
        elif main_type == "audio":
            fp = open(file,'rb')
            inner_msg = MIMEImage(fp.read(), _subtype = sub_type)
            fp.close()
        
        elif main_type == "application" and sub_type != "octet-stream":
            fp = open(file,'rb')
            inner_msg = MIMEApplication(fp.read, _subtype= sub_type)
            fp.close()
            
        else:
            fp = open(file,'rb')
            inner_msg = MIMEBase(main_type, sub_type)
            inner_msg.set_payload(fp.read)
            fp.close()
            encoders.encode_base64(inner_msg)
            
        self._message.add_header('Content-Disposition','attachment', filename = file)
        self._message.attach(inner_msg)
        
    def parse_message(self, message):
        try:
            fp = open(message)
            self._message = email.message_from_file(fp)
        except MessageError as msg_error:
            print("!! There was an error in parsing message::\n" + msg_error)
        finally:
            fp.close()
   
    def all_from_dir(self,dir):
        for file in os.listdir(dir):
            self.attach(file)
    
    def set_login(self,user,passwd=None):
        self._user = user
        self._passwd = passwd
    
    def set_subject(self,subject):
        self._message["Subject"] = subject
    
    def set_msg_data(self,data_file):
        pass
    
    def save_message(self,output):
        output.write(self._message.as_string(False))
    
if __name__ == '__main__' :
    
    opt_parser = OptionParser(usage="""\
Craft a email message from a list of attachments and send or output to file.

Usases: %prog [options]

Unless the host or -o option is given, the mail will be sent using mail server on "localhost".
Hence local machine must be running a mail server or -H HOST must be specified.

    """)
    
    opt_parser.add_option('-d','--directory',
                      type ='string', action='store', metavar='DIRECTORY',
                      help ="""Mail the contents of specified directory.
                      Only the regular files are mailed. No recursive directory listing"""
                      )
    opt_parser.add_option('-f','--file',
                      type='string', action='append', metavar='FILE',
                      default=[], dest='files',
                      help="""File to be sent as attachment with mail"""
                      )
    opt_parser.add_option('-m','--message',
                      type='string', action='store', metavar='FILE', default=None,
                      help="""Path to the pre-formatted email message to send.
                      """
                      )
    opt_parser.add_option('-H', '--host',
                      type='string', action='store', metavar='HOSTNAME', default=None,
                      help="""Hostname of the SMTP server to connect with. Default is 'localhost'"""
                      )
    opt_parser.add_option('-P', '--port',
                      type='int', metavar='PORT', default=None,
                      help="""SMTP Port of the server. Required if SMTP is not available on default port"""
                      )
    opt_parser.add_option('-l', '--login',
                      type='string', action="store", metavar='LOGIN:PASSWORD',
                      help="""USERNAME:PASSWORD if server requires authentication"""
                      )
    opt_parser.add_option('-s', '--sender',
                      type='string', action='store', metavar='MAIL-FROM',
                      help="""Sender of the email. used as FROM header of email envelop (Required) """
                      )
    opt_parser.add_option('-r', '--recipient',
                      type='string', action='append', metavar='RCPT-TO',
                      default=[], dest='recipients',
                      help="""Recipients of the mail. Used in RCPT TO header of email envelop
                      (at least 1 required)"""
                      )
    opt_parser.add_option('-o', '--output',
                      type='string', action='store', metavar='FILE',
                      help="""Output the composed Message to a file."""
                      )
    '''
        opt_parser.add_option('-d', '--data',
                      type='string', action='store', metavar="TEXT_FILE",
                      help="""A file containing additions fields for DATA block.
                      Must be in format of FIELD:Value1[, Value2, ...]."""
                      )
    '''
    opts, arg = opt_parser.parse_args()
    if not opts.sender or not opts.recipients:
        opt_parser.print_help()
        sys.exit(1)
    
    mailer = Mailer(opts.message, opts.host, opts.port)
    mailer.set_subject("Just testing the mail API")
    mailer.set_from(opts.sender)
    
    for recpt in opts.recipients:
        mailer.add_to(recpt)
    
    for file in opts.files:
        mailer.attach(file)
        
    if opts.login:
        user,passwd = opts.login.split(':',1)
        mailer.set_login(user, passwd)
    fp = open("C:\\RevEng\\mail.txt",'w')    
    mailer.save_message(fp)
    fp.close()
    # mailer.send()