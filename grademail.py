#! /usr/bin/env python

# Import smtplib for the actual sending function
import smtplib
import re
import os
import time
import mimetypes
import getpass
import platform
import csv

from optparse import OptionParser

from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

DEBUG = False

CURRDIR = './Output'
COMMASPACE = ', '

# the prefix and suffix for the code files
CODE_P = 'Grade_'
CODE_S = '.pdf'

# the prefix and suffix for the code files
GRADE_P = 'Grade_'
GRADE_S = '.txt'

#time to sleep between opening files
SLEEP_TIME = 0.1

# the labnumber
labnum = ""

def main():
     parser = OptionParser("""\
Send the canned response emails

Usage: ./grademail.py -e [email] -t [text] -l [labnumber]

Note that above options are not optional.
""")
     parser.add_option('-e', '--email',
                      type='string', action='store',
                      help="""The email adress of the sender""")
     parser.add_option('-t', '--text',
                      type='string', action='store',
                      help="""The text of the email""")
     parser.add_option('-l', '--labnum',
                      type='string', action='store',
                      help="""The two-digit lab number""")
     parser.add_option('-D', '--DEBUG', action='store_true',
                      help= """Enable if you want to debug.""")

     opts, args = parser.parse_args()

     # python -m smtpd -n -c DebuggingServer localhost:1025
     # enter the above line at terminal to run debugging server
     DEBUG = opts.DEBUG
     if DEBUG:
          print("DEBUG ENABLED")
          PORT = 1025
          SERVER = 'localhost'
     else:
          print("NO DEBUG")
          PORT = 587
          SERVER = "smtp.gmail.com:587"

     # check for incomplete arguments
     if not opts.email or not opts.text or not opts.labnum:
          raise IncompleteArgumentsException();

     usr = opts.email
     pw = getpass.getpass('Password for %s:' % usr)

     # create connection to the server
     server = connect(usr, pw)
     sendEmails(__parseCSV(opts.labnum), opts.text, usr, server, opts.labnum)
     server.quit()

def sendEmails(mapping, text, sender, server, labnum):
     for filename in mapping:
          rubric = "%s/%s%s" % (CURRDIR, filename, GRADE_S)
          comments = "%s/%s%s" % (CURRDIR, filename, CODE_S)
          if os.path.isfile(rubric) and os.path.isfile(comments):
               print("Send mail to %s" % (filename))
               sendmail(getName(filename), text, sender, mapping[filename], server, labnum)
          else:
               print("Missing .txt or .pdf file for %s.*" % (filename))

# create connection to the server
def connect(usr, pw):
     if DEBUG: # localhost
          return smtplib.SMTP(SERVER,PORT)
     else: # log in to external mail server
          server = smtplib.SMTP(SERVER)
          server.ehlo()
          server.starttls()
          server.login(usr, pw)
          return server


def sendmail(name, textbody, sender, receiver, server, num):
     # Open a plain text file for reading. For this example, assume that
     # the text file contains only ASCII characters

     msg = MIMEMultipart("Hello, this is your grading email!")

     # create the proper email
     msg['Subject'] = 'Lab %s grade' % num
     msg['From'] = sender
     msg['To'] = receiver
     msg.preamble = 'Grading email \n'

     bodyfp = open(textbody, 'rb')
     text = getBody(bodyfp, num)
     msg.attach(text)
     bodyfp.close()

     # get the files
     codeFile  = findPair(CODE_P,  name)
     gradeFile = findPair(GRADE_P, name)
     codePath  = os.path.join(CURRDIR, codeFile)
     gradePath = os.path.join(CURRDIR, gradeFile)

     # check for errors
     if not os.path.isfile(codePath):
          raise BadPathException(codePath)
     if not os.path.isfile(gradePath):
          raise BadPathException(gradePath)

     # add the attachments
     # ASSUME GRADE IS .txt AND CODE IS .pdf
     # attach the .txt
     ctype_g, encoding_g = mimetypes.guess_type(gradePath)
     if ctype_g is None or encoding_g is not None:
          # No guess could be made, or the file is encoded (compressed), so
          # use a generic bag-of-bits type.
          ctype_g = 'text/plain'
     maintype_g, subtype_g = ctype_g.split('/', 1)
     gradefp = open(gradePath)
     # Note: we should handle calculating the charset
     grade = MIMEText(gradefp.read(),_subtype=subtype_g)
     gradefp.close()
     grade.add_header('Content-Disposition', 'attachment', filename=gradeFile)
     msg.attach(grade)

     # attach the .pdf
     if not DEBUG:
          ctype_c, encoding_c = mimetypes.guess_type(codePath)
          if ctype_c is None or encoding_c is not None:
               # No guess could be made, or the file is encoded (compressed), so
              # use a generic bag-of-bits type.
               ctype_c = 'application/pdf'
          maintype_c, subtype_c = ctype_c.split('/', 1)
          codefp = open(codePath, 'rb')
          # Note: we should handle calculating the charset
          code = MIMEBase(maintype_c, _subtype=subtype_c)
          code.set_payload(codefp.read())
          encoders.encode_base64(code)
          codefp.close()
          code.add_header('Content-Disposition', 'attachment', filename=codeFile)
          msg.attach(code)

     # send the message
     server.sendmail(sender, [receiver], msg.as_string())


def __parseCSV(num):
     csvfd = open('./Emails.csv', 'r')
     dictionary = {}
     with csvfd as emails:
          reader = csv.DictReader(emails)
          for row in reader:
               dictionary[__nameToFile(row['Name'], num)] = row['Email']
     return dictionary

def __nameToFile(name, num):
     first = __getFirst(name)
     last = __getLast(name)
     return "Grade_Lab%s%s%s" % (num,last,first)

def __getFirst(name):
     brkIdx = name.index(" ")
     return name[:brkIdx]

def __getLast(name):
     brkIdx = -name[::-1].index(" ")
     return name[brkIdx:]

# a function to open all the grade files
def openGradeFiles():
     __openFiles(GRADE_P)

# function to open all the Code files
def openCodeFiles():
     __openFiles(CODE_P)

# private helper function to open up files with
# a given prefix.
def __openFiles(prefix):
     for filename in os.listdir(CURRDIR):
          if re.match(prefix, filename) != None:
               time.sleep(SLEEP_TIME)
               command = ""
               if platform.system() == "Windows":
                    command = "start "
               else:
                    command = "open "
               os.system("%s%s/%s " % (command,CURRDIR,filename));


# search file for code/grade pairs
def findPair (prefix, name):
     regex = prefix+name
     pattern = re.compile(regex)
     for filename in os.listdir(CURRDIR):
          match = pattern.match(filename)
          if match != None:
               return filename

# retrieve the name from the file name
def getName (filename):
     if filename[:4] == 'Code':
          return filename[5:-4]
     elif filename[:5] == 'Grade':
          return filename[6:-4]
     else:
          raise ImproperNameException(filename)

def getBody(fd, num):

     nextline = fd.readline()
     content = "";
     while nextline:
          print(type(nextline.decode('utf-8')))
          content = "%s%s" % (content, nextline.decode('utf-8'))
          nextline = fd.readline()

     # content = content.replace('XX', num)
     text = MIMEText(content, _subtype='plain')
     return text

def isCodeFile(filename):
     return re.match(CODE_P, filename) != None

def isGradeFile(filename):
     return re.match(GRADE_P, filename) != None

# exception when naming convention has not been adhered to
class ImproperNameException(Exception):
     def __init__(self, value):
          self.value = value
          print('ImproperNameException: %s not properly named' % value)
     def __str__(self):
          return repr(self.value)

# exception for when the options are incomplete
class IncompleteArgumentsException(Exception):
     def __init__(self):
          self.value = "An option was omitted."
     def __str__(self):
          return repr(self.value)

# exception if there is no file at path
class BadPathException(Exception):
     def __init__(self, value):
          self.value = value
     def __str__(self):
          return repr(self.value)

# execute main
if __name__ == "__main__": main()
