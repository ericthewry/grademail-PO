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
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# the prefix and suffix for the Commented Code
CODE_P = 'Grade_'
CODE_S = '.pdf'

# the prefix and suffix for the Rubric files
GRADE_P = 'Grade_'
GRADE_S = '.txt'

def main():
     parser = OptionParser("""\
./grademail.py -e [email] -t [text] -l [labnumber]
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
                      help="""Enable if you want to debug.""")
     parser.add_option('-d', '--dir',
                      type='string', action='store',
                      help="""Specify Grading Dir, default is current dir""")

     opts, args = parser.parse_args()

     # check for incomplete arguments
     if not opts.email or not opts.text or not opts.labnum:
          raise IncompleteArgumentsException();

     curDir = opts.dir if opts.dir != None else '.'
     debug = opts.DEBUG
     usr = opts.email
     pw = getpass.getpass('Password for %s:' % usr)

     # create connection to the server
     server = connect(usr, pw, debug)
     sendEmails(parseCSV(opts.labnum, curDir), opts.text, usr, server, opts.labnum, curDir)
     server.quit()

def sendEmails(mapping, text, sender, server, labnum, curDir):
     for filename in mapping:
          rubric = "%s/%s%s" % (curDir, filename, GRADE_S)
          comments = "%s/%s%s" % (curDir, filename, CODE_S)
          if os.path.isfile(rubric) and os.path.isfile(comments):
               # print("Send mail to %s" % (filename))
               sendmail(filename, text, sender, mapping[filename], server, labnum, curDir)
          else:
               print("Missing .txt or .pdf file for %s.*" % (filename))

# create connection to the server
def connect(usr, pw, debug):
     if debug: # localhost
          print("DEBUG: Using Local Server")
          PORT = 1025
          SERVER = 'localhost'
          return smtplib.SMTP(SERVER,PORT)
     else: # log in to gmail smtp server
          print("Connecting to mail server...")
          PORT = 587
          SERVER = "smtp.gmail.com:587"
          server = smtplib.SMTP(SERVER)
          server.ehlo()
          server.starttls()
          server.login(usr, pw)
          print("Connected!")
          return server

def sendmail(name, textbody, sender, receiver, server, num, curDir):
     # Open a plain text file for reading. For this example, assume that
     # the text file contains only ASCII characters
     msg = newMessage(sender, receiver, num)
     addCannedText(msg, curDir, textbody, num)

     # get the files
     codeFile  = name + CODE_S
     gradeFile = name + GRADE_S
     codePath  = os.path.join(curDir, codeFile)
     gradePath = os.path.join(curDir, gradeFile)

     # check for errors
     if not os.path.isfile(codePath):
          raise BadPathException(codePath)
     if not os.path.isfile(gradePath):
          raise BadPathException(gradePath)

     # add the attachments
     # attach the .txt rubric
     attachRubric(gradePath, gradeFile, msg)
     attachCommentedCode(codePath, codeFile, msg)

     # send the message
     server.sendmail(sender, [receiver], msg.as_string())

## MESSAGE METHODS ##

def newMessage(sender, receiver, num):
     # create the proper email
     msg = MIMEMultipart("")
     msg['Subject'] = 'Lab %s grade' % num
     msg['From'] = sender
     msg['To'] = receiver
     msg.preamble = 'Grading email \n'
     return msg

def addCannedText(msg, curDir, text, num):
     bodyfp = open("%s/%s" % (curDir,text), 'rb')
     text = getBody(bodyfp, num)
     msg.attach(text)
     bodyfp.close()

def attachRubric(gradePath, gradeFile, msg):
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

def attachCommentedCode(codePath, codeFile, msg):
     # attach the .pdf
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

## CSV HELPERS ##

def parseCSV(num, curDir):
     csvfd = open(curDir + '/Emails.csv', 'r')
     dictionary = {}
     with csvfd as emails:
          reader = csv.DictReader(emails)
          for row in reader:
               dictionary[nameToFile(row['Name'], num)] = row['Email']
     return dictionary

def nameToFile(name, num):
     first = getFirst(name)
     last = getLast(name)
     return "Grade_Lab%s%s%s" % (num,last,first)

def getFirst(name):
     brkIdx = name.index(" ")
     return name[:brkIdx]

def getLast(name):
     brkIdx = -name[::-1].index(" ")
     return name[brkIdx:]

def getBody(fd, num):
     nextline = fd.readline()
     content = "";
     while nextline:
          content = "%s%s" % (content, nextline.decode('utf-8'))
          nextline = fd.readline()
     # content = content.replace('XX', num)
     text = MIMEText(content, _subtype='plain')
     return text

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
