#! /usr/bin/env python

## The purpose of this script is to automate the process of sending
## out grades via email to students.  This version specifically supports
## sending out a commented .pdf file and a rubric .txt file.
##
## Eric Campbell, 2015

import smtplib # access smtp server
import re # regex
import os # make system calls
import mimetypes # determine types of input
import getpass # read password
import csv # access csvs

# pass options via commandline
from optparse import OptionParser

# import email building
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

# Send the collection of emails from the mapping
# PARAMS:
#    Dictionary mapping -- a dictionary.  KEYS are names, VALUES are Emails
#    str text           -- the canned text file path
#    str sender         -- the sending email
#    SMTP server        -- the SMTP server being contacted
#    str labnum         -- the labnumber
#    str curDir         -- the path to the grade files
def sendEmails(mapping, text, sender, server, labnum, curDir):
     print(type(server))
     for filename in mapping:
          rubric = "%s/%s%s" % (curDir, filename, GRADE_S)
          comments = "%s/%s%s" % (curDir, filename, CODE_S)
          if os.path.isfile(rubric) and os.path.isfile(comments):
               # print("Send mail to %s" % (filename))
               sendmail(filename, text, sender, mapping[filename], server, labnum, curDir)
          else:
               print("Missing .txt or .pdf file for %s.*" % (filename))

# create connection to the server
# PARAMS:
#     str usr    -- the email address of the user
#     str pw     -- the users password
#     bool debug -- true if --DEBUG flag triggered
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

# Send a single email
# PARAMS:
#    str name           -- the name of the file (without extension)
#    str text           -- the canned text file path
#    str sender         -- the sending email
#    str receiver       -- the email of the recipient
#    SMTP server        -- the SMTP server being contacted
#    str num            -- the labnumber
#    str curDir         -- the path to the grade files
def sendmail(name, text, sender, receiver, server, num, curDir):
     # Open a plain text file for reading. For this example, assume that
     # the text file contains only ASCII characters
     msg = newMessage(sender, receiver, num)
     addCannedText(msg, curDir, text, num)

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

# initialize a new message
# PARAMS:
#     str sender   -- the email of the sender
#     str receiver -- the email of the recipient
# return MIMEMultipart, the base of the message to be sent
def newMessage(sender, receiver, num):
     # create the proper email
     msg = MIMEMultipart("")
     msg['Subject'] = 'Lab %s grade' % num
     msg['From'] = sender
     msg['To'] = receiver
     msg.preamble = 'Grading email \n'
     return msg

# append the canned text to the email object
# PARAMS:
#   MIMEMultipart msg -- the message
#   str curDir        -- the path to the code directory
#   str text          -- the path to the canned text
#   str num           -- the lab number
def addCannedText(msg, curDir, text, num):
     bodyfp = open("%s/%s" % (curDir,text), 'rb')
     text = getBody(bodyfp, num)
     msg.attach(text)
     bodyfp.close()

# Read the canned text from the file
# PARAMS
#    BufferedReader fd -- input stream for canned response
#    str num           -- the lab number
def getBody(fd, num):
     print(type(fd))
     nextline = fd.readline()
     content = "";
     while nextline:
          content = "%s%s" % (content, nextline.decode('utf-8'))
          nextline = fd.readline()
     # content = content.replace('XX', num)
     text = MIMEText(content, _subtype='plain')
     return text

# Add the .txt rubric attachment to an input message
# PARAMS:
#     str gradePath     -- the path to the grade file
#     str gradeFile     -- the name of the grade file
#     MIMEMultipart msg -- the message to send
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


# Add the .pdf commented code attachment to an input message
# PARAMS:
#     str codePath      -- the path to the code file
#     str codeFile      -- the name of the code file
#     MIMEMultipart msg -- the message to send
def attachCommentedCode(codePath, codeFile, msg):
     # try to guess the mimetype of the file
     ctype_c, encoding_c = mimetypes.guess_type(codePath)
     if ctype_c is None or encoding_c is not None:
          # No guess could be made, or the file is encoded (compressed), so
          # force pdf
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

# Parse the csv file into a key-value store.  The file name is the key,
# and the email is the value.
# PARAMS:
#     str num    -- the lab number
#     str curDir -- the grade file directory
# return Dictionary -- keys are filenames and values are emails
def parseCSV(num, curDir):
     csvfd = open(curDir + '/Emails.csv', 'r')
     dictionary = {}
     with csvfd as emails:
          reader = csv.DictReader(emails)
          for row in reader:
               dictionary[nameToFile(row['Name'], num)] = row['Email']
     return dictionary

# Convert a person's name into the appropriate Grade_* file name
# PARAMS:
#      str name -- the student's name
#      str num  -- the lab number
# return str -- the file name corresponding to the Name and lab number
def nameToFile(name, num):
     first = getFirst(name)
     last = getLast(name)
     return "Grade_Lab%s%s%s" % (num,last,first)

# Gets the first name of the full-name input
#      str name -- a full name
# return str -- the first name
def getFirst(name):
     brkIdx = name.index(" ")
     return name[:brkIdx]

# Gets the last name of the full-name input
#      str name -- a full name
# return str -- the last name
def getLast(name):
     brkIdx = -name[::-1].index(" ")
     return name[brkIdx:]

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
