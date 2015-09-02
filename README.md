# grademail-PO

## Setup
Download the file `grademail.py`.  Then used the provided `canned.txt` or create your own (you may call it whatever you like). Optionally add `grademail.py` to your PATH to facilitate command line usage. Befure use, ensure that you have a directory with the students' commented code (.pdf), and the filled-in rubric (.txt), lets call this `$GRD_DIR`.  The files should have the format `Grade_LabXXLastnameFirstname.*`. If they do not, the script will not work properly.

You will also need a `.csv` file in this directory, called `Emails.csv`.  The file at `$GRD_DIR/Emails.csv` will have two columns, one specified `Name` and the other `Email`.  Each row will have a student's name (first and last only) and their email.

## Usage
In order to run the Grademail script, simply enter this into the command line (if you have `grademail.py` in your `PATH`)

    grademail.py -e [email] -t [text] -l [labnumber] -d $GRD_DIR
    
Where `email` is the email from which you wish to send your emails (should be the TA grading email address), `canned.txt` is the file from which you wish to pull your canned response, and `labnumber` is the number of the lab. The `-d` flag specifies the directory in which the script will run.  If you omit this flag, it will run in the current directory.

If you do not add `grademail.py` to your path, you will need to include the path to the script when you call it, i.e. 

    ~/Downloads/grademail.py -e [email] -t canned.txt -l [labnumber] -d $GRD_DR

It will then prompt you for your email password. Once you input it, your emails will be sent.

For more information about usage, run `grademail.py --help`.

## Development

In order to develop Grademail, you'll want to run a local debugging server, before attempting to send out files. Before you `./grademail.py` run a local debugging SMTP server using the command

    python -m smtpd -n -c DebuggingServer localhost:1025

Then, run the file with the `--DEBUG` flag.

## Known Issues
+ Grademail is only compatible with gmail.
+ To run `grademail.py`, you must enable less secure apps on your gmail account, or enable 2-factor authentication and provide Grademail with a specifically-generated app password.
+ Gmail's spam filters catch the emails sent by this script.  Outlook's filters do not, students are encouraged to provide their 5C address.
