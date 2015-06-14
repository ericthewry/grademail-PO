# grademail-PO

## Usage

In order to run the GradEmail Script, simply enter this into the command line

    ./grademail.py -e [email] -t [text] -l [labnumber]
    
Where `email` is the email from which you wish to send your emails (currently only works for gmail), `text` is the file from which you wish to pull your canned response, and `labnumber` is the number of the lab.

It will then prompt you for your email password.

## Development

In order to develop GradEmail, you'll want to run a local debugging server, before attempting to send out files. Before you `./grademail.py` run a local debugging SMTP server using the command

    python -m smtpd -n -c DebuggingServer localhost:1025

Then, run the file with the -DEBUG flag
