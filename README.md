# Python - Outlook Reader and Sender

Send Emails and Download Attachments from Outlook using Python

## How To Use
Email is a full program with UI and customizable settings.  
PCC_Reader is a batch programs, it requires to be opened with CMD. 

The diferences are:
PCC_Reader.exe downloads the message as a .msg file;
Email.exe downloads all unread emails attachments.

#### Run PCC_Reader.exe: (CMD COMMAND)
".\file_path\PCC_Reader.exe" "email_to_check_from@domain.com" ".\file_path\\<email-subject\>.msg"

#### Run PCC_Reader.py: (CMD COMMAND)
python -m ".\file_path\PCC_Reader.exe" "email_to_check_from@domain.com" ".\file_path\\<email-subject\>.msg"

#### Run Email.exe:
Just run the program... It will download all attachments from newer emails and mark as read.
