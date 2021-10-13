# GmailAPInachos
Gmail inbox management.  This will help you to inventory your gmail inbox by sender.  See how many messages you are getting from each email domain.

Gmail Api through python with domain counts of inbox data, csv export and count of domains sender
MAIL report maker using Python and Gmail API.  This script can handle very large gmail inboxes.

    - Richard 

Based on Original GMAIL reader by 
    - Imran Momin and other similar python scripts
'''

'''
This script does the following:
- Go to Gmail inbox
- Find a Date range of E-mails
- Extract details (Date, Subject, Send) and export them to a .csv file
--Read details from csv file and Provide a 'Pretty Table' report of the sender or receiver domains during the given period
--Converts all date and timestamps to a sortable format that works with Excel
--Converts subject texts to utf-8-sig to support the most languages
--provides live status of processing activities
--allows for processing over the 500 message limit of the Gmail API 

'''

'''
Before running this script, the user should get the authentication by following
the link: https://developers.google.com/gmail/api/quickstart/python
Also, client_secret.json should be saved in the same directory as this file

To run script go to python terminal and type:

python gmail_export_all_emails2.py (or whatever name you called this file)
