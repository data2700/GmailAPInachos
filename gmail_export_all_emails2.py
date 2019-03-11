'''
GMAIL report maker using Python and Gmail API

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

'''
# -*- coding: utf-8 -*-
from apiclient import discovery
from apiclient import errors
from httplib2 import Http
from oauth2client import file, client, tools
import base64
from bs4 import BeautifulSoup
import dateutil.parser as parser
import csv
from time import strftime, gmtime
import sys
from datetime import datetime
import collections
from collections import Counter
from prettytable import PrettyTable

def ReadEmailDetails(service,user_id,msg_id):


  temp_dict = { }

  try:
      message = service.users().messages().get(userId=user_id,id=msg_id).execute() # fetch the message using API
      payld = message['payload'] # get payload of the message
      headr = payld['headers'] # get header of the payload

      #print('[message: %s],' % len(message))
      sys.stdout.flush()

      for one in headr: # getting the Subject
          if one['name'] == 'Subject':
              msg_subject = one['value']
              temp_dict['Subject'] = msg_subject
          #else:
              #pass


      for two in headr: # getting the date
          if two['name'] == 'Date':
              msg_date = two['value']
              #date_parse = (parser.parse(msg_date))
              #msg_date = (date_parse.datetime())
              if not "," in msg_date:
                 msg_date = msg_date.rsplit(":",1)[0].rstrip()

                 print('Date string fix 1. rare')
              if "," in msg_date:
                 msg_date = msg_date.rsplit(":",1)[0].rstrip()
                 msg_date = msg_date.split(",",1)[1].strip()

                 print('Date string fix 2 common ')

              msg_date = datetime.strptime(msg_date,'%d %b %Y %H:%M')
              temp_dict['DateTime'] = msg_date
          else:
              pass
                
      for three in headr: # getting the Sender
          if three['name'] == 'From': #Use either 'From' or 'To' depending on which messages you want to count
              msg_from = three['value']
              #clean up Gmail's changing address fields to get domain
              #msg_from = msg_from.split(">", 1)[0].rstrip()
              msg_from = msg_from.split("@")[1].lstrip()
              msg_from = msg_from.split("<")[0].rstrip()
              msg_from = msg_from.split(">")[0].rstrip()
              if msg_from.endswith('"'):msg_from = msg_from[:-1]

              #msg_from = msg_from.split(" ", 1)[0].rstrip()
              print('Message Domain:    %s,' % (msg_from))


              temp_dict['Sender'] = msg_from

          else:
              pass

      # Fetching message body
      # email_parts = payld['parts'] # fetching the message parts
      # part_one  = email_parts[0] # fetching first element of the part
      # part_body = part_one['body'] # fetching body of the message
      # part_data = part_body['data'] # fetching data from the body
      # clean_one = part_data.replace("-","+") # decoding from Base64 to UTF-8
      # clean_one = clean_one.replace("_","/") # decoding from Base64 to UTF-8
      # clean_two = base64.b64decode (bytes(clean_one, 'UTF-8')) # decoding from Base64 to UTF-8
      # soup = BeautifulSoup(clean_two , "lxml" )
      # message_body = soup.body()
      # message_body is a readible form of message body
      # depending on the end user's requirements, it can be further cleaned
      # using regex, beautiful soup, or any other method
      # temp_dict['Message_body'] = message_body
      

  except Exception as e:
      print(e)
      temp_dict = None
      pass

  finally:
      return temp_dict


def ListMessagesWithLabels(service, user_id, q='',label_ids=[]):
  """List all Messages of the user's mailbox with label_ids applied.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    label_ids: Only return Messages with these labelIds applied.

  Returns:
    List of Messages that have all required Labels applied. Note that the
    returned list contains Message IDs, you must use get with the
    appropriate id to get the details of a Message.
  """
  try:
    response = service.users().messages().list(userId=user_id,
                                               q='after:2018/09/01 before:2018/09/02',
                                               labelIds=label_ids,
                                               maxResults=500).execute()


    messages = []
    if 'messages' in response:
      messages.extend(response['messages'])

    while 'nextPageToken' in response:
      page_token = response['nextPageToken']

      response = service.users().messages().list(userId=user_id,
                                                 q='after:2018/09/01 before:2018/09/02',
                                                 labelIds=label_ids,
                                                 pageToken=page_token,
                                                 maxResults=500).execute()

      messages.extend(response['messages'])

      print('... total %d emails on next page [page token: %s], %d listed so far' % (len(response['messages']), page_token, len(messages)))
      sys.stdout.flush()

    return messages

  except errors.HttpError as error:
    print('An error occurred: %s' % error)


if __name__ == "__main__":
  print('\n... start')

  # Creating a storage.JSON file with authentication details
  SCOPES = 'https://www.googleapis.com/auth/gmail.readonly' # we are using readonly, as we will be marking the messages Read
  store = file.Storage('storage.json')
  creds = store.get()

  if not creds or creds.invalid:
      flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
      creds = tools.run_flow(flow, store)

  GMAIL = discovery.build('gmail', 'v1', http=creds.authorize(Http()))

  user_id =  'me'
  label_id_one = 'INBOX'
  label_id_two = 'UNREAD'

  print('\n... list all emails')

  # email_list = ListMessagesWithLabels(GMAIL, user_id, [label_id_one,label_id_two])  # to read unread emails from inbox
  email_list = ListMessagesWithLabels(GMAIL, user_id, [])

  final_list = [ ]

  print('\n... fetching all emails data, this will take some time')
  sys.stdout.flush()


  #exporting the values as .csv
  rows = 0
  file = 'emails_%s.csv' % (strftime("%Y_%m_%d_%H%M%S", gmtime()))
  #file = 'emails_%s.csv' % (strftime("%Y_%m_%d_%H%M"))

"""
  with open(file, 'w', encoding='utf-8', newline = '') as csvfile:
      #fieldnames = ['Subject','DateTime','Message_body']
      fieldnames = ['Subject','DateTime','Sender']
      writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter = ',')
      writer.writeheader()

      for email in email_list:
        msg_id = email['id'] # get id of individual message

        email_dict = ReadEmailDetails(GMAIL,user_id,msg_id)

        if email_dict is not None:
          writer.writerow(email_dict)
          rows += 1

        if rows > 0 and (rows%50) == 0:
          print('... total %d read so far' % (rows))
          sys.stdout.flush()

  print('... emails exported into %s' % (file))
  print("\n... total %d message retrived" % (rows))
  sys.stdout.flush()"""


  #exporting the values as .csv
with open(file, 'w', newline = '', encoding='utf-8-sig') as csvfile: 
      fieldnames = ['Subject','DateTime','Sender']
      writer = csv.DictWriter(csvfile, extrasaction='ignore', fieldnames=fieldnames, delimiter = ',')
      writer.writeheader()
      print ("Length of Email List: %d" % len (email_list))
      #counter=collections.Counter(email_list)
      #print (email_list)
      a = 0
      for email in email_list:
         msg_id = email['id'] # get id of individual message
         email_dict = ReadEmailDetails(GMAIL,user_id,msg_id)
         #print ("Length : %d" % len (email_dict))
         #counter=collections.Counter(email_list)
         # print (email_dict)
         if email_dict is not None:
          writer.writerow(email_dict)
          rows += 1
          a += 1
          print ("Message counter: %d " % (a) + "\n")

         if rows > 0 and (rows%50) == 0:
          print('... total %d read so far' % (rows))
sys.stdout.flush()


print('... emails exported into %s' % (file))
print("\n... total %d message retrieved" % (rows) + "\n")
sys.stdout.flush()

print('... Frequency counting starts now - ugly table)' + "\n") # no library required for this table
with open(file, newline='', encoding='utf-8-sig') as csvfile:
  x = PrettyTable(["domain", "Count"])            #create instance of pretty table with two table columns
  senderreader = csv.reader(csvfile, delimiter=',')
  next(senderreader) #skip header
  Sender = [row[2] for row in senderreader]
  for (k,v) in Counter(Sender).items():
      print ("%s appears %d times" % (k, v))
#------loop below is to create the pretty table output -- need to import pretty table for this one to work

  for (k,v) in Counter(Sender).items():
      x.align = "l"                               #align columns for pretty table to the left
      x.add_row([k,v])                            #populate pretty table with the domain and sender
  print("\n" + ' - pretty table of previous data-' + "\n")
  print(x)                                          #print pretty table of domains and count of domains




'''
def CountFrequency(email_dict): 
      
    # Creating an empty dictionary  
    freq = {} 
    for items in email_dict: 
        freq[items] = email_dict.count(items) 
      
    for key, value in [list(freq.items())]: 
        print ("% d : % d"%(key, value)) 
#in_hist = [list(in_degrees.values()).count(x) for x in in_values]

# Driver function 
if __name__ == "__main__":  
        CountFrequency(email_dict) 
'''
print('... all Done!')
