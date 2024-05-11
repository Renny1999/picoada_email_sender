import os.path
import base64
import mimetypes
from os.path import isfile, join
from os import listdir
from os.path import splitext


from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from email.message import EmailMessage
import base64
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText


import google.auth
import csv

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
year = 2024
month = 4

def gmail_send_message(service, message):
  """Create and insert a draft email.
   Print the returned draft's message and id.
   Returns: Draft object, including draft id and message meta data.

  Load pre-authorized user credentials from the environment.
  TODO(developer) - See https://developers.google.com/identity
  for guides on implementing OAuth2 for the application.
  """
  try:
    # pylint: disable=E1101
    send_message = (
        service.users()
        .messages()
        .send(userId="me", body=message)
        .execute()
    )

  except HttpError as error:
    send_message = None
    raise error

  return send_message

def connect():
  """Shows basic usage of the Gmail API.
  Lists the user's Gmail labels.
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    # Call the Gmail API
    service = build("gmail", "v1", credentials=creds)
    return service

  except HttpError as error:
    # TODO(developer) - Handle errors from gmail API.
    print(f"An error occurred: {error}")
    return None


def create_message(subject:str, \
                   sender_mailaddr, \
                   mailaddr:list[str], \
                   fullpath:str, \
                   filename:str, \
                   facility='', \
                   name='', \
                   cc:list[str] = [], \
                   id="NONE"):
  """
  IN:
    sender_mailaddr: sender email address
    mailaddr: receiver fmail address
    fullpath: full path of the attachment
    filename: name of the attachment
  """
  attachment_filename = fullpath
  message = EmailMessage()

  message["To"] = mailaddr
  message["From"] = "株式会社ピコ・エイダ <"+sender_mailaddr+">"
  message["CC"] = cc
  message["Subject"] = subject
  
  with open('template.txt','r') as f:
    content = f.read()
    content = content.replace('[FACILITY]',facility)
    content = content.replace('[NAME]',name)
    # message.add_alternative(content,subtype='html')
    # message.add_alternative(content,subtype='raw')
    message.add_alternative(content)

  # attachment
  # guessing the MIME type
  type_subtype, _ = mimetypes.guess_type(attachment_filename)
  maintype, subtype = type_subtype.split("/")

  attachment_data = None
  with open(attachment_filename, "rb") as fp:
    attachment_data = fp.read()
    message.add_attachment(attachment_data, maintype, subtype, filename=filename)
  if (attachment_data == None):
    return None

  # encoded message
  encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

  create_message = {"raw": encoded_message}

  return create_message

if __name__ == "__main__":
  service = connect()
  if (service == None):
    print("failed to connect")
    exit(1)

  cwd = os.getcwd()
  mypath = str(cwd) + '/pdf'
  onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]

  # get the customer info 
  # contains CustomerID, CustomerName, CustomerEmail
  customer_info = open('customers.csv', 'r')

  # contains CustomerID, FileName
  file_maps = open('log.csv', 'r')
    
  # create a mapping between ID and filename
  file_maps_dict = {}
  for l in file_maps.readlines():
    if len(l) <= 1:
      break
    l = l.strip().split(',')
    file_maps_dict[l[0]] = l[2].replace(".xlsx", "")
  file_maps.close()

  myfiles = {}
  # myfiles maps file name to the file full path
  for f in onlyfiles:
    myfiles[f.replace(".pdf", "")] = join(mypath,f)

  outputlog = open('output.log', 'w')

  csvfile = open('emails.csv','r')
  reader = csv.reader(csvfile, delimiter=',', quotechar='\"')
  email_data = {}
  for r in reader:
    customer_id_email = r[0]
    customer_facility1_email = r[1]
    customer_facility2_email = r[2]
    customer_name_email = r[3]
    customer_address_email = r[4].split('\n')
    customer_cc_email = r[5].split('\n')

    # if (customer_facility2_email == ''):
    #   continue
    #
    # if (customer_name_email == ''):
    #   continue

    email_data[customer_id_email] = {'facility':customer_facility2_email,
                                     'name': customer_name_email,
                                     'address': customer_address_email,
                                     'cc': customer_cc_email}

  # read the lines from the cache
  for cid in email_data.keys():
    cus = email_data[cid] # customer info
    print(cus)
    logmsg = "{},{},".format(cid,cus["facility"])
    outputlog.write(logmsg)
    print(len(cus["address"]))
    if (len(cus["address"]) == 0):
      print("NO ADDRESS")

    # get the customer mapping (id, name, email)
    if (cid not in file_maps_dict):
      errormsg = 'did not find {} for ID->filename mapping'.format(cid)
      print(errormsg)
      outputlog.write(',0,'+errormsg+'\n')
      continue
    filename: str = file_maps_dict[cid]

    # get full file path using the filename
    if (filename not in myfiles):
      errormsg = 'could not find pdf for id={}, filename={}'.format(cid, filename)
      print(errormsg)
      outputlog.write(',0,'+errormsg+'\n')
      continue

    fullpath = myfiles[filename]
    print("file full path: ", fullpath)
    
    message = None
    try:
      subject = "{}様{}年{}月月報のご送付".format(cus["facility"],year,month)
      message = create_message(subject,
                              'rennyhong1103@gmail.com',  # sender address
                               cus['address'],  # receiver address
                               fullpath, 
                               facility=cus['facility'],
                               name=cus['name'],
                               cc=cus['cc'],
                               filename=filename+".pdf")

    except Exception as e:
      print('failed to create message,', e)
      outputlog.write(',0,failed to create message\n')
      continue

    try:
      gmail_send_message(service, message)
      outputlog.write('1,message successfully sent\n')
    except Exception as e:
      print('failed to send message,', e)
      errormsg = '0,failed to send message, reason: \"{}\"\n'.format(e)
      outputlog.write(errormsg)
      continue
  outputlog.close()


