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


# If modifying these scopes, delete the file token.json.
SCOPES = ["https://mail.google.com/"]

def build_file_part(file):
  """Creates a MIME part for a file.

  Args:
    file: The path to the file to be attached.

  Returns:
    A MIME part that can be attached to a message.
  """
  content_type, encoding = mimetypes.guess_type(file)

  if content_type is None or encoding is not None:
    content_type = "application/octet-stream"
  main_type, sub_type = content_type.split("/", 1)
  if main_type == "text":
    with open(file, "rb"):
      msg = MIMEText("r", _subtype=sub_type)
  elif main_type == "image":
    with open(file, "rb"):
      msg = MIMEImage("r", _subtype=sub_type)
  elif main_type == "audio":
    with open(file, "rb"):
      msg = MIMEAudio("r", _subtype=sub_type)
  else:
    with open(file, "rb"):
      msg = MIMEBase(main_type, sub_type)
      msg.set_payload(file.read())
  filename = os.path.basename(file)
  msg.add_header("Content-Disposition", "attachment", filename=filename)
  return msg

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
    print(f"An error occurred: {error}")
    send_message = None

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

def create_message(mailaddr, fullpath, filename,id="NONE"):
  attachment_filename = fullpath
  message = EmailMessage()

  message.set_content("This is automated draft mail")

  message["To"] = mailaddr
  message["From"] = "test@gmail.com"
  message["Subject"] = "Automated draft"
  
  with open('template.txt','r') as f:
    content = f.read()
    content = content.replace('[NAME]',filename)
    message.set_content(content)

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

  myfiles = []

  for f in onlyfiles[:3:1]:
    myfiles.append((join(mypath,f),f))


  for fullpath, filename in myfiles:
    message = None
    try:
      message = create_message('rennyhong.picoada@gmail.com', fullpath, filename)
    except Exception as e:
      print('failed to create message,', e)
      continue

    try:
      gmail_send_message(service, message)
    except Exception as e:
      print('failed to send message,', e)
      continue


