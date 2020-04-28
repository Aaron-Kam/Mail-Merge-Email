import base64
from googleapiclient import discovery
from httplib2 import Http
from oauth2client import file, client, tools
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import time

# User manipulated variables
# In Google Sheet URL: .../d/SHEET_FILE_ID/edit#gid=0
SHEETS_FILE_ID = '1TkEP8fPBPLuYbkOBJIBsPRx7HpdJiMQXlFpugtsOvSk'

# Your Email Address 
SENDER = 'ark476@gmail.com'

# Subject of Email
SUBJECT = 'Hi there, this is a test'

# Email Purpose and date (e.g. "Mandatory 4/11/20")
PURPOSE = 'Mandatory adv 4/20/20'

# Number of emails per batch
BATCH_AMT = 5

# Time between email batches in seconds
SLEEP_TIME = 5

# General API constants
CLIENT_ID_FILE = 'credentials.json'
TOKEN_STORE_FILE = 'token.json'
SCOPES = (  # iterable or space-delimited string
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/gmail.send'
)




def get_sheets_data(service):
    """(private) Returns data from Google Sheets source. It gets all rows of
        'Sheet1' (the default Sheet in a new spreadsheet), but drops the first
        (header) row. Use any desired data range (in standard A1 notation).
    """
    headers = service.spreadsheets().values().get(spreadsheetId=SHEETS_FILE_ID,
            range='Sheet1').execute().get('values')[0] # only header row 
    recipient_info = service.spreadsheets().values().get(spreadsheetId=SHEETS_FILE_ID,
            range='Sheet1').execute().get('values')[1:] # skip header row
    
    final_data = []
    for person in recipient_info:
        final_data.append(dict(zip(headers, person)))
    return final_data

def get_http_client():
    """Uses project credentials in CLIENT_ID_FILE along with requested OAuth2
        scopes for authorization, and caches API tokens in TOKEN_STORE_FILE.
    """
    store = file.Storage(TOKEN_STORE_FILE)
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_ID_FILE, SCOPES)
        creds = tools.run_flow(flow, store)
    return creds.authorize(Http())

HTTP = get_http_client()
GMAIL = discovery.build('gmail', 'v1', http=HTTP)
SHEETS = discovery.build('sheets', 'v4', http=HTTP)


def create_message(sender, to, subject, message_text):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.

  Returns:
    An object containing a base64url encoded email object.
  """
  
  message = MIMEMultipart('related')
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject
  message.preamble = 'This is a multi-part message in MIME format.'  
  
  messageAlternative = MIMEMultipart('alternative')
  message.attach(messageAlternative)  
  
  text="Hi!\nHow are you?\nBye"
  part1=MIMEText(text, 'plain')
  messageAlternative.attach(part1)
  
  partHtml=MIMEText(message_text, 'html')
  messageAlternative.attach(partHtml)   
  
  fp = open('logo.png','rb')
  messageImage = MIMEImage(fp.read())
  fp.close
  
  messageImage.add_header('Content-ID', '<image1>')
  message.attach(messageImage)  
  
  return {'raw': base64.urlsafe_b64encode(message.as_string())}

def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def send_message(service, user_id, message):

  """Send an email message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

  Returns:
    Sent Message.
  """
  try:
    message = (service.users().messages().send(userId=user_id, body=message)
               .execute())
    print('Message Id: %s' % message['id'])
    return time.asctime()
  except(errors.HttpError, error):
    print('An error occurred: %s' % error)
    return False



def main():
  with open('message.html', 'r') as file:
      template = file.read()
  data = get_sheets_data(SHEETS)
  email_sent = [[PURPOSE]]

  num_sent = 0
  for person in data:
      message_body = template
      for item in person:
          search_str = '{{' + item + '}}'
          message_body = message_body.replace(search_str, str(person[item]))

      work = send_message(GMAIL, 'me', create_message(SENDER, person['Email'], SUBJECT, message_body))
      email_sent.append([str(work)])
      num_sent =+ 1

      if num_sent % SLEEP_TIME == 0: time.sleep(BATCH_AMT)
  
  last_column = colnum_string(len(data[0]) + 1)
  range_ =  last_column + ':' + last_column
  body={
    'range':range_,
    'majorDimension': 'ROWS',
    'values': email_sent}
  #SHEETS.spreadsheets().values().get(spreadsheetId=SHEETS_FILE_ID)
  SHEETS.spreadsheets().values().update(
    spreadsheetId=SHEETS_FILE_ID,
    range=range_,
    valueInputOption='RAW',
    body=body).execute()
    
    
if __name__ == '__main__':
    main()
