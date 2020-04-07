import base64
from googleapiclient import discovery
from httplib2 import Http
from oauth2client import file, client, tools
from email.mime.text import MIMEText

# User manipulated variables
# In Google Sheet URL: .../d/SHEET_FILE_ID/edit#gid=0
SHEETS_FILE_ID = 'INSERT SHEET ID HERE'

# Your Email Address 
SENDER = 'INSERT EMAIL ADDRESS'

# Subject of Email
SUBJECT = 'INSERT SUBJECT EMAIL'


# General API constants
CLIENT_ID_FILE = 'credentials.json'
TOKEN_STORE_FILE = 'token.json'
SCOPES = (  # iterable or space-delimited string
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/gmail.send'
)

HTTP = get_http_client()
GMAIL = discovery.build('gmail', 'v1', http=HTTP)
SHEETS = discovery.build('sheets', 'v4', http=HTTP)



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
  message = MIMEText(message_text)
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject
  return {'raw': base64.urlsafe_b64encode(message.as_string())}

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
    return message
  except(errors.HttpError, error):
    print('An error occurred: %s' % error)




def main():
  with open('template.txt', 'r') as file:
      template = file.read()  

  data = get_sheets_data(SHEETS)
  for person in data:
      message_body = template
      for item in person:
          search_str = '{{' + item + '}}'
          message_body = message_body.replace(search_str, str(person[item]))
      send_message(GMAIL, 'me', create_message(SENDER, person['Email'], SUBJECT, message_body))
    
if __name__ == '__main__':
    main()
