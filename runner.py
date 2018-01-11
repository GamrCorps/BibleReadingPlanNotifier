import base64
import datetime
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes
import os
import time

from apiclient import errors, discovery
import argparse
import httplib2
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import schedule

SCOPES = ('https://www.googleapis.com/auth/gmail.compose',
          'https://www.googleapis.com/auth/spreadsheets.readonly')
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Bible Reading Plan Notifier'


def send_message(service, user_id, message):
    try:
        message = service.users().messages().send(userId=user_id,
                                                  body=message).execute()
        print('Message Id: %s' % message['id'])
        return message
    except errors.HttpError as error:
        print('An error occurred: %s' % error)


def create_message(sender, to, subject, message_text):
    message = MIMEText(message_text)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}


def create_message_html(sender, to, subject, html):
    message = MIMEText(html, 'html')
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}


def create_message_attachment(sender, to, subject, message_text, file_dir,
                              filename):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    msg = MIMEText(message_text)
    message.attach(msg)

    path = os.path.join(file_dir, filename)
    content_type, encoding = mimetypes.guess_type(path)

    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'text':
        fp = open(path, 'rb')
        msg = MIMEText(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'image':
        fp = open(path, 'rb')
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'audio':
        fp = open(path, 'rb')
        msg = MIMEAudio(fp.read(), _subtype=sub_type)
        fp.close()
    else:
        fp = open(path, 'rb')
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        fp.close()

    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)

    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}


def get_credentials():
    home_dir = os.path.expanduser('~/workspace')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'bible-reading-plan-notifier.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store, flags)
        print('Storing credentials to ' + credential_path)
    return credentials


def send_email(email_service, sheet_service):
    spreadsheet_id = '1hSLmVVBJBwClc2tqMX0mwvQ4b3kwrRq4QJGuh_By26A'
    verse_range = 'Verses!A:B'
    email_range = 'Users!A:A'

    today = str(datetime.date.today())

    request = sheet_service.spreadsheets().values().batchGet(
        spreadsheetId=spreadsheet_id, ranges=[verse_range, email_range])
    response = request.execute()

    verse_list = response['valueRanges'][0]['values']
    email_list = response['valueRanges'][1]['values']

    verse = None
    for l in verse_list:
        d, v = l
        if d == today:
            verse = v

    if verse:
        for e in email_list:
            msg = create_message('refinerybiblereadingplan@gmail.com',
                                 e[0], 'Verse for ' + today, verse)
            send_message(email_service, 'refinerybiblereadingplan@gmail.com',
                         msg)

    return None


if __name__ == '__main__':
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    email_service = discovery.build('gmail', 'v1', http=http)

    sheet_service = discovery.build('sheets', 'v4', credentials=credentials)

    schedule.every().day.at('7:00').do(send_email, email_service,
                                       sheet_service)

    while True:
        schedule.run_pending()
        time.sleep(300)
