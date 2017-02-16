import requests
import json
import csv
import argparse
import os
import httplib2
import mimetypes
import smtplib
import base64
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from urllib2 import HTTPError

SCOPES = 'https://www.googleapis.com/auth/gmail.send'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'CS52 grade sender'
SENDER = 'cs052grading@gmail.com'


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("subject",
                        help="subject of email to be sent to all students")
    parser.add_argument("file",
                        help="name of file to be sent to all students with ext.")
    parser.add_argument("--config",
                        help="path to config file")
    parser.add_argument("--roster",
                        help="path to csv file with roster")
    parser.add_argument("--message",
                        help="path to .txt file with message for everyone")
    parser.add_argument("--directory",
                        help="directory to look for files")

    args = parser.parse_args()

    # preliminary settings
    cur_dir = os.getcwd()
    message_text = os.path.join(cur_dir, 'working_directory', 'message.txt')
    csv_path = os.path.join(cur_dir, 'working_directory', 'students.csv')
    config_path = os.path.join(cur_dir, 'working_directory', 'config.json')
    file_directory = cur_dir
    desired_file = args.file
    subject = args.subject
    recipients = {}

        # check optional arguments
    if (args.config != None):
        config_path = args.config
    if (args.roster != None):
        csv_path = args.roster
    if (args.message != None):
        message_text = args.message
    if (args.directory != None):
        file_directory = args.directory

    # get standard message to be included in body of email
    with open (message_text, "r") as message_file:
        message_text = message_file.read()

    # get students and emails
    recipients = csv_parser(csv_path)

    #
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)

    # check to see if the files are going to be good
    recipients = check_batch(recipients, desired_file, file_directory)[0]

    # send emails if user confirms
    if (raw_input('Type confirm to send emails:').lower() == 'confirm'):
        print("confirm")
        send_all(SENDER, recipients, subject, message_text, desired_file, file_directory, service, 'me')

# generate file path for directory structure
def make_path(the_file, user, directory):
    return os.path.join(directory, user, the_file)

# get user and email addresses form CSV into a dictionary
def csv_parser(file_path):
    csv_file = open(file_path, 'r')
    dictionary = {}
    with csv_file as emails:
        reader = csv.DictReader(emails)
        for row in reader:
            dictionary[row['USER']] = row['EMAIL']
    return dictionary

def check_batch(recipients, desired_file, parent_directory):
    succeed_recipients = {}
    failed_recipients = {}

    # try opening each file that we are going to send
    for user, email in recipients.iteritems():
        try:
            with open(make_path(desired_file, user, parent_directory)) as attachment:
                succeed_recipients[user] = email
        except IOError:
            failed_recipients[user] = email

    # print results
    print("Files found for the following students:")
    for usr, eml in succeed_recipients.iteritems():
        print( '\033[92m' + usr + '<'+ eml + '>' + '\033[0m')

    print("Files NOT found for the following student:")
    for usr, eml in failed_recipients.iteritems():
        print('\033[91m' + usr + '<'+ eml + '>' + '\033[0m')

    return [succeed_recipients, failed_recipients]

# send personalized email to student with grade
def send_grade(sender, recipient, subject, message_text, attachment):
    message = MIMEMultipart()
    message['To'] = recipient
    message['From'] = sender
    message['Subject'] = subject

    msg = MIMEText(message_text)
    message.attach(msg)

    content_type, encoding = mimetypes.guess_type(attachment)

    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'text':
        fp = open(attachment, 'rb')
        msg = MIMEText(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'image':
        fp = open(attachment, 'rb')
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'audio':
        fp = open(attachment, 'rb')
        msg = MIMEAudio(fp.read(), _subtype=sub_type)
        fp.close()
    else:
        fp = open(attachment, 'rb')
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        fp.close()
    filename = os.path.basename(attachment)
    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)
    return {'raw': base64.urlsafe_b64encode(message.as_string())}

# send emails to all students
def send_all(sender_info, recipients, subject, message, desired_file, parent_directory, service, user_id):
    print ("Sending...")
    for user, email in recipients.iteritems():
            msg = send_grade(sender_info, email, subject, message, make_path(desired_file, user, parent_directory))
            try:
                result = (service.users().messages().send(userId=user_id, body=msg).execute())
                print ('\033[92m' + 'SENT  ' + user + '<' + email + '>' + '\033[0m')
            except HTTPError as error:
                print ('\033[91m'+ 'FAILED ' + user + '<' + email + '>' + '\033[0m')
            except:
                print ('\033[91m' + 'Other error ' + '\033[0m')
    print ("Complete.")

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'grader.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store, tools.argparser.parse_args(args=[]))
        print('Storing credentials to ' + credential_path)
    return credentials



if __name__ == '__main__':
    main()
