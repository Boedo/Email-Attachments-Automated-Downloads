import boto3
import imaplib
import email
import os, sys
from os import listdir
from os.path import isfile, join
import datetime

svdir = 'C:\\Users\\nicol\\Desktop\\Scripts\\S3_Control\\Files'
date = (datetime.date.today() - datetime.timedelta(1)).strftime("%d-%b-%Y")

##########################################
# Connecting to the server, to the account
##########################################
mail = imaplib.IMAP4_SSL('outlook.office365.com',993)
mail.login('Bob@gmail.com','password')
mail.select("Inbox")

##########################################
# Email to check for attachment
##########################################
liste = [
'(TEXT "Avocet Delivery Report" SENTSINCE {date})',
'(TEXT "Avocet Conversion Report" SENTSINCE {date})',
'(TEXT "Facebook Delivery" SENTSINCE {date})',
'(TEXT "Daily Search Delivery Report" SENTSINCE {date})',
'(TEXT "Daily Search Conversion Report" SENTSINCE {date})',
'(BODY "Bing Daily Delivery Data"  SENTSINCE {date})',
'(BODY "Bing Daily Conversion Report"  SENTSINCE {date})']


##########################################
# Emails to check for attachment
##########################################
i = 0
for query in range(len(liste)):

    print query

    result, data = mail.uid('search', None, liste[i].format(date=date))
    print result, data
    target_email = data[0].split()
    print target_email

    for emailid in target_email:
        result, data = mail.uid('fetch', emailid, '(RFC822)') # fetch the email body (RFC822) for the given ID
        print result, data
        raw_email = data[0][1]
        email_msg = email.message_from_string(raw_email)
        print email_msg

        if email_msg.get_content_maintype() != 'multipart':
            continue

        for part in email_msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()
            print filename
            sv_path = os.path.join(svdir, filename)

            if not os.path.isfile(sv_path):
                fp = open(sv_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()

            i += 1

