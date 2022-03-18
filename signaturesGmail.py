from __future__ import print_function
from googleapiclient.discovery import build
from apiclient import errors
from httplib2 import Http
from email.mime.text import MIMEText
import base64
from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/gmail.settings.basic']

# This is your credentials file you downloaded from the Platform Console.
credsFile = 'service-key.json'

def signatureUpdater():
 #...
  emailAddresses = ["Populate", "an", "array", "with", "your", "user", "email", "addresses", "using", "Users: list", "method", "of", "Admin", "SDK", "API"]
  q = {"signature": "The Signature You want your users to have <b>can include some html</b>"}

  for x in emailAddresses:
    response = serviceAccountLogin(x).users().settings().sendAs().update(userId = x, q = q).execute()

def serviceAccountLogin(email):
  credentials = service_account.Credentials.from_service_account_file(credsFile, scopes = SCOPES)
  delegatedCreds = credentials.with_subject(email)
  service = build('gmail', 'v1', credentials = delegatedCreds)
  return service

def main():
  signatureUpdater()