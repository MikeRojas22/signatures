from string import Template
import time

import pytest
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import *
from google.auth.exceptions import *
import easygui
from string import Template
import pandas as pd


def get_api_credentials(key_path):
    API_scopes = ['https://www.googleapis.com/auth/gmail.settings.basic',
                  'https://www.googleapis.com/auth/gmail.settings.sharing']
    credentials = service_account.Credentials.from_service_account_file(key_path, scopes=API_scopes)
    return credentials


def update_sig(full_name, job_title, telephone, username, sig_template, credentials, live=False):
    sig = sig_template.substitute(full_name=full_name, job_title=job_title,
                                  telephone=telephone)
    print("Username: %s Fullname: %s Job title: %s" % (username, full_name, job_title))
    print("Username:")
    if live:
        credentials_delegated = credentials.with_subject(username)
        gmail_service = build("gmail", "v1", credentials=credentials_delegated)
        addresses = gmail_service.users().settings().sendAs().list(userId='me',
                                                                   fields='sendAs(isPrimary,sendAsEmail)').execute().get(
            'sendAs')
        # this way of getting the primary address is copy & paste from google example

        address = None
        for address in addresses:
            if address.get('isPrimary'):
                break
        if address:
            rsp = gmail_service.users().settings().sendAs().patch(userId='me',
                                                                  sendAsEmail=address['sendAsEmail'],
                                                                  body={'signature': sig}).execute()
            print(f"Signature changed for: {username}")
        else:
            print(f"Could not find primary address for: {username}")


def main():
    excelfile = easygui.fileopenbox(msg="Please open the .xlsx file with signature",
                                    title="Choose signatures .xlsx file")
    if not excelfile:
        print("No signature .xlsx file selected, so stopping")
        return
    user_data = pd.ExcelFile(excelfile)
    # df = user_data.parse("testsheet")
    df = user_data.parse("livedata")
    key_path = easygui.fileopenbox(msg="Please open the confidential Google secret .json file",
                                   title="Choose Google json secrets")
    credentials = get_api_credentials(key_path=key_path)
    if not credentials:
        print("No credential file selected, so stopping")
        return

    try:
        # build a signature
        sig_template = Template("""
            <html>
            <head>
            <meta charset="UTF-8">
            <title>Email Signature</title>
            </head>

            <body style="margin:0px; padding:0px;">
            ...

            </body>
            </html>
            """)
    except (FileNotFoundError, IOError):
        print("Could not open the template file")
        raise
    for r in df.values:
        print("Hello")
        username = r[0]
        first_name = "%s" % r[1]
        if first_name == "nan":
            first_name = ''
        second_name = "%s" % r[2]
        if second_name == "nan":
            second_name = ''
        full_name = "%s %s" % (first_name, second_name)
        job_title = "%s" % r[3]
        if job_title == 'nan':
            job_title = ''
        telephone = "%s" % r[4]
        if telephone == 'nan':
            telephone = "1300 863 824"
        retry_count = 0


        while retry_count < 3:
            try:
                update_sig(full_name=full_name, job_title=job_title, username=username, telephone=telephone,
                           sig_template=sig_template, credentials=credentials, live=True)
                break
            except (RefreshError, TransportError) as e:
                retry_count += 1
                print(f"Error encountered for: {username}, retrying (attempt {retry_count}). Error was: {e}")
                time.sleep(2)
                continue
            except Exception as e:
                raise
        else:
            print(f"Failed to update {username}")


if __name__ == '__main__':
    main()


@pytest.fixture
def get_test_api_credentials():
    return get_api_credentials(key_path='signatures-344518-de248afd6fb1.json')


def test_fetch_user_info(credentials):
    credentials_delegated = credentials.with_subject("admin-470@signatures-344518.iam.gserviceaccount.com")
    gmail_service = build("gmail", "v1", credentials=credentials_delegated)
    addresses = gmail_service.users().settings().sendAs().list(userId='me').execute()
    assert gmail_service


def test_fetch_another_user(credentials):
    credentials_delegated = credentials.with_subject("admin-470@signatures-344518.iam.gserviceaccount.com")
    gmail_service = build("gmail", "v1", credentials=credentials_delegated)
    addresses = gmail_service.users().settings().sendAs().list(userId='me').execute()
    assert gmail_service