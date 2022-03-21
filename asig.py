from __future__ import print_function
import pandas as pd
import httplib2
from google.oauth2 import service_account
import gdata.apps.emailsettings.client
import gdata.gauth
import ssl

from string import Template

excelfile = "email_list.xlsx"

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


def update_sig(full_name, job_title, telephone, username, sig_template, gmail_client=None):
    sig = sig_template.substitute(full_name=full_name, job_title=job_title,
                                  telephone=telephone)

    if gmail_client.UpdateSignature(username=username, signature=sig):
        return "Signature changed"
    else:
        return None


def setup_credentials(delegate_email=None):
    # delegated credentials seem to be required for email_settings api
    key_path = 'signatures-344518-de248afd6fb1.json'
    API_scopes = ['https://apps-apis.google.com/a/feeds/emailsettings/2.0/']
    credentials = service_account.Credentials.from_service_account_file(key_path, scopes=API_scopes)
    if delegate_email:
        return credentials.create_delegated(delegate_email)

    return credentials


def test_1(credentials, domain, user):
    """ a low level test of something simple. If this works, you have the authentication working
    User is the name of the user, not the full email address. e.g. tim """
    http = httplib2.Http()
    # http = credentials.authorize(http)
    http = credentials.authorize(http)  # this is necessary
    url_get_sig = 'https://apps-apis.google.com/a/feeds/emailsettings/2.0/{domain}/{user}/signature'.format(**locals())
    r = http.request(url_get_sig, "GET")
    return r


if __name__ == "__main__":
    delegated_credentials = setup_credentials('admin-470@signatures-344518.iam.gserviceaccount.com')  # this user is a super admin.

    # test with a low level attempt
    (res1, res2) = test_1(delegated_credentials)
    assert (res1['status'] == '200')

    # try with the EmailSettingsClient
    client = gdata.apps.emailsettings.client.EmailSettingsClient(domain='vci.com.au')
    auth2token = gdata.gauth.OAuth2TokenFromCredentials(delegated_credentials)
    auth2token.authorize(client)
    r = client.retrieve_signature('tim')
    print(client)

    # now read a spreadsheet and bulk update
    user_data = pd.ExcelFile(excelfile)
    df = user_data.parse("Sheet1")
    for r in df.values:
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

        try:
            print(update_sig(full_name=full_name, job_title=job_title, gmail_client=client,
                             telephone=telephone, username=username, sig_template=sig_template))
        except gdata.client.RequestError:
            print("Client Error (user not found probably) for {}".format(full_name))
            continue
        except ssl.SSLEOFError:
            print("SSL handshake Error for {}".format(full_name))
            continue