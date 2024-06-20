#!/usr/bin/env python3

import dropbox
from dropbox import DropboxOAuth2FlowNoRedirect

import dotenv
from dotenv import load_dotenv

import os

load_dotenv()
APP_KEY = os.environ['APP_KEY']

'''
Populate your app key in order to run this locally
'''
auth_flow = DropboxOAuth2FlowNoRedirect(APP_KEY, use_pkce=True, token_access_type='offline')

authorize_url = auth_flow.start()
print("1. Go to: " + authorize_url)
print("2. Click \"Allow\" (you might have to log in first).")
print("3. Copy the authorization code.")
auth_code = input("Enter the authorization code here: ").strip()

try:
    oauth_result = auth_flow.finish(auth_code)
except Exception as e:
    print('Error: %s' % (e,))
    exit(1)

with dropbox.Dropbox(oauth2_refresh_token=oauth_result.refresh_token, app_key=APP_KEY) as dbx:
    dbx.users_get_current_account()
    print("Successfully set up client!\n")
    print(oauth_result)
    print("Refresh token:\n");
    print(oauth_result.refresh_token)

    dotenv_file = dotenv.find_dotenv()
    dotenv.set_key(dotenv_file, "REFRESH_TOKEN", oauth_result.refresh_token)
    print("Refresh Token has been saved to: " + dotenv_file);