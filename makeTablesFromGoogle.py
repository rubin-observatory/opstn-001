#!/usr/bin/env python
# TO access gogole you have to do some setup https://developers.google.com/sheets/api/quickstart/python
# rougly you muct create a client secret for OAUTH using this wizard https://console.developers.google.com/start/api?id=sheets.googleapis.com
# Accept the blurb and go to the APIs
# click CANCLE on the next screen to ge tto the "create credentials"
# hit create credentials choose OATH client
# Configure a product - just put in a name like "LSST DOCS"
# Creat web application id
# Give it a name hit ok on the next screen
# now you can download the client id - call it client_secret.json as expected below.
# You have to do this once to allow the script to access yuour google stuff from this machine


from __future__ import print_function
import httplib2
import os
import sys
import pickle

from oauth2client import client
from oauth2client import tools
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Python LSST'


def getCredentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    credentials = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            credentials = pickle.load(token)

    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES )
            flow.user_agent = APPLICATION_NAME
            credentials = flow.run_local_server()
        with open('token.pickle', 'wb') as token:
            pickle.dump(credentials, token)

    return credentials


def outtail( tout):
    print (r'\end{longtable}', file=tout)
    tout.close()
    return

def outhead(ncols, tout,name, cap, width):

    print (" \\begin{longtable} {", file=tout, end='')
    c =1
    print (" |p{%f\\textwidth} "%(width), file=tout, end='')
    for c in range(1,ncols,+1):
       print (" |r ", file=tout, end='')
    print ("|} ", file=tout, )
    print ('\\caption{%s \\label{tab:%s}}\\\\ ' % (cap,name), file=tout, )
    print ("\hline ", file=tout)
    return

def outputrow(tout, pre, row, cols):
    for i in range(cols):
        #print(i)
        try:
            print("%s{%s}" % (pre, fixTex(row[i])), end='', file=tout)
        except IndexError:
            pass
        if i < cols-1:
            print("&", end='', file=tout)
    print (' \\\\ \hline', file=tout)


def fixTex(text):
    ret = text.replace("_", "\\_")
    ret = ret.replace("/", "/ ")
    ret = ret.replace("$", "\\$")
    return ret



def main():
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    creds = getCredentials()

    service = build('sheets', 'v4', credentials=creds)


    spreadsheetId = '1s-JJW2v-1AIxno0yL5pJzORxWSyxLLSXyHIxgHvQJE4'
    rangeName = 'DP!A1:H'

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    name=''
    tout = ''

    rowc=0
    if not values:
        print('No data found.')
    else:
        print(values)
        cols = 0
        for row in values:
            if (row and 'Table' in row[0] ):  # got a new table
                if  name :
                   outtail(tout)
                vals=row[0].split(' ')
                name= vals[1]
                print ("Create new table %s" % name)
                tout = open(name+'.tex','w')

                cap=row[1]
                cols=int(row[2])
                width=float(row[3])
                outhead(cols, tout, name, cap, width)
                rowc=0
            else:
                if name and row:
                    if row[0].startswith('Year') or row[0].startswith('Total') or rowc == 0:
                        # print headeri/total in bold
                        outputrow(tout, "\\textbf", row, cols)
                    else:
                        outputrow(tout, "", row, cols)
                rowc = rowc + 1
    outtail(tout)


if __name__ == '__main__':
    main()

