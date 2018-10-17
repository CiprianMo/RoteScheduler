from __future__ import print_function
import datetime
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

SCOPES = 'https://www.googleapis.com/auth/calendar'

def colors():

    
   
    store = file.Storage('token.json')
    cred = store.get()
    if not cred or cred.invalid:
        flow = client.flow_from_clientsecrets('credentials.json',SCOPES)
        cred = tools.run_flow(flow,store)

    service = build('calendar', 'v3', http=cred.authorize(Http()))

    colors = service.colors().get().execute()

    for id, color in colors['event'].iteritem():
        print ('colorId: %s' % id)
        print ('  Background: %s' % color['background'])
        print ('  Foreground: %s' % color['foreground'])

colors()