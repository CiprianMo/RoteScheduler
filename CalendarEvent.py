from __future__ import print_function
import datetime
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

SCOPES = 'https://www.googleapis.com/auth/calendar'

class CalendarEvent():

    def __init__(self,credentialsJson,calendarId = 'primary'):

        self.store = file.Storage('token.json')
        self.cred = self.store.get()

        if not self.cred or self.cred.invalid:
            flow = client.flow_from_clientsecrets(credentialsJson,SCOPES)
            self.cred = tools.run_flow(flow,self.store)

        self.service = build('calendar', 'v3', http=self.cred.authorize(Http()))

        self.clID = calendarId #example : '709l1hm8igeid4b0d0onpl8q5o@group.calendar.google.com'

    def InsertEvent(self, event):    

        #call the Calendar API
        now = datetime.datetime.utcnow().isoformat()+'Z' #Z indicates UTC time     

        event_result = self.service.events().insert(calendarId=self.clID,body = event).execute()
        print ('Event created: %s' % (event_result.get('htmlLink')))

    def valiDate(self,date_text):
        try :
            datetime.datetime.strptime(date_text, '%Y-%m-%d')
        except ValueError:
            raise ValueError("Incorrect data format, should be YYYY-MM-DD")

    def GetEventsInDateRange(self,dateFrom,dateTo):

        event_results = self.service.events().list(calendarId=self.clID,timeMin = dateFrom,timeMax = dateTo).execute()
        return event_results

    def DeleteEventWithId(self,calId):
       return self.service.events().delete(calendarId=self.clID, eventId=calId).execute()

    def GetUpComing(self,numOfEvents):
        #call the Calendar API
        now = datetime.datetime.utcnow().isoformat()+'Z' #Z indicates UTC time

        print(f"Getting the upcoming {numOfEvents} events")

        event_result = self.service.events().list(calendarId=self.clID,timeMin = now,
                                            maxResults=numOfEvents, singleEvents = True,
                                            orderBy='startTime').execute()

        events = event_result.get('items',[])

        if not events:
            print('No upcoming events found')
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(start,event['summary'])
        
        return events

    def createEvent(self,datetimeFrom,datetimeTo,color,description, title):

        event = {
            'summary':title,
            'location':'Kumi, Oslo, Norway',
            'colorId': color,
            
            'description':description,
            'start':{
                'dateTime':datetimeFrom ,
                'timeZone':'Europe/Oslo'
            },
            'end':{
                'dateTime':datetimeTo,
                'timeZone':'Europe/Oslo'
            },
            'reminders': {
                'useDefault': True,
            }
        }

        return event

    