from CalendarEvent import CalendarEvent
from ExcelParser import ExcelParser
import datetime

class CalendarEventManager():
    
    def __init__(self, credentials):
        self.calendarEvent = CalendarEvent(credentials)

    def deleteEventsInDateRange(self,dateFrom,dateTo):
        
            event_results = self.calendarEvent.GetEventsInDateRange(dateFrom,dateTo)

            eventItems = event_results.get('items',[])

            eventsToDelete=[]

            for event in eventItems:
                if 'location' in event:
                    if "Kumi" in event['location']:
                        eventsToDelete.append(event['id'])
            
            for eventId in eventsToDelete:
                deleted = self.calendarEvent.DeleteEventWithId(eventId)

                print(f"Event delete {deleted}")

class Rote():

    def __init__(self, credentials, excelfile):
        self.calendarEvent = CalendarEvent(credentials)
        self.excelParser = ExcelParser(excelfile)
    
    def createEvents(self, place,color):
        
        schedule = self.excelParser.getSchedule(place)

        for day, time in schedule.items():
            
            description=place+' Schedule \n'

            if time[0] == 'A':
                ATime = self.excelParser.getAuxLegend(day,time)
                description += f"check the roter, legend says {time}.\n"
                if ATime == "all day":
                    time="07:00-19:00"
                else :
                    time = ATime
            
            fromToTime = time.split('-')

            fromDate = self.excelParser.getDateTime(day,fromToTime[0]).isoformat()   
            toDate = self.excelParser.getDateTime(day,fromToTime[1]).isoformat()

            event = self.calendarEvent.createEvent(fromDate,toDate, color ,description,"Shift in the "+place)

            self.calendarEvent.InsertEvent(event)

if __name__ == "__main__":
    credentialsFilePath = 'credentials.json'
    excelFilePath = 'Shift list - Oct 2018.xlsx'
    
    kitchen_color = 10
    front_color = 5
    
    roter = Rote(credentialsFilePath,excelFilePath)

    roter.createEvents('Kitchen',kitchen_color)
    roter.createEvents('Front',front_color)

    # now = datetime.datetime(2018,10,1).isoformat()+'Z'

    # dateTo = datetime.datetime(2018,11,1).isoformat()+'Z'

    # calManager = CalendarEventManager(credentialsFilePath)
    # calManager.deleteEventsInDateRange(now,dateTo)
