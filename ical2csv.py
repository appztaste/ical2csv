import sys
import os.path
from icalendar import Calendar
import csv
from pathlib import Path

filePath = Path(sys.argv[1])
outputFileName = str(filePath.absolute()) + '/output.csv'
headers = ('Summary', 'UID', 'Description', 'Location', 'Start Time', 'End Time', 'URL')



class CalendarEvent:
    """Calendar event class"""
    summary = ''
    uid = ''
    description = ''
    location = ''
    start = ''
    end = ''
    url = ''

    def __init__(self, name):
        self.name = name

events = []


def open_cal(fileName):
    if os.path.isfile(fileName):
        f = open(fileName, 'rb')
        gcal = Calendar.from_ical(f.read())

        for component in gcal.walk():
            event = CalendarEvent("event")
            if component.get('SUMMARY') == None: continue #skip blank items
            event.summary = component.get('SUMMARY')
            event.uid = component.get('UID')
            if component.get('DESCRIPTION') == None: continue #skip blank items
            event.description = component.get('DESCRIPTION')
            event.location = component.get('LOCATION')
            if hasattr(component.get('dtstart'), 'dt'):
                event.start = component.get('dtstart').dt
            if hasattr(component.get('dtend'), 'dt'):
                event.end = component.get('dtend').dt


            event.url = component.get('URL')
            events.append(event)
        f.close()
    else:
        print("I can't find the file ", fileName, ".")
        print("Please enter an ics file located in the same folder as this script.")
        exit(0)


def csv_write(outputFileName):
    try:
        with open(outputFileName, 'w') as myfile:
            wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
            wr.writerow(headers)
            for event in events:
                values = (event.summary.encode('utf-8'), event.uid, event.description.encode('utf-8'), event.location, event.start, event.end, event.url)
                wr.writerow(values)
            print("Wrote to ", outputFileName, "\n")
    except IOError:
        print("Could not open file! Please close Excel!")
        exit(0)


def debug_event(class_name):
    print("Contents of ", class_name.name, ":")
    print(class_name.summary)
    print(class_name.uid)
    print(class_name.description)
    print(class_name.location)
    print(class_name.start)
    print(class_name.end)
    print(class_name.url, "\n")

for fileName in filePath.glob('*.ics'):
    open_cal(fileName)
    
csv_write(outputFileName)

#debug_event(event)
