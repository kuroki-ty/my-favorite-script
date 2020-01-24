from __future__ import print_function
import datetime
from datetime import timedelta
from dateutil import parser
import pickle
import os.path

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
JST = datetime.timezone(datetime.timedelta(hours=+9), 'JST')

class Event:
    def __init__(self, title, start, end):
        self.title = title
        self.start = start
        self.end = end

    def output(self):
        return '- ' + self.start + ' ~ ' + self.end + ' : ' + self.title

def get_schedules():
    creds = None
    dir = os.path.dirname(__file__)
    pickle_path = os.path.join(dir, 'token.pickle')
    if os.path.exists(pickle_path):
        with open(pickle_path, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            cred_path = os.path.join(dir, 'credentials.json')
            flow = InstalledAppFlow.from_client_secrets_file(cred_path, SCOPES)
            creds = flow.run_local_server()
        with open(pickle_path, "wb") as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    today = datetime.date.today()
    yesterday = datetime.datetime.today() - timedelta(days=1)
    time_min = datetime.datetime(yesterday.year, yesterday.month,
                                 yesterday.day, 15, 0, 0).isoformat() + 'Z'
    time_max = datetime.datetime(today.year, today.month,
                                 today.day, 14, 59, 59).isoformat() + 'Z'

    events_result = service.events().list(calendarId='primary',
                                          timeMin=time_min,
                                          timeMax=time_max,
                                          singleEvents=True,
                                          orderBy='startTime').execute()
    events = events_result.get('items', [])
    schedules = []
    for event in events:
        iso_start = event['start'].get('dateTime')
        if iso_start is None:
            continue
        start = parser.parse(iso_start).astimezone(JST).strftime('%H:%M')
        iso_end = event['end'].get('dateTime')
        end = parser.parse(iso_end).astimezone(JST).strftime('%H:%M')

        output = Event(event['summary'], start, end).output()
        schedules.append(output)

    return schedules

def main():
    schedules = get_schedules()
    print('## 【スケジュール】本日の予定')
    for schedule in schedules:
        print(schedule)

if __name__ == '__main__':
    main()
