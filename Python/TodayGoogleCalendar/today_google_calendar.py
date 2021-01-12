from __future__ import print_function
import datetime
from datetime import timedelta
from dateutil import parser
import pickle
import os.path

from httplib2 import Http
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

JST = datetime.timezone(datetime.timedelta(hours=+9), 'JST')

class Event:
    def __init__(self, title, start, end):
        self.title = title
        self.start = start
        self.end = end

    def output(self):
        return '- ' + self.start + ' ~ ' + self.end + ' : ' + self.title

def authorize_google_service(cred_dir):
    SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

    creds = None
    pickle_path = os.path.join(cred_dir, 'token.pickle')
    if os.path.exists(pickle_path):
        with open(pickle_path, 'rb') as token:
            creds = pickle.load(token)
    if not creds:
        # credentials.json was created by GCP service account
        cred_path = os.path.join(cred_dir, 'credentials.json')
        creds = ServiceAccountCredentials.from_json_keyfile_name(cred_path, scopes=SCOPES)
        with open(pickle_path, "wb") as token:
            pickle.dump(creds, token)

    http_auth = creds.authorize(Http())
    service = build('calendar', 'v3', http=http_auth)
    return service

    today = datetime.date.today()
    yesterday = datetime.datetime.today() - timedelta(days=1)
    time_min = datetime.datetime(yesterday.year, yesterday.month,
                                 yesterday.day, 15, 0, 0).isoformat() + 'Z'
    time_max = datetime.datetime(today.year, today.month,
                                 today.day, 14, 59, 59).isoformat() + 'Z'

    events_result = service.events().list(calendarId=calendarId,
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
