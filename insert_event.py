from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']


def insert_event(start_time='2019-11-11T16:00:00', end_time='2019-11-11T16:00:00', summary='あかり',
                 calendarid='4rju4fvdis8ti9cc2dckvc1hj8@group.calendar.google.com'):
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    event = {
        'summary': f'{summary}',
        'start': {
            'dateTime': f'{start_time}',
            'timeZone': 'Japan',
        },
        'end': {
            'dateTime': f'{end_time}',
            'timeZone': 'Japan',
        },
    }

    event = service.events().insert(calendarId=calendarid,
                                    body=event).execute()
    print(event['id'])


if __name__ == '__main__':
    insert_event()
