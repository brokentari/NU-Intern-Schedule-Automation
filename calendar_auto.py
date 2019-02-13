from __future__ import print_function
import datetime
import pickle
import os.path
from pprint import pprint
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ["https://www.googleapis.com/auth/calendar"]


def set_work_event(month, day, sport, time):
    service = credentials_check()
    credentials_check()
    if sport != "WBB":
        location = "Matthews Arena"
    else:
        location = "Cabot Center"
    event = {
        "summary": "{}/{} {}".format(month, day, sport),
        "location": location,
        "description": "",
        "colorId": "3",
        "start": {
            "dateTime": "2019-0{}-{}T17:00:00-05:00".format(month, day),
            "timeZone": "America/New_York",
        },
        "end": {
            "dateTime": "2019-0{}-{}T21:00:00-05:00".format(month, day),
            "timeZone": "America/New_York",
        },
        "reminders": {
            "useDefault": False,
            "overrides": [{"method": "popup", "minutes": 30}],
        },
    }

    event = service.events().insert(calendarId="primary", body=event).execute()
    print("Event created: {} ({})".format(event["summary"], event.get("htmlLink")))


def credentials_check():
    creds = None
    if os.path.exists("token_calendar.pickle"):
        with open("token_calendar.pickle", "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials_calendar.json", SCOPES
            )
            creds = flow.run_local_server()
        with open("token_calendar.pickle", "wb") as token:
            pickle.dump(creds, token)

    service = build("calendar", "v3", credentials=creds)
    return service


def check_existing_event(event_summary):
    check_marker = 0
    now = datetime.datetime.utcnow().isoformat() + "Z"
    service = credentials_check()
    existing_events = service.events().list(calendarId="primary", timeMin=now).execute()
    for event in existing_events["items"]:
        if event["summary"] == event_summary:
            pprint(event["summary"])
            check_marker = 1
            break
        else:
            check_marker = 0
        
    if check_marker == 1:
        return True
    else:
        return False


if __name__ == "__main__":
    main()
