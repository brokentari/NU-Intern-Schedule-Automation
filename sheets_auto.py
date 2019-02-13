from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from pprint import pprint
import json
from calendar_auto import check_existing_event
from calendar_auto import set_work_event

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

calendar_spreadsheet = "1fsNioO3MI2sjDFGFJTXqh5R9tcvpgyyMxvEi3hFXXm4"
name_occurences = "Master Calendar!B12:J12"
work_details = "Master Calendar!B1:J1"

def main():
    creds = None
    if os.path.exists("token_sheets.pickle"):
        with open("token_sheets.pickle", "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials_sheets.json", SCOPES)
            creds = flow.run_local_server()
        with open("token_sheets.pickle", "wb") as token:
            pickle.dump(creds, token)
    
    service = build("sheets", "v4", credentials=creds)

    names = service.spreadsheets().get(spreadsheetId=calendar_spreadsheet, ranges=name_occurences, includeGridData=True)
    names_response = names.execute()

    details = service.spreadsheets().get(spreadsheetId=calendar_spreadsheet, ranges=work_details, includeGridData=True)
    details_response = details.execute()

    color_check = [1, 1, 1]
    range_dates = ord(work_details.split('!')[1].split(':')[1][0]) - ord(work_details.split('!')[1].split(':')[0][0])

    
    for name in range(range_dates):
        try:
            background_color = names_response["sheets"][0]["data"][0]["rowData"][0]["values"][name]["effectiveFormat"]["backgroundColor"]
            blue = background_color["blue"]
            red = background_color["red"]
            green = background_color["green"]
            
        except KeyError as error:
            blue, red, green = 0, 0, 0
        
        rgb = [red, blue, green]

        if rgb != color_check:
            work_day = details_response["sheets"][0]["data"][0]["rowData"][0]["values"][name]["effectiveValue"]["stringValue"]
            month = work_day.split(' ')[0].split('/')[0]
            day = work_day.split(' ')[0].split('/')[1]
            sport = work_day.split(' ')[1]
            sport_time = work_day.split(' ')[4]

            work_summary = "{}/{} {}".format(month, day, sport)

            if check_existing_event(work_summary) is False:
                set_work_event(month, day, sport, sport_time)


if __name__ == "__main__":
    main()
