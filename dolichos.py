""" Dolichos

    Synchronize Strava activities with "The Doc"

    Author: Owen Webb (owebb@umich.edu)
    Date: 1/13/2020
"""
import pickle
import os.path
from dataclasses import dataclass
from datetime import datetime
from string import ascii_uppercase

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

import strava

##################
# GLOBAL OPTIONS #
##################
THE_DOC_ID = ''
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']



@dataclass
class DocActivity:
    """Store activity data relevant to The Doc."""
    title: str
    description: str
    distance: float
    date: str


def main():
    # get most recent activities and their identifiers
    s = strava.Strava()
    recents = s.get_activity_list()
    identifiers = [act['id'] for act in recents]

    # get full data for each activity to synchronize or update
    to_sync = []
    for act_id in identifiers:
        act = s.get_detailed_activity(act_id)
        to_sync.append(DocActivity(act['name'], 
                                   act['description'], 
                                   meters_to_miles(act['distance']),
                                   tstamp_to_dt(act['start_date'])))

    # send to The Doc
    gsheets = gsheets_init()
    body = {"range":'Owen!C3', "values":[["hello, world"]]}
    cells = gsheets.values().update(spreadsheetId=THE_DOC_ID, range='Owen!C3', body=body, valueInputOption='USER_ENTERED').execute()
    #value = cells.get('values', [])
    #print(value)


def gsheets_init():
    """Initialize/Authenticate the Google Sheets API. Most code here is stolen 
    from Google's quickstart.py sample code."""
    global THE_DOC_ID
    with open('the_doc.identity', 'r') as fin:
        THE_DOC_ID = fin.read().strip()

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('sheets_token.pickle'):
        with open('sheets_token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'sheets_auth.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('sheets_token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # return sheet
    # see documentation at: 
    # https://developers.google.com/resources/api-libraries/documentation/sheets/v4/python/latest/sheets_v4.spreadsheets.values.html
    sheet = service.spreadsheets()
    return sheet


def meters_to_miles(meters):
    """Convert meters to miles. Conversion factor = 0.00062137."""
    return round(meters * 0.00062137, 2)


def tstamp_to_dt(timestamp):
    """Convert a Strava timestamp to a Python datetime object."""
    return datetime.strptime(timestamp, "%Y-%m-%dT%H:%M:%SZ")


def date_to_cell(timestamp):
    """Convert a date to the proper cell on The Doc."""
    start_date = datetime(2020, 1, 6)
    days_after = (timestamp - start_date).days

    row_num = ((days_after / 7) * 2) + 3
    col_num = (days_after % 7) + 3
    col_alpha = ascii_uppercase[col_num]

    return row_num, col_alpha


if __name__ == '__main__':
    main()
