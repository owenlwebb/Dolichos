"""Google Sheets API Wrapper

   Author: Owen Webb
   Date: 1/20/2020
"""
import os.path
import pickle
import re

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


class GoogleSheetsError(BaseException):
   pass


class GoogleSheets(object):
   def __init__(self, sheet_id, rw=True, fromfile=True):
      """Initialize and authenticate Google Sheets API.

         sheet_id : if fromfile=True name of file containing spreadsheetID,
                     if fromfile=False sheet_id is string literal
         rw : read/write access? False implies read-only
      """
      # load sheet_id and set scope
      if fromfile:
         with open(sheet_id, 'r') as fin:
            self.sheet_id = fin.read().strip()
      else:
         self.sheet_id = sheet_id

      # set scope
      if rw:
         self.scope = 'https://www.googleapis.com/auth/spreadsheets'
      else:
         self.scope = 'https://www.googleapis.com/auth/spreadsheets.readonly'

      # set file vars
      self.token_file = 'sheets_token.pickle'
      self.auth_file = 'sheets_auth.json'

      self.authenticate()

   def name_to_sheetid(self, name):
      """Return integer sheet id of a particular sheet within the 
      spreadsheet."""
      req = self.sheet.get(spreadsheetId=self.sheet_id, 
                           fields='sheets/properties')
      sheet_props = req.execute()

      for sheet in (s['properties'] for s in sheet_props['sheets']):
         if sheet['title'] == name:
            return sheet['sheetId']

      raise GoogleSheetsError(f'Could not find sheet: {name}')

   @staticmethod
   def column_letter_to_num(col):
      index = 0
      for i, char in enumerate(reversed(col)):
            index += ((ord(char) - 96) * (26**i))
      return index

   def a1_to_gridrange(self, a1coords):
      """Convert A1 notation coordinates of the form SheetName!COORDS to
      a GridRange object required by some v4 Sheets API calls.
      """
      try:
         name, coords = a1coords.split('!')
      except ValueError:
         msg = 'Coordinates must be of the form SHEETNAME!COORDS'
         raise GoogleSheetsError(msg)

      # assume only one cell and adjust to a range later if needed
      start_cell = coords.split(':')[0].lower()
      start_col, start_row = re.match('([a-z]+)([0-9]+)', start_cell).groups()

      startRowIndex = int(start_row) - 1
      endRowIndex = startRowIndex + 1
      startColumnIndex = self.column_letter_to_num(start_col) - 1
      endColumnIndex = startColumnIndex + 1

      if ':' in a1coords:
         end_col, end_row = re.match('([a-z]+)([0-9]+)', coords).groups()
         endRowIndex = end_row
         endColumnIndex = self.column_letter_to_num(end_col)
      
      target_id = self.name_to_sheetid(name)
      return {
         "sheetId" : target_id,
         "startRowIndex": startRowIndex,
         "endRowIndex": endRowIndex,
         "startColumnIndex": startColumnIndex,
         "endColumnIndex": endColumnIndex
      }

   def authenticate(self):
      """Authenticate the Google Sheets API, most code here stolen from:
      https://developers.google.com/resources/api-libraries/documentation/sheets/v4/python/latest/sheets_v4.spreadsheets.values.html"""
      creds = None
      if os.path.exists(self.token_file):
         with open(self.token_file, 'rb') as token:
            creds = pickle.load(token)
      # If there are no (valid) credentials available, let the user log in.
      if not creds or not creds.valid:
         if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
         else:
            flow = InstalledAppFlow.from_client_secrets_file(self.auth_file, 
                                                             self.scope)
            creds = flow.run_local_server(port=0)
         # Save the credentials for the next run
         with open(self.token_file, 'wb') as token:
            pickle.dump(creds, token)

      # see documentation at: 
      self.sheet = build('sheets', 'v4', credentials=creds).spreadsheets()

   def get_cell(self, coords):
      """Return value of cell belonging to the authenticated sheet."""
      req = self.sheet.values().get(spreadsheetId=self.sheet_id, range=coords)
      res = req.execute()
      return res.get('values', [])[0][0]

   def set_cell(self, coords, val):
      """Set value of cell belonging to the authenticated sheet."""
      body = { 'range' : coords, 'values' : [[val]] }
      req = self.sheet.values().update(spreadsheetId=self.sheet_id, 
                                       range=coords, 
                                       body=body, 
                                       valueInputOption='USER_ENTERED')
      req.execute()

   def get_cell_note(self, coords):
      """Get note at the single cell belonging to coords in the authenticated 
         sheet.

         coords should be in A1 format (eg. Owen!A1:A1)
      """
      req = self.sheet.get(spreadsheetId=self.sheet_id,
                           ranges=coords,
                           fields='sheets/data/rowData/values/note').execute()
      return req['sheets'][0]['data'][0]['rowData'][0]['values'][0]['note']

   def set_cell_note(self, coords, text, append=True):
      """Set note at the single cell belonging to coords in the authenticated
      sheet. If append==True, text will be appended to the existing note.
      """
      req_body = {
         "requests": [
            {
                  "repeatCell": {
                     "range": self.a1_to_gridrange(coords),
                     "cell": {"note": text},
                     "fields": "note",
                  }
            }
         ]
      }
      req = self.sheet.batchUpdate(spreadsheetId=self.sheet_id, body=req_body)
      req.execute()


if __name__ == '__main__':
   # TESTING
   # gs = GoogleSheets('the_doc.identity')
   # gs.get_cell_note('abc', 'def')
   GoogleSheets('the_doc.identity').set_cell_note('Owen!a1', "")
