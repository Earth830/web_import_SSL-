from google.oauth2 import service_account
from googleapiclient.discovery import build
import pandas as pd

def load_sheet(json_keyfile: str, spreadsheet_id: str, range_name: str) -> pd.DataFrame:
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds = service_account.Credentials.from_service_account_file(
        json_keyfile, scopes=SCOPES
    )
    service = build('sheets', 'v4', credentials=creds)
    result = (
        service.spreadsheets()
               .values()
               .get(spreadsheetId=spreadsheet_id, range=range_name)
               .execute()
    )
    values = result.get('values', [])
    if not values:
        return pd.DataFrame()
    df = pd.DataFrame(values[1:], columns=values[0])
    return df
