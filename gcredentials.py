import os.path
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

_r, _ = os.path.split(__file__)

# add credentials to the account
creds = ServiceAccountCredentials.from_json_keyfile_name(
    os.path.join(_r, 'secrets/nca-holiday-tools-dee322de04d1.json'), scope)

# authorize the clientsheet
client = gspread.authorize(creds)
