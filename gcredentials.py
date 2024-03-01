# define the scope
import gspread
from gspread.utils import ExportFormat
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# add credentials to the account
creds = ServiceAccountCredentials.from_json_keyfile_name('secrets/nca-holiday-tools-dee322de04d1.json', scope)

# authorize the clientsheet
client = gspread.authorize(creds)
