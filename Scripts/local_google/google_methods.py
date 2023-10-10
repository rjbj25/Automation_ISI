from googleapiclient.discovery import build
from google.oauth2 import service_account
import logging
from datetime import date
import sys
from pathlib import Path
BASE_DIR=Path(__file__).resolve().parent.parent.parent
sys.path.append(str(BASE_DIR))
from Scripts import common
logger = common.config()['logger']
logging.basicConfig(filename=f'{common.config()["config"]["routes"]["logging"]}{date.today()}.log', format='%(asctime)s.%(msecs)03d %(levelname)s {%(module)s} [%(funcName)s] %(message)s', datefmt='%Y-%m-%d,%H:%M:%S' ,level=logging.INFO)


SAF = common.config()['base_dir']/'Scripts/local_google/keys/keys.json'
SCOPES = common.config()['config']['google']['scopes']['global']
SPID=common.config()['config']['google']['spids']['test']
creds = service_account.Credentials.from_service_account_file(SAF,scopes=SCOPES)


def get_data_from_sheet():
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    lstrw = sheet.values().get(spreadsheetId=SPID, range=f'Data!A:A').execute()
    lstcol = 'R'
    long = common.get_len(lstrw)
    headers=f'Data!A1:{lstcol}1'
    data=f'Data!A{long}:{lstcol}{long}'

    dfn = common.get_data_from_sheets(sheet,[headers,data],SPID)
    return dfn


if __name__ == "__main__":
    print(get_data_from_sheet())