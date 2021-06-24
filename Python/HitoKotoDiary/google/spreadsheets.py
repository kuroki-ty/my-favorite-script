import pickle
import os.path
from httplib2 import Http
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

class GoogleSpreadSheet():
    def __init__(self, spreadsheet_id, cred_dir):
        self.service = self.authorize_service(cred_dir)
        self.spreadsheet_id = spreadsheet_id
        self.spreadsheets = self.service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()

    def authorize_service(self, cred_dir):
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

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
        service = build('sheets', 'v4', http=http_auth)
        return service

    def create_sheet(self, sheet_name):
        if (self.is_already_sheet(sheet_name)):
            return

        body = {
            'requests': [
                {
                    'addSheet': {
                        'properties': {
                            'title': sheet_name,
                        }
                    }
                }
            ]
        }
        dst_sheet = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheet_id, body=body).execute()['replies'][0]['addSheet']
        template_sheet = [sheet for sheet in self.spreadsheets['sheets'] if 'Template' in sheet['properties']['title']][0]
        self.copy_sheet(template_sheet, dst_sheet)

    def copy_sheet(self, src_sheet, dst_sheet):
        body = {
            'requests': [
                {
                    'copyPaste': {
                        'source': {
                            'sheetId': src_sheet['properties']['sheetId'],
                            'startRowIndex': 0,
                            'endRowIndex': src_sheet['properties']['gridProperties']['rowCount'],
                            'startColumnIndex': 0,
                            'endColumnIndex': src_sheet['properties']['gridProperties']['columnCount']
                        },
                        'destination': {
                            'sheetId': dst_sheet['properties']['sheetId'],
                            'startRowIndex': 0,
                            'endRowIndex': dst_sheet['properties']['gridProperties']['rowCount'],
                            'startColumnIndex': 0,
                            'endColumnIndex': dst_sheet['properties']['gridProperties']['columnCount']
                        }
                    }
                }
            ]
        }
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheet_id, body=body).execute()

    def is_already_sheet(self, sheet_name):
        sheet = [sheet for sheet in self.spreadsheets['sheets'] if sheet_name == sheet['properties']['title']]
        return len(sheet) != 0

    def read_table(self, sheet_name, anchor_col_alpha, anchor_row_num, end_col_alpha, end_row_num):
        range = f'{sheet_name}!{anchor_col_alpha}{anchor_row_num}:{end_col_alpha}{end_row_num}'
        return self.read_values(range)

    def read_values(self, range):
        response = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheet_id,
                                                            range=range).execute()
        return response['values']

    def write_table(self, title, df, sheet_name, index_title, anchor_col_alpha, anchor_row_num):
        df_anchor_row_num = str(int(anchor_row_num) + 1)
        self.write_value(title, sheet_name, anchor_col_alpha, anchor_row_num)
        self.write_dataframe_with_index_columns(df, sheet_name, index_title, anchor_col_alpha, df_anchor_row_num)

    def write_tables(self, title, df_list, sheet_name, index_title, anchor_col_alpha, anchor_row_num):
        df_index_anchor_row_num = str(int(anchor_row_num) + 1)
        df_anchor_row_num = str(int(anchor_row_num) + 2)
        self.write_value(title, sheet_name, anchor_col_alpha, anchor_row_num)
        for index, df in enumerate(df_list):
            if index == 0:
                self.write_dataframe_with_index_columns(df, sheet_name, index_title, anchor_col_alpha, df_index_anchor_row_num)
            else:
                self.write_dataframe_with_index(df, sheet_name, anchor_col_alpha, df_anchor_row_num)
            df_anchor_row_num = str(int(df_anchor_row_num) + len(df))

    def write_value(self, value, sheet_name, col_alpha, row_num):
        range = f'{sheet_name}!{col_alpha}{row_num}'
        list = [[value]]
        return self.update_values(range, list)

    def write_dataframe_with_index_columns(self, df, sheet_name, index_title, anchor_col_alpha, anchor_row_num):
        list = df.reset_index().T.reset_index().T.values.tolist()
        list[0][0] = index_title
        range = self.list2range(sheet_name, anchor_col_alpha, anchor_row_num, list)
        return self.update_values(range, list)

    def write_dataframe_with_index(self, df, sheet_name, anchor_col_alpha, anchor_row_num):
        list = df.reset_index().values.tolist()
        range = self.list2range(sheet_name, anchor_col_alpha, anchor_row_num, list)
        return self.update_values(range, list)

    def write_dataframe(self, df, sheet_name, anchor_col_alpha, anchor_row_num):
        list = df.values.tolist()
        range = self.list2range(sheet_name, anchor_col_alpha, anchor_row_num, list)
        return self.update_values(range, list)

    def update_values(self, range, list):
        body = {
            'majorDimension': 'ROWS',
            'values': list
        }
        response = self.service.spreadsheets().values().update(spreadsheetId=self.spreadsheet_id,
                                                               range=range,
                                                               valueInputOption='USER_ENTERED',
                                                               body=body).execute()
        return response['updatedRange']

    def alpha2num(self, alpha):
        num = 0
        for index, item in enumerate(list(alpha)):
            num += pow(26, len(alpha) - index - 1) * (ord(item) - ord('A') + 1)
        return num

    def num2alpha(self, num):
        if num <= 26:
            return chr(64 + num)
        elif num % 26 == 0:
            return self.num2alpha(num // 26 - 1) + chr(90)
        else:
            return self.num2alpha(num // 26) + chr(64 + num % 26)

    def list2range(self, sheet_name, anchor_col_alpha, anchor_row_num, list):
        anchor_col_num_int = self.alpha2num(anchor_col_alpha)
        len_row = len(list)
        len_col = [len(oneline) for oneline in list][0]
        end = self.num2alpha(anchor_col_num_int + len_col) + str(int(anchor_row_num) + len_row)
        range = f'{sheet_name}!{anchor_col_alpha}{anchor_row_num}:{end}'
        return range
