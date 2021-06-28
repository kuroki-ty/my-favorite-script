from google import spreadsheets as spsh
from config import config

import os.path
import argparse
import datetime
import locale
import json

COLUMN_LIST = {
    'PRIVATE': {'alpha': 'B', 'index': 1},
    'WORK': {'alpha': 'C', 'index': 2},
    'MEMO': {'alpha': 'D', 'index': 3},
}

def get_spreadsheet():
    cred_dir = os.path.join(os.path.dirname(__file__), config.SA_KEY_PATH)
    return spsh.GoogleSpreadSheet(config.SPREADSHEET_ID, cred_dir)

def write_diary(date, text, column, force):
    year, month, day = toYYYYMMDDElement(date)
    target = toFormatDateForDiary(year, month, day)

    spreadsheet = get_spreadsheet()
    diary = spreadsheet.read_table(year, 'A', '1', 'E', '367')
    target_index, target_row = [(index, row) for index, row in enumerate(diary) if row[0] == target][0]

    col_index = COLUMN_LIST.get(column).get('index')
    is_writable = col_index >= len(target_row) or not target_row[col_index]

    response = f'{target}  {column}  に対して\n  {text}\nを書き込みます\n\n'
    if is_writable or force:
        updated_range = spreadsheet.write_value(text, year, COLUMN_LIST.get(column).get('alpha'), target_index + 1)
        response += f'...... {updated_range}  に正常に書き込まれました'
    else:
        response += '...... 書き込みに失敗しました\n書き込み対象のセルにはすでに値が入力されています'
    return response

def read_diary(date):
    year, month, day = toYYYYMMDDElement(date)
    target = toFormatDateForDiary(year, month, day)

    spreadsheet = get_spreadsheet()
    diary = spreadsheet.read_table(year, 'A', '1', 'E', '367')

    target_index = [index for index, row in enumerate(diary) if row[0] == target][0]
    days, *contents = diary[target_index]

    response = f'{days}\n'
    for index, item in enumerate(contents):
        response += f'  {diary[0][index + 1]}: {item}\n'

    return response

def toYYYYMMDDElement(date):
    year, month, day = [int(v) for v in date.split('-')]
    return [year, month, day]

def toFormatDateForDiary(year, month, day):
    return datetime.datetime(year, month, day).strftime('%Y/%m/%d (%a)')

def main():
    locale.setlocale(locale.LC_TIME, 'ja_JP.UTF-8')

    today = datetime.datetime.today().strftime('%Y-%m-%d')
    date_type = lambda s: datetime.datetime.strptime(s, '%Y-%m-%d')

    # コマンドライン引数設定
    parser = argparse.ArgumentParser(description='ヒトコト日記 - Google Spreadsheet に対しての読み込み、書き込み操作を行う')
    parser.add_argument('-t', '--text', default='', help='書き込みたい文字列を指定する')
    parser.add_argument('-d', '--date', type=date_type, default=today, help='読み書き対象の日記の日付 yyyy-mm-dd')
    parser.add_argument('-c', '--column', default='PRIVATE', choices=['PRIVATE', 'WORK', 'MEMO'], help='書き込むカラムを指定する')
    parser.add_argument('-f', '--force', default=False, action='store_true', help='書き込み操作実行時に値の上書きを強制する')

    args = parser.parse_args()
    formated_date = args.date.strftime('%Y-%m-%d')
    text = args.text

    if not text:
        response = read_diary(formated_date)
        print(response)
    else:
        response = write_diary(formated_date, text, args.column, args.force)
        print(response)

if __name__ == '__main__':
    main()
