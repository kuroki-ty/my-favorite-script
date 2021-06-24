import os
from os.path import join, dirname
from dotenv import load_dotenv

dotenv_path = join(dirname(__file__), '.env')
load_dotenv(dotenv_path)

SA_KEY_PATH = os.environ.get('SA_KEY_PATH')
SPREADSHEET_ID = os.environ.get('SPREADSHEET_ID')
