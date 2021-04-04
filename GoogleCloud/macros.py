import locale
import apiclient
import httplib2
import sys
from oauth2client.service_account import ServiceAccountCredentials
import os
import argparse
parser = argparse.ArgumentParser(epilog="""
                         instruction: add the arguments in strict sequence - sheet name→ range data→ secret key→ 
                         id sheet in link→ filename.\n
                         example: [path]/python3 macros.py Лист1 A1:C7 secret_key.json 273839dss3dad033901 report.txt
                         """)
parser.add_argument("name_sheet", help="name sheet in Google Sheets")
parser.add_argument("range_data", help="specify the data range in the sheet in Google Sheets")
parser.add_argument("secret_key", help="the name of the downloaded file with the private key")
parser.add_argument("id_sheet", help="the id name from your table link")
parser.add_argument("file_name", help="file name with its extension")
args = parser.parse_args()

locale.setlocale(locale.LC_TIME, 'ru_RU.UTF-8')
# 1← имя скаченного файла с закрытым ключом
CREDENTIALS_FILE = args.secret_key

if os.path.exists(os.path.join(os.getcwd(), args.secret_key)):
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                                   ['https://www.googleapis.com/auth/spreadsheets',
                                                                    'https://www.googleapis.com/auth/drive'])
else:
    print('Передан несуществующий ключ. Файл должен быть в текущем папке в формате json')
    quit()
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)
# 2← имя id из ссылки вашей таблицы
range_name = args.name_sheet + '!' + args.range_data
spreadsheetId = args.id_sheet
try:
    table = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=range_name).execute()
except Exception as e:
    print(f'Переданы невалидные значения {spreadsheetId} {range_name}. Должно быть в формате - A1:С7 или правильный id',
          e.args)
    quit()
lines = table.get('values', None)
if not lines:
    print(f'Такого листа {args.name_sheet} у вас нету')
    quit()
if os.path.exists(os.path.join(os.getcwd(), args.file_name)):
    with open(file=args.file_name, mode='w+', encoding='utf-8') as f:
        strings = ''
        for line in lines:
            if '' in line:
                continue
            strings += '\t'.join(line)+'\n'
        f.write(strings)
    sys.stdout.write(strings)
else:
    print('Данного файла не существует. Файл должен быть в текущем папке')
