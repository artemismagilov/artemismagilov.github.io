from oauth2client.service_account import ServiceAccountCredentials
import apiclient
import httplib2
import argparse
import csv
import os
import re
import sys

parser = argparse.ArgumentParser(epilog="""
Instruction: add the arguments in strict sequence - [sheet name,sheet name2,...] or all→ [range data,range data2,...]→ 
secret key→ id sheet in link→ filename(optional). 
Example: [path]/python3 spreadsheet.py [Sheet1,Sheet2,...] [A1:C7,...] secret_key.json 273839dsdad033901 
[path/]report.[txt|csv](optional)
""")

parser.add_argument("name_sheet", help="name sheet in Google Sheets")
parser.add_argument("range_data", help="specify the data range in the sheet in Google Sheets")
parser.add_argument("secret_key", help="the name of the downloaded file with the private key")
parser.add_argument("id_sheet", help="the id name from your table link")
parser.add_argument("--file_name", "-f", help="file name with its extension (optional)", default=0)
args = parser.parse_args()

if os.path.exists(args.secret_key):
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        args.secret_key, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    )
else:
    print(f'secret_key {args.secret_key} not found.')
    quit()
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

if args.name_sheet == 'all':
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=args.id_sheet).execute()
    except:
        print('Wrong id_sheet. The link may be restricted in access ')
        quit()
    names = [sheet['properties']['title'] for sheet in spreadsheet['sheets']]
else:
    names = re.split(',', args.name_sheet.strip('[]'))
    assert '' not in names, f'Wrong sheet format - {names}'
ranges_data = re.split(',', args.range_data.strip('[]'))
assert '' not in ranges_data, f'Wrong range format - {ranges_data}'
for name, ranges in zip(names, ranges_data*len(names)):
    range_name = name+'!'+ranges
    try:
        table = service.spreadsheets().values().get(spreadsheetId=args.id_sheet, range=range_name).execute()
    except:
        print(f'Invalid values passed . Wrong range_name - {ranges} or wrong id_sheet - {args.id_sheet} or '
              f'wrong name_sheet - {name}. The link may be restricted in access ')
        quit()
    lines = table.get('values', None)
    if not lines:
        print(f'Empty sheet with title {name} and with a gap {ranges}')
        quit()
    if os.path.isfile(args.file_name) and args.file_name != 0:
        filename, file_extension = os.path.splitext(args.file_name)
        file = open(file=args.file_name, mode='a', newline='')
        if file_extension == '.csv':
            csv_writer = csv.writer(file)
            csv_writer.writerows(lines)
        elif file_extension == '.txt':
            for line in lines:
                file.write(' '.join(line) + '\n')
        file.close()
    elif not os.path.isfile(args.file_name) and args.file_name != 0:
        print(f'No such file: {args.file_name}. Optional-stdout')
        [sys.stdout.write(' '.join(line) + '\n') for line in lines]
    elif args.file_name == 0:
        print('You did not pass a parameter:file_name. Optional-stdout')
        [sys.stdout.write(' '.join(line) + '\n') for line in lines]
