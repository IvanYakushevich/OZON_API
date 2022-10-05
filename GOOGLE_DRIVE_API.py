# библиотеки для google drive API
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd


# функция для сохранения CSV для Google Drive
def xlsx_to_csv_pd():
    data_xls = pd.read_excel('ozon_fbs_data.xlsx')
    data_xls.to_csv('OZON_FBS_CSV.csv')

# функция перезаписи файла CSV в гугл драйв через pydrive
def upload_to_google_drive():
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)
    fileList = drive.ListFile({'q': "'root' in parents and mimeType='text/csv'"}).GetList()
    # Удаляем файл для загрузки обновленного файла
    for x in fileList:
        drive.CreateFile({'id': x['id']}).Delete()
    # Загружаем обновленный файл
    ozon_fabs_file = drive.CreateFile({"mimeType": "text/csv"})
    ozon_fabs_file.SetContentFile('OZON_FBS_CSV.csv')
    ozon_fabs_file.Upload()
    print('!!!Файл загружен и обновлен!!!')

# создание GOOGLE STEETS на google drive с данными из таблицы и
# # обновление нового файла google sheets OZON_FBS_TEST1 на основе CSV
def upload_to_google_drive_serv_acc():
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
            "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

    credentials = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    client = gspread.authorize(credentials)

    spreadsheet = client.open('OZON_FBS_TEST1')

    with open(r'OZON_FBS_CSV.csv', encoding='latin1') as file_obj:
        content = file_obj.read()
        client.import_csv(spreadsheet.id, data=content)


# Выполнение функций записи CVS и загрузки его в Google Drive, также записи и
# обновления нового файла google sheets OZON_FBS_TEST1 на основе CSV
xlsx_to_csv_pd()
#upload_to_google_drive()
upload_to_google_drive_serv_acc()
