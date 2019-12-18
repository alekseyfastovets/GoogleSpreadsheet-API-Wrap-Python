#Полезные ссылки:
#https://habr.com/ru/post/305378/
#https://developers.google.com/sheets/api/reference/rest/
#https://github.com/Tsar/Spreadsheet/blob/master/Spreadsheet.py
#https://console.developers.google.com/project


import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials



class sheetAPI:

  tableId = 0
  vSheet = []

  def __init__(self,CREDENTIALS_FILE):
    self.credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive'])
        
    self.httpAuth = self.credentials.authorize(httplib2.Http())
        
    self.service = apiclient.discovery.build('sheets', 'v4', http = self.httpAuth)

    self.driveService = apiclient.discovery.build('drive', 'v3', http = self.httpAuth)


  def createTable(self,title,locale,nCol,nRow, sheetId = 0):
        
    query = self.service.spreadsheets().create(body = {
    'properties': {'title': title, 'locale': locale},
    'sheets': [{'properties': {'sheetType': 'GRID',
    'sheetId': sheetId,
    'title': title,
    'gridProperties': {'rowCount': nRow, 'columnCount': nCol}}}]
    })
    result = query.execute()
    self.tableId, self.vSheet = result['spreadsheetId'], result
    

  def getPermissions(self, vType = 'anyone',vRole = 'reader', emailAddress = 'none',spreadsheetId = 'none'):
    if spreadsheetId == 'none':
        spreadsheetId = self.tableId    
        
    if spreadsheetId == 0:
      print('Укажите корректный id таблицы')
    elif vType == 'anyone':
      drive = self.driveService.permissions().create(
      fileId = self.tableId,
      # доступ на чтение кому угодно
      body = {'type': vType, 'role': vRole},  
      fields = 'id'
      ).execute()
    elif vType == 'user' and emailAddress != 'none':
      drive = self.driveService.permissions().create(
      fileId = self.tableId,
      #доступ конкретному пользователю {'type': 'user', 'role': 'writer', 'emailAddress': 'pochtamorion@gmail.com'}
      body = {'type': vType, 'role': vRole, 'emailAddress': emailAddress},
      fields = 'id'
      ).execute()

  def updateSheet(self,vRange,data,spreadsheetId = 'none'):
  	#if spreadsheetId == 'none':
  		#spreadsheetId = self.tableId
    results = self.service.spreadsheets().values().batchUpdate(spreadsheetId = spreadsheetId, body = {
    "valueInputOption": "USER_ENTERED",
    "data": [
        {"range": vRange,
         "majorDimension": "ROWS",     # сначала заполнять ряды, затем столбцы (т.е. самые внутренние списки в values - это ряды)
         "values": data}

   
    ]
    }).execute()

  def sheetInfo(self,sId,meta_data):
  	request = self.service.spreadsheets().developerMetadata().get(spreadsheetId=sId, metadataId=meta_data).execute()
  	return request

  def getDataFromSheet(self,spreadsheetId,vRange):
    value_render_option = 'FORMATTED_VALUE'
    date_time_render_option = 'SERIAL_NUMBER'
    request = self.service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=vRange, valueRenderOption=value_render_option, dateTimeRenderOption=date_time_render_option)
    response = request.execute()
    return response

if __name__ == '__main__':
  props = {
  'spreadsheetId' : '',
  'GOOGLE_CREDENTIALS_FILE' : '',
  'locale' : 'ru_RU',
  'nCol' : 5,
  'nRow' : 100,
  'title' : 'TestSheet',
  'permissions_type' : ['anyone','user'],
  'permissions_role' : ['reader','writer'],
  'emailAddress' : ''
  }

  CREDENTIALS_FILE = ''  # имя файла с закрытым ключом
  test = sheetAPI(props.get('GOOGLE_CREDENTIALS_FILE'))
  test.createTable(
    props.get('title'),
    props.get('locale'),
    props.get('nCol'),
    props.get('nRow')
  )
  print(test.vSheet)
  
  test.getPermissions()
  






    
    
        
        
  