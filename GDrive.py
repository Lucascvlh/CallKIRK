from __future__ import print_function

import os.path
import pandas as pd
import numpy as np
import json

from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/drive.metadata.readonly']

caminho = 'Matriz_Chamado.xlsx'

def sheet():
  planilha = pd.read_excel(caminho, sheet_name='Planilha1', usecols=[0,1,2], engine='openpyxl')
  planilha['Data de vencimento'] = pd.to_datetime(planilha['Data de vencimento'], format='%Y-%m-%d')
  planilha['Data de vencimento'] = planilha['Data de vencimento'].apply(lambda x: datetime.strftime(x, '%d/%m/%Y'))
  planilha['Gdrive'] = ''

  for i in range(len(planilha)):
    if np.isnan(planilha.at[i,'Documento']):
      continue
    else:
      planilha.at[i, 'Gdrive'] = 'Link do Drive aqui'

  array_dict = planilha.to_dict(orient='records')

  jsonKey = json.dumps(array_dict)
  return(jsonKey)

def search_files(name, mime_type, folder_id):
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('drive', 'v3', credentials=creds)

        query = f"name contains '{name}' and mimeType='{mime_type}' and parents in '{folder_id}'"
        results = service.files().list(
            q=query,
            fields="nextPageToken, files(id, name, webViewLink)").execute()
        items = results.get('files', [])

        if not items:
            print('No files found.')
            return('No files found.')
        else:
            print('Files:')
            for item in items:
                print(f"{item['name']} - {item['webViewLink']}")
                return item['webViewLink']
    except HttpError as error:
        print(f'An error occurred: {error}')

if __name__ == '__main__':
    jsonKey = json.loads(sheet())
    key = input('Chave da pasta: ')

    for i in range(len(jsonKey)):
        name = '{}.pdf'.format(jsonKey[i]["Documento"])
        mime_type = 'application/pdf'
        folder_id = key
        web_view_link = search_files(name, mime_type, folder_id)
        jsonKey[i]['Gdrive'] = web_view_link

    # Agora você pode converter o jsonKey novamente para JSON para salvá-lo em um arquivo ou fazer algo com os dados.
    jsonAPI = json.dumps(jsonKey, indent=2)

    planilha = pd.DataFrame(jsonKey)
    with pd.ExcelWriter(caminho, mode='w') as writer:
        planilha.to_excel(writer, sheet_name='Planilha1', index=False)


