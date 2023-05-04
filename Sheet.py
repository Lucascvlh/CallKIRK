import pandas as pd
from datetime import datetime
import numpy as np

def main():
  import json
  # Lendo a planilha em excel
  planilha = pd.read_excel('Robo2/excel.xlsx', sheet_name='excel', usecols=[0,1,2], engine='openpyxl')

  # Convertendo a coluna 'Data' para o tipo datetime
  planilha['Data de vencimento'] = pd.to_datetime(planilha['Data de vencimento'], format='%Y-%m-%d')

  # Formatando a coluna 'Data' para dd/mm/aaaa
  planilha['Data de vencimento'] = planilha['Data de vencimento'].apply(lambda x: datetime.strftime(x, '%d/%m/%Y'))

  #Adicionando a nova coluna e escrevendo nela
  planilha['Gdrive'] = ''

  for i in range(len(planilha)):
    if np.isnan(planilha.at[i,'Documento']):
      continue
    else:
      planilha.at[i, 'Gdrive'] = 'Link do Drive aqui'

  # Salvar o DataFrame atualizado na mesma planilha Excel
  with pd.ExcelWriter('Robo2/excel.xlsx', mode='w') as writer:
      planilha.to_excel(writer, sheet_name='excel', index=False)

  # Convertendo o dataframe para um array de dicionários
  array_dict = planilha.to_dict(orient='records')

  # Convertendo o array de dicionários para json
  jsonKey = json.dumps(array_dict)
  print(jsonKey)

if __name__ == '__main__':
  main()