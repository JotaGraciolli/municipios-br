import requests
import openpyxl
import time
from tqdm import tqdm 

# Declaração de variáveis
linha = 2
countMunicipio = 1
FilePath = r'C:\Users\sidnei.graciolli\OneDrive - Avanade\Documents\PythonProjects\Distancias\Cidades_Nosocomios.xlsx'
api_key = 'AIzaSyCboUFVqK8j-iUZoHsYXedLR9ZdvH3DvtI'
url = "https://maps.googleapis.com/maps/api/place/textsearch/json"

def Populate(pUrl, pParams, pWsRes, pWs, pRowNum):
    global linha

    response = requests.get(pUrl, params=pParams)
    results = response.json().get('results', [])

    for result in results:
        pWsRes.cell(row=linha, column=1).value = pWs.cell(row=pRowNum+2, column=4).value
        pWsRes.cell(row=linha, column=2).value = pWs.cell(row=pRowNum+2, column=2).value
        pWsRes.cell(row=linha, column=3).value = result.get('name')
        pWsRes.cell(row=linha, column=4).value = result.get('formatted_address')
        
        linha += 1

    return response.json().get('next_page_token')

def CreateSheet(pWorkbook, pSheetName):
    global linha
    global FilePath

    result = pWorkbook.create_sheet(str(pSheetName))
    result['A1'] = 'Município'
    result['B1'] = 'UF'
    result['C1'] = 'Nosocômio'
    result['D1'] = 'Endereço'

    linha = 2
    pWorkbook.save(FilePath)
    return result

DataFile = openpyxl.load_workbook(FilePath)

ws = DataFile['Base']

uf = ws['B2'].value

ufAnterior = uf

wsRes = CreateSheet(DataFile,uf)

totalMunicipio = ws.max_row

for RowNum, RowVal in tqdm(enumerate(ws.iter_rows(max_col=ws.max_column,min_row=2,max_row=totalMunicipio)), 
                           desc='Processando...', total=totalMunicipio):
    municipio = f'{ws.cell(row=RowNum+2, column=4).value}, {ws.cell(row=RowNum+2, column=2).value}, Brasil'
    
    #Verifica se mudou o estado
    uf = ws.cell(row=RowNum+2, column=2).value
    if uf != ufAnterior:
        ufAnterior = uf
        wsRes = CreateSheet(DataFile,uf)

    # Parâmetros da requisição
    params = {
        'query': f'hospitais em {municipio}',
        'key': api_key
    }

    next_page_token = Populate(url, params, wsRes, ws, RowNum)

    time.sleep(2)

    while next_page_token:
        params['pagetoken'] = next_page_token
        next_page_token = Populate(url, params, wsRes, ws, RowNum)

        time.sleep(2)

DataFile.save(FilePath)