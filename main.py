import json

#imports do selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time

#imports do para webscrapping sem selenium
from bs4 import BeautifulSoup
import requests as r

#imports para acessar a api do google calendar
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import os
from datetime import datetime, timedelta


#acessando AutoSAU, fazendo login e abrindo horarios e salas de aula para pegar todas as materias nas quais sou inscrito
service = Service('chromedriver_win32/chromedriver')
browser = webdriver.Chrome(service=service)

browser.get('https://sau.puc-rio.br/WebLoginPucOnline/Default.aspx?sessao=VmluY3Vsbz1BJlNpc3RMb2dpbj1QVUNPTkxJTkVfQUxVTk8mQXBwTG9naW49TE9HSU4mRnVuY0xvZ2luPUxPR0lOJlNpc3RNZW51PVBVQ09OTElORV9BTFVOTyZBcHBNZW51PU1FTlUmRnVuY01lbnU9TUVOVQ__')
time.sleep(3)
input_mat = browser.find_element(by=By.ID, value="txtLogin")
input_mat.send_keys("2211068")
input_senha = browser.find_element(by=By.ID, value="txtSenha")
input_senha.send_keys("*As^0evV") #nao me desmatricula em nenhuma materia por favor kkkkkkkk
botao_logar = browser.find_element(by=By.ID, value="btnOk")
botao_logar.click()
time.sleep(3)

horarios_aulas = browser.find_element(by=By.XPATH, value='//*[@id="trv_Area_ESQUERDA_UL"]/li[2]/a')
horarios_aulas.click()
time.sleep(3)

l_minhas_materias_response = browser.find_elements(by=By.CLASS_NAME, value="listaDisciplinasLinha")
l_minhas_materias = list()
for elem in l_minhas_materias_response:
    l_minhas_materias.append(str(elem.text.split("/")[0].strip()))

#fazendo login no ead pq nao tem data da prova de inf1009 no cbctc
browser.get("https://ead.puc-rio.br/loginccead/")
time.sleep(25)

input_mat = browser.find_element(by=By.ID, value="username")
input_mat.send_keys("2211068")
input_senha = browser.find_element(by=By.ID, value="password")
input_senha.send_keys("*As^0evV")
botao_logar = browser.find_element(by=By.ID, value="btnlogin")
botao_logar.click()
time.sleep(6)

browser.get("https://ead.puc-rio.br/course/view.php?id=78260")
time.sleep(5)
materia__ = browser.find_element(by=By.CLASS_NAME, value='h2').text.split("-")[0].strip()
l_minhas_provas = list()
for i in range(2,5):
    d_prova = dict()
    d_prova['materia'] = materia__
    d_prova['prova'] = browser.find_element(by=By.XPATH, value=f'//*[@id="coursecontentcollapse1"]/div/div[1]/div/div/table[2]/tbody/tr[{i}]/td[1]').text
    d_prova['data'] = browser.find_element(by=By.XPATH, value=f'//*[@id="coursecontentcollapse1"]/div/div[1]/div/div/table[2]/tbody/tr[{i}]/td[3]').text + '/2023'
    d_prova['horario'] = '13-15:00H'
    l_minhas_provas.append(d_prova)


#utilizando requests e beautifulsoup pra fazer web scrapping pq pra pegar todas as informaçoes sobre todas as provas que estao na tabela do cbctc nao achei uma forma de fazer no selenium
response_cbctc = r.get("https://www.cbctc.puc-rio.br/Paginas/MeuCB/Provas.aspx")
parsedHTML_cbctc = BeautifulSoup(response_cbctc.text, "html.parser")
l_materia_data_prova = parsedHTML_cbctc.find_all("tr")[1:]

#salvando as provas em dicionarios numa lista
for elem in l_materia_data_prova:
    materia = elem.find("td", {"data-title":"Disciplina"}).text.strip()
    if materia in l_minhas_materias:
        d_materia_prova = dict()
        d_materia_prova['materia'] = materia
        d_materia_prova['prova'] = elem.find("td", {"data-title":"Avaliação"}).text.strip()
        d_materia_prova['data'] = elem.find("td", {"data-title":"Data"}).text.strip()[0:10]
        d_materia_prova['horario'] = elem.find("td", {"data-title":"Horário"}).text.strip()
        l_minhas_provas.append(d_materia_prova)
json.dump(l_minhas_provas, open("provas.json","w"))

#salvando uma copia das provas no crud crud caso nao exista (essa parte so roda quando pego um novo link do crud crud ate por que as datas que estao no site do cbctc nao vao mudar)
CRUDCRUD_URL = "https://crudcrud.com/api/8d7f9064e1a1480ab0329df30ce70b8a/materias"        
if len(json.loads(r.get(CRUDCRUD_URL).text)) == 0:
    for d_materia_prova in l_minhas_provas:
        r.post(CRUDCRUD_URL, json=d_materia_prova)


SCOPES = [
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/calendar.readonly", 
    "https://www.googleapis.com/auth/calendar.events.readonly", 
    "https://www.googleapis.com/auth/calendar.events"
    ]

#criando credenciais e token (o programa verifica antes se o arquivo existe ou nao se nao existe cria e pega a credencial automaticamente do google toda vez que roda)
credenciais = None
if os.path.exists('token.json'):
    credenciais = Credentials.from_authorized_user_file('token.json', scopes=SCOPES)
if not credenciais or not credenciais.valid:
    if credenciais and credenciais.expired and credenciais.refresh_token:
        credenciais.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
        credenciais = flow.run_local_server(port=0)
    with open("token.json", "w") as token_json:
        token_json.write(credenciais.to_json())

#acessando API do google calendar e adicionando as provas que estavam no site do ciclo basico
service_calendar = build('calendar', 'v3', credentials=credenciais)
for prova in l_minhas_provas:
    ano = int(prova['data'][-4:])
    mes = int(prova['data'][3:5])
    dia = int(prova['data'][0:2])
    if prova['materia'] == 'MAT4202':
        titulo = f"ALG LIN 2 {prova['prova']}"    
    elif prova['materia'] == 'MAT1162':
        titulo = f'CALC 2 {prova["prova"]}'
    elif prova['materia'] == 'INF1009':
        titulo = f'LOGICA 1 {prova["prova"]}'
    if prova['horario'] == "Hor\u00e1rio de aula": #<-- apenas alg lin2 tem esse "horario" por isso automaticamente converti pro horario 9-11:00h
        hora_inicio = datetime(ano, mes, dia, 9, 0, 0)
        hora_fim = hora_inicio + timedelta(hours=2)
    else:
        hora_inicio = datetime(ano, mes, dia, int(prova['horario'][0:2]), 0, 0)
        hora_fim = datetime(ano, mes, dia, int(prova['horario'][3:5]), int(prova['horario'][6:8]),0)
    evento = {
        'summary': titulo,
        'colorId': '11',
        'location': 'PUC-Rio',
        'start': {
            'dateTime': hora_inicio.strftime("%Y-%m-%dT%H:%M:%S"),
            'timeZone': 'America/Sao_Paulo',
        },
        'end': {
            'dateTime': hora_fim.strftime("%Y-%m-%dT%H:%M:%S"),
            'timeZone': 'America/Sao_Paulo',
        },
        'reminders': 
        {'useDefault': False, 
        'overrides': [
                {'method': 'popup', 'minutes': 30}, 
                {'method': 'popup', 'minutes': 1440}, 
                {'method': 'popup', 'minutes': 5}, 
                {'method': 'popup', 'minutes': 10080}
                ],
        },
    }
    service_calendar.events().insert(calendarId="primary", body=evento).execute()