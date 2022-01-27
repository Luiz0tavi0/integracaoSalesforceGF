from __future__ import print_function
import os.path
from pathlib import Path

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import requests
import datetime

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'credentials.json'
path_service_account_file = Path("credentials.json").resolve()

def get_caspio_token_access() -> str:
    url = "https://c2aca196.caspio.com/oauth/token"
    payload = "grant_type=client_credentials&client_id=af322ec0fa4e48e5d6da975c52b30a20129c6496892e616c9e&client_secret" \
              "=3722b2b5c83344d9ae21a91a909eef8788ac94141b6667d47e"

    headers = {
        'Content-Type': 'text/plain',
        'Cookie': 'AWSALB=vhgcWRnIUX2Z6uVKHDtpOC4YR8o5ddyFjb5llP00NVq6s23xj9BqeJ9TyvAaK8wcw3pJVy3oH6YyDD9SH+6EtFKuy2YKqx0mbHcW6KzH+1zaMDhmQPonEuFMMPqQ; AWSALBCORS=vhgcWRnIUX2Z6uVKHDtpOC4YR8o5ddyFjb5llP00NVq6s23xj9BqeJ9TyvAaK8wcw3pJVy3oH6YyDD9SH+6EtFKuy2YKqx0mbHcW6KzH+1zaMDhmQPonEuFMMPqQ'
    }

    response = requests.request("POST", url, headers=headers, data=payload)
    access_token = response.json().get('access_token')
    return access_token


def search_in_bazar_vw_oportunidades_salesforc(params: dict):
    token = get_caspio_token_access()
    payload = {}
    headers = {
        'accept': 'application/json',
        'Authorization': f"bearer {token}"
    }
    url = "https://c2aca196.caspio.com/rest/v2/views/bazar_vw_oportunidades_salesforc/records"

    response = requests.request("GET", url, headers=headers, data=payload, params=params)

    if response.status_code == 200:
        return response.json().get('Result')
    else:
        print("Erro na função search_in_bazar_vw_oportunidades_salesforc")
        print(response.json())
        return ""


def instantiate_spreadsheet():
    """"
    Instancia uma spreadsheet a partir de um token ou de credenciais.
    As credenciais foram geradas no Google plataform com a conta da zeta

    Mais informações em: https://developers.google.com/sheets/api/quickstart/python"""

    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                str(path_service_account_file), SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    service = build('sheets', 'v4', credentials=creds)
    # Call the Sheets API
    try:
        sheet = service.spreadsheets()
    except:
        sheet = ""

    return sheet


def main(RANGE_NAME: str, values: list):
    """" Função que recebe um range (no formato sheet!A:Z passando o nome da tabela e o range em que os valores serão adicionados)
    e uma list values que em que cada lista interna representa uma linha que vai ser adicionada na tabela.

    Cada linha precisa ter exatamente a mesma quantidade de valores premitida no RANGE_NAME
    """

    sheet = instantiate_spreadsheet()

    if not sheet:
        print("Falha ao instanciar a spreadsheet, verifique as credenciais")
        return False

    # The ID of "Leads Salesforce" spreadsheet.
    SPREADSHEET_ID = '1OAornGa3shYiKMXG1h7B_A0ROAYx3PWIjBpF5AeZ6qw'

    body = {
        'values': values
    }
    try:
        result = sheet.values().append(spreadsheetId=SPREADSHEET_ID,
                                       range=RANGE_NAME,
                                       body=body,
                                       valueInputOption="RAW",
                                       insertDataOption="INSERT_ROWS").execute()

        print(result)
    except:
        print(f"Erro ao adicionar linhas na tabela: \n {values} \n\n")


yesterday = datetime.datetime.now() - datetime.timedelta(days=1)

# Cria os parametros que vão ser utilizados  para fazer a pesquisa na view do caspio, usando a data do dia anterior e
# o status da doacão
params_agendamento = {
    'q.where': f"bazar_tb_doacao_Entry_DateUpdated LIKE '%{yesterday:%Y-%m-%d}%' AND bazar_tb_doacao_status_doacao = 'Cadastrada'"}
params_baixa = {
    'q.where': f"bazar_tb_doacao_Entry_DateUpdated LIKE '%{yesterday:%Y-%m-%d}%' AND bazar_tb_doacao_status_doacao = 'Concluída'"}

oportunidades_agendadas = search_in_bazar_vw_oportunidades_salesforc(params_agendamento)
oportunidades_coletadas = search_in_bazar_vw_oportunidades_salesforc(params_baixa)

print("\n[" + str(datetime.datetime.now()) + "] " + "Doações com Oportunidade Agendadas:")
# Se houver agendamento de doações com oportunidade no dia anterior
if oportunidades_agendadas:
    oportunidades_agendadas_rows = []

    for oportunidade in oportunidades_agendadas:
        row = [str(oportunidade['bazar_tb_doacao_id_doacao']),
               str(oportunidade['bazar_tb_contatos_id_doador']),
               str(oportunidade['bazar_tb_contatos_id_oportunidade_salesforce']),
               str(oportunidade['bazar_tb_doacao_Entry_DateUpdated'])]
        print(row)
        oportunidades_agendadas_rows.append(row)
    print("\nResultado da inserção na tabela de agendamentos:")
    main("agendamentos_de_coletas!A:D", oportunidades_agendadas_rows)
else:
    print("Sem novos agendamentos")

print("\n[" + str(datetime.datetime.now()) + "] " + "Doações com Oportunidade Baixadas:")
# Se houver baixa de doações com oportunidade no dia anterior
if oportunidades_coletadas:
    oportunidades_coletadas_rows = []

    # organiza os dados da doacão do caspio em rows que serão adicionadas
    for oportunidade in oportunidades_coletadas:
        row = [str(oportunidade['bazar_tb_doacao_id_doacao']),
               str(oportunidade['bazar_tb_contatos_id_doador']),
               str(oportunidade['bazar_tb_contatos_id_oportunidade_salesforce']),
               str(oportunidade['bazar_tb_doacao_Entry_DateUpdated'])]
        print(row)
        oportunidades_coletadas_rows.append(row)
    print("\nResultado da inserção na tabela de baixas:")
    main("baixas_de_coletas!A:D", oportunidades_coletadas_rows)
else:
    print("Sem novas baixas")

print("-------------------------------------------------------------------------------------------------")
