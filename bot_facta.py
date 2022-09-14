from __future__ import print_function
from selenium import webdriver as opcoes
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from time import sleep
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

options = Options()
options.add_argument('--start-maximized')

driver = opcoes.Chrome(ChromeDriverManager().install(), options=options)
driver.implicitly_wait(120)
driver.get('https://desenv.facta.com.br/sistemaNovo/login.php')

driver.find_element(
    By.XPATH, '//*[@id="login"]').send_keys('login')

driver.find_element(By.XPATH, '//*[@id="senha"]').send_keys('senha')

driver.find_element(By.XPATH, '//*[@id="btnLogin"]').click()
# esperar aqui
while len(driver.find_elements(By.XPATH, '//*[@id="main-nav"]/div/ul/li[2]/a')) == 0:
    sleep(1)
driver.get('https://desenv.facta.com.br/sistemaNovo/propostaSimulador.php')

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1Fvyvn2RM1nGkq-trQP4zm-etC4ubj-pewT3MEUuioI0'
SAMPLE_RANGE_NAME = 'dados!A2:T'


def main():
    for i in range(50000):
        creds = None

        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'client_secret.json', SCOPES)
                creds = flow.run_local_server(port=0)

            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        values = result.get('values', [])

        for x in range(len(values)):

            nome = values[x][0]
            cpf = values[x][1]
            nascimento = values[x][2]
            rg = values[x][3]
            exp = values[x][4]
            uf = values[x][5]
            sexo = values[x][6]
            natura = values[x][7]
            pai = values[x][8]
            mae = values[x][9]
            cdf = values[x][10]
            cep = values[x][11]
            rua = values[x][12]
            bairro = values[x][13]
            nu = values[x][14]
            comp = values[x][15]
            cidade = values[x][16]
            fone = values[x][17]
            ag = values[x][18]
            conta = values[x][19]

            with open('clientes_digitados.txt', 'r') as clientes_digitados_r:
                clientes_lista = clientes_digitados_r.readlines()

            clientes_lista = list(
                map(lambda x: x.replace('\n', ''), clientes_lista))

            if not cpf in clientes_lista:
                driver.find_element(By.XPATH, '//*[@id="produto"]').click()

                driver.find_element(
                    By.XPATH, '//*[@id="produto"]/option[7]').click()

                driver.find_element(
                    By.XPATH, '//*[@id="tipoOperacao"]').click()

                driver.find_element(
                    By.XPATH, '//*[@id="tipoOperacao"]/option[6]').click()

                driver.find_element(By.XPATH, '//*[@id="averbador"]').click()

                driver.find_element(
                    By.XPATH, '//*[@id="averbador"]/option[4]').click()

                driver.find_element(By.XPATH, '//*[@id="banco"]').click()

                driver.find_element(
                    By.XPATH, '//*[@id="banco"]/option[2]').click()

                driver.find_element(By.XPATH, '//*[@id="cpf"]').send_keys(cpf)

                driver.find_element(
                    By.XPATH, '//*[@id="nomeCliente"]').send_keys(nome)

                driver.find_element(
                    By.XPATH, '//*[@id="valor"]').send_keys('16000')
                sleep(5)

                driver.find_element(
                    By.XPATH, '//*[@id="dataNascimento"]').click()
                driver.find_element(
                    By.XPATH, '//*[@id="dataNascimento"]').send_keys(nascimento)

                driver.find_element(By.XPATH, '//*[@id="opcaoValor"]').click()

                driver.find_element(
                    By.XPATH, '//*[@id="opcaoValor"]/option[3]').click()

                driver.find_element(
                    By.XPATH, '//*[@id="prazo"]').send_keys('24')

                driver.find_element(By.XPATH, '//*[@id="pesquisar"]').click()

                driver.find_element(
                    By.XPATH, '//*[@id="myTable"]/tbody/tr[2]').click()

                driver.find_element(By.XPATH, '//*[@id="etapa1"]').click()

                while len(driver.find_elements(By.XPATH, '//*[@id="vendedor"]')) == 0:
                    sleep(1)
                # -------------------------------------------------------------------

                driver.find_element(By.XPATH, '//*[@id="vendedor"]').click()

                driver.find_element(
                    By.XPATH, '//*[@id="vendedor"]/option[6]').click()

                driver.find_element(By.XPATH, '//*[@id="etapa_2"]').click()

                while len(driver.find_elements(By.XPATH, '//*[@id="sexo"]')) == 0:
                    sleep(1)

                # ----------------------------------------------------------------------

                driver.find_element(
                    By.XPATH, '//*[@id="sexo"]').send_keys(sexo)

                driver.find_element(
                    By.XPATH, '//*[@id="estadoCivil"]').send_keys('SOLTEIRO')

                driver.find_element(By.XPATH, '//*[@id="rg"]').clear()
                driver.find_element(By.XPATH, '//*[@id="rg"]').send_keys(rg)

                driver.find_element(
                    By.XPATH, '//*[@id="orgaoEmissor"]').send_keys('SSP')

                driver.find_element(
                    By.XPATH, '//*[@id="estadoRg"]').send_keys(uf)

                driver.find_element(By.XPATH, '//*[@id="emissaoRg"]').click()

                driver.find_element(
                    By.XPATH, '//*[@id="emissaoRg"]').send_keys(exp)

                driver.find_element(
                    By.XPATH, '//*[@id="nacionalidade"]').send_keys('BRASILEIRA')

                driver.find_element(
                    By.XPATH, '//*[@id="estadoNatural"]').send_keys(uf)

                driver.find_element(
                    By.XPATH, '//*[@id="cidadeNatural"]').click()
                sleep(3)
                for i in natura:
                    driver.find_element(
                        By.XPATH, '//*[@id="cidadeNatural"]').send_keys(i)

                driver.find_element(By.XPATH, '//*[@id="nomeMae"]').clear()
                driver.find_element(
                    By.XPATH, '//*[@id="nomeMae"]').send_keys(mae)

                driver.find_element(By.XPATH, '//*[@id="nomePai"]').clear()
                driver.find_element(
                    By.XPATH, '//*[@id="nomePai"]').send_keys(pai)

                driver.find_element(
                    By.XPATH, '//*[@id="valorPatrimonio"]').click()

                driver.find_element(
                    By.XPATH, '//*[@id="valorPatrimonio"]/option[2]').click()

                driver.find_element(
                    By.XPATH, '//*[@id="clienteAnalfabeto"]').send_keys('Não')

                driver.find_element(
                    By.XPATH, '//*[@id="matricula"]').send_keys(cdf)

                driver.find_element(
                    By.XPATH, '//*[@id="ufBeneficio"]').send_keys(uf)

                driver.find_element(
                    By.XPATH, '//*[@id="valorDoBeneficio"]').send_keys('400.00')

                driver.find_element(By.XPATH, '//*[@id="cep"]').clear()

                for x in range(9):
                    driver.find_element(
                        By.XPATH, '//*[@id="cep"]').send_keys(Keys.ARROW_LEFT)

                #mouse.typewrite('61760031', interval=0.1)

                driver.find_element(By.XPATH, '//*[@id="cep"]').send_keys(cep)

                driver.find_element(By.XPATH, '//*[@id="endereco"]').clear()
                driver.find_element(
                    By.XPATH, '//*[@id="endereco"]').send_keys(rua)

                driver.find_element(By.XPATH, '//*[@id="numero"]').clear()
                driver.find_element(
                    By.XPATH, '//*[@id="numero"]').send_keys(nu)

                driver.find_element(By.XPATH, '//*[@id="complemento"]').clear()
                driver.find_element(
                    By.XPATH, '//*[@id="complemento"]').send_keys(comp)

                driver.find_element(By.XPATH, '//*[@id="bairro"]').clear()
                driver.find_element(
                    By.XPATH, '//*[@id="bairro"]').send_keys(bairro)

                driver.find_element(By.XPATH, '//*[@id="cidade"]').click()
                sleep(3)
                for i in cidade:
                    driver.find_element(
                        By.XPATH, '//*[@id="cidade"]').send_keys(i)

                driver.find_element(By.XPATH, '//*[@id="celular"]').clear()
                driver.find_element(By.XPATH, '//*[@id="celular"]').click()

                for x in range(15):
                    driver.find_element(
                        By.XPATH, '//*[@id="celular"]').send_keys(Keys.ARROW_LEFT)

                #mouse.typewrite('85984064386', interval=0.1)

                driver.find_element(
                    By.XPATH, '//*[@id="celular"]').send_keys(fone)

                driver.find_element(By.XPATH, '//*[@id="etapa_3"]').click()

                while len(driver.find_elements(By.XPATH, '//*[@id="agenciaLiberacao"]')) == 0:
                    sleep(1)

                # -------------------------------------------------------------------------------

                driver.find_element(
                    By.XPATH, '//*[@id="agenciaLiberacao"]').send_keys(ag)

                driver.find_element(
                    By.XPATH, '//*[@id="contaLiberacao"]').send_keys(conta)

                driver.find_element(
                    By.XPATH, '//*[@id="id_tipo_profissao"]').click()

                driver.find_element(
                    By.XPATH, '//*[@id="id_tipo_profissao"]/option[5]').click()

                driver.find_element(By.XPATH, '//*[@id="etapa_4"]').click()

                # ----------------------------------------------------------------------------------

                driver.find_element(
                    By.XPATH, '//*[@id="content-wrapper"]/div/div/div/div[3]/div[2]/div[1]/input').click()

                driver.find_element(
                    By.XPATH, '//*[@id="btnRealizaFormalizacao"]').click()

                driver.find_element(
                    By.XPATH, '//*[@id="corpo"]/div[4]/div[2]/a').click()

                driver.find_element(By.XPATH, '//*[@id="novo"]').click()

                with open('clientes_digitados.txt', 'a') as clientes_digitados_r:
                    clientes_digitados_r.write(f'{cpf}\n')

            else:
                driver.refresh()
                print(f'cliente {cpf} que consta na lista nao foram digitados')
        driver.refresh()
        sleep(200)


if __name__ == '__main__':
    main()
