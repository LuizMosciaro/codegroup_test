from dataclasses import dataclass
import logging
import os
from time import sleep, time
import urllib3
from dotenv import load_dotenv
from pathlib import Path
import xlsxwriter
from openpyxl import load_workbook
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.relative_locator import locate_with
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.firefox import GeckoDriverManager
from typing import Union, Optional,List, Set, Dict, Tuple

load_dotenv()

@dataclass
class Pessoa:
    nome: str
    cidade: str
    estado: str
    descricao: str
    cargo: str
    empresa: str
    clima: str
    
class Crawler:
    def __init__(self,url:str,login:str,password:str,file:str,headless:Optional[bool]=None) -> None:
        self.url = url
        self.login = login
        self.password = password
        self.headless = headless
        self.file = file

    def run(self,):
        print('Inicializando busca de informações de usuários')
        logger = self.get_logger()
        driver = self.get_driver(self.headless,False)
        
        success = False
        tries = 0
        while not success and tries <=5:
            success = self.login_site(driver, logger, self.url, self.login, self.password)
            tries += 1

        if self.get_users_information(driver,logger,self.file):
            self.send_email(os.getenv('RECIPIENT'),'Pesquisa de usuarios','Ola, segue os arquivos em anexo',self.file)

    def get_logger(self,):
        logging.basicConfig(level=logging.WARNING,filename='file.log')
        log = logging.getLogger('Log')
        return log
    
    def get_driver(self, headless: bool = False, maximize: bool = True):
        # Configurações do driver
        firefox_options = Options()
        if headless:
            firefox_options.add_argument('-headless')  # Modo invisível
        firefox_options.add_argument('--ignore-certificate-errors')
        firefox_options.add_argument('--log-level=3')
        firefox_options.add_argument('--disable-gpu')

        pasta_download_relativa = str(os.path.join(Path.home(), "Downloads"))
        firefox_options.set_preference("browser.download.folderList", 2)
        firefox_options.set_preference("browser.download.dir", pasta_download_relativa)
        firefox_options.set_preference("browser.download.useDownloadDir", True)
        firefox_options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")

        # Configuração para o Geckodriver
        service = Service(GeckoDriverManager().install())
        driver = webdriver.Firefox(service=service, options=firefox_options)
        if maximize:
            driver.maximize_window()

        print('Configurações do webdriver realizadas')
        return driver

    def read_file(self,logger:logging.Logger,file:str):
        try:
            workbook = load_workbook(file)
            ws = workbook.active

            users = []
            for row in ws.iter_rows(min_row=2, max_col=7, values_only=True):
                users.append(Pessoa(*row))

            return users
        
        except Exception as Err:
            print(Err)
            logger.warning(f'Erro ao realizer leitura do arquivo: {file}\nErro: {Err}')

    def get_temperature(self,city:str,api_key:str):
        url = f"http://api.weatherapi.com/v1/current.json?key={api_key}&q={city},BR"

        response = requests.get(url)

        if response.status_code == 200:
            temperatura = response.json()['current']['temp_c']
            return str(temperatura)
        else:
           return f"Erro ao obter dados de clima: {response.status_code}"
    
    def login_site(self, driver: webdriver.Chrome, logger: logging.Logger, url: str, login: str, password: str) -> bool:
        print(f'Realizando Login: {url}')
        success = False
        tries = 0
        while not success and tries <= 5:
            try:
                driver.get(url)
                sleep(5)
                exists = driver.find_element(By.ID, 'session_key')
                if exists:
                    success = True
                    if success:
                        # Credenciais
                        driver.find_element(By.ID, 'session_key').send_keys(login)
                        driver.find_element(By.ID, 'session_password').send_keys(password)
                        driver.find_element(By.XPATH, '//button[normalize-space()="Entrar"]').click()
                        WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.CLASS_NAME, 'search-global-typeahead__input')))
                        print('Usuário conectado')
            except NoSuchElementException:
                logger.warning(f'Elemento não encontrado. Tentativa {tries} falhou.')
                success = False
            tries += 1
        return success

    def get_users_information(self,driver:webdriver.Chrome,logger:logging.Logger,file:str) -> None:
        try:
            users = self.read_file(logger,file)
            users_filled_list = []
            success = True

            for user in users:
                input_box = driver.find_element(By.CLASS_NAME,'search-global-typeahead__input')
                input_box.send_keys(user.nome)
                input_box.send_keys(Keys.ENTER)

                #Confirmar aba 'Pessoas'
                WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,'//button[normalize-space()="Pessoas"]')))
                driver.find_element(By.XPATH,'//button[normalize-space()="Pessoas"]').click()

                #Acessando o perfil do usuario
                print('Acessando perfil: ',user.nome)
                WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME,'entity-result__item')))
                
                if user.nome == 'Tiago Pamplona':
                    elements = driver.find_elements(By.CLASS_NAME,'entity-result__item')
                    if len(elements) >= 3:
                        third_element = elements[2]
                        third_element.click()
                else:
                    driver.find_element(By.CLASS_NAME,'entity-result__item').click()

                WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,"//h2/span[contains(text(),'Sobre')][1]")))

                #Descrição
                descricao_text = driver.find_element(By.XPATH,'//div[contains(@class,"pv-shared-text-with-see-more full-width t-14 t-normal t-black display-flex align-items-center")]').text
                user.descricao = descricao_text.split('\n')[0]
                #Empresa
                empresa_text = driver.find_element(By.XPATH,'//span[contains(@class,"pv-text-details__right-panel-item-text")]').text
                user.empresa = empresa_text
                #Cargo
                cargo_text = driver.find_element(By.XPATH,'//div[contains(@class,"text-body-medium break-words")]').text
                user.cargo = cargo_text

                #Clima na cidade do usuario
                temperatura = self.get_temperature(user.cidade,os.getenv('WEATHER_API_KEY'))
                user.clima = temperatura + 'C'

                users_filled_list.append(user)
            
            driver.close()

            book = load_workbook(file)
            worksheet = book.active
            row = 2
            for user in users_filled_list:
                worksheet.cell(row=row, column=1).value = user.nome
                worksheet.cell(row=row, column=2).value = user.cidade
                worksheet.cell(row=row, column=3).value = user.estado
                worksheet.cell(row=row, column=4).value = user.descricao
                worksheet.cell(row=row, column=5).value = user.cargo
                worksheet.cell(row=row, column=6).value = user.empresa
                worksheet.cell(row=row, column=7).value = user.clima
                row += 1
            
            book.save(file)
            print(f'Dados salvos em: {file}')
            success = True

            return success
        except Exception as Err:
            print(Err)
            logger.warning(f'Erro ao pesquisar dados de usuario\nErro:{Err}')
                    
    def send_email(self, recipient:str, subject:str, text:str, file:str) -> None:
        # Configurações do email
        remetente = os.getenv('EMAIL')
        senha_remetente = os.getenv('EMAIL_PWD')

        msg = MIMEMultipart()
        msg['From'] = remetente
        msg['To'] = recipient
        msg['Subject'] = subject

        msg.attach(MIMEText(text, 'plain'))

        # Abre o arquivo anexo em modo binário
        with open(file, 'rb') as anexo:
            attachment = MIMEBase('application', 'octet-stream')
            attachment.set_payload(anexo.read())
            
            # Codifica o anexo em base64
            encoders.encode_base64(attachment)
            filename = file.split("\\")[-1]
            attachment.add_header('Content-Disposition', f'attachment; filename={filename}')
            msg.attach(attachment)

        # Configuração do servidor SMTP do Outlook
        servidor_smtp = 'smtp.office365.com'
        porta_smtp = 587

        with smtplib.SMTP(servidor_smtp, porta_smtp) as servidor:
            servidor.starttls()
            servidor.login(remetente, senha_remetente)
            servidor.send_message(msg)
            print('Email enviado com sucesso')


if __name__=='__main__':
    rpa = Crawler('https://www.linkedin.com/?original_referer=',os.getenv('EMAIL'),os.getenv('PASSWORD'),'Colaboradores.xlsx',False)
    rpa.run()
