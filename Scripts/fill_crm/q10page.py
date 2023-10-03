from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
#from selenium.webdriver.chrome.service import Service
from time import sleep
from selenium.webdriver.support.wait import WebDriverWait # Modulo para  manejo de esperas explicitas
from selenium.webdriver.support import expected_conditions as EC #esperas explicitas
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from datetime import date, timedelta
import logging
import sys
from pathlib import Path
BASE_DIR=Path(__file__).resolve().parent.parent.parent
sys.path.append(str(BASE_DIR))
from Scripts import common
from Scripts.local_google import google_methods
logger = common.config()['logger']
from datetime import date
logging.basicConfig(filename=f'{common.config()["config"]["routes"]["logging"]}{date.today()}.log', format='%(asctime)s.%(msecs)03d %(levelname)s {%(module)s} [%(funcName)s] %(message)s', datefmt='%Y-%m-%d,%H:%M:%S' , level=logging.INFO)

class Q10Page(object):

    def open(self):
        self._driver.get(self._url)

    def __init__(self):
        self._WINDOW_SIZE = "1920,1080"
        self._chrome_options = Options()
        self._chrome_options.add_argument("--window-size=%s" % self._WINDOW_SIZE)
        self._driver = webdriver.Chrome(options=self._chrome_options)
        self._driver.maximize_window()
        self._url = "https://site2.q10.com/login?ReturnUrl=%2F&aplentId=5f0cac06-a506-459a-a7b8-364b50574728"
        xpaths_path = common.config()['base_dir']/'Scripts/fill_crm/xpaths/xpaths.yml'
        self.xpaths = common.open_yaml(xpaths_path)
        self.student = google_methods.get_data_from_sheet()

    def login(self):
        WebDriverWait(self._driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, self.xpaths['TXTBOX_USERNAME'])))
        text_field_username = self._driver.find_element(By.CSS_SELECTOR, self.xpaths['TXTBOX_USERNAME'])
        text_field_password = self._driver.find_element(By.CSS_SELECTOR, self.xpaths['TXTBOX_PASSWORD'])

        text_field_username.clear()
        text_field_password.clear()
        user = common.config()['config']['q10user']['user']
        password = common.config()['config']['q10user']['password']
        text_field_username.send_keys(user)
        text_field_password.send_keys(password)

        self.click_element_by_ccs_sel(self.xpaths['BUTTON_SUBMIT'])
        self.click_element_by_ccs_sel(self.xpaths['BUTTON_ADMINISTRATIVOS'])

        sleep(38)

    def go_to_oportunities(self):
        self.click_element_by_ccs_sel(self.xpaths['MENU_INSTITUCIONAL'])
        WebDriverWait(self._driver, 3).until(EC.visibility_of_element_located((By.CSS_SELECTOR, self.xpaths['MENU_COMERCIAL_CRM'])))
        menu_comercial_crm = self._driver.find_element(By.CSS_SELECTOR, self.xpaths['MENU_COMERCIAL_CRM'])
        actions = ActionChains(self._driver)
        actions.move_to_element(menu_comercial_crm).perform()
        self.click_element_by_ccs_sel(self.xpaths['MENU_OPORTUNITIES'])

        sleep(2)

    def register_oportunitie(self):
        self.click_element_by_ccs_sel(self.xpaths['BUTTON_REGISTRAR_OPORTUNIDAD'])
        name = str(self.student['Nombre'].iloc[0].iloc[0])+' '+str(self.student['Apellido'].iloc[0].iloc[0])
        self.send_text_by_ccs_sel(self.xpaths['TXTBOX_OPOTUNIDAD_NOMBRE'], name)
        self.send_text_by_ccs_sel(self.xpaths['TXTBOX_IDENTIFICACION'], self.student['Cedula'].iloc[0])
        self.send_text_by_ccs_sel(self.xpaths['TXTBOX_EMAIL'], self.student['Correo'].iloc[0])
        self.send_text_by_ccs_sel(self.xpaths['TXTBOX_CELULAR'], self.student['Celular'].iloc[0])
        sleep(10)

    def send_text_by_ccs_sel(self, element, txt):
        WebDriverWait(self._driver, 3).until(EC.visibility_of_element_located((By.CSS_SELECTOR, element)))
        e = self._driver.find_element(By.CSS_SELECTOR, element)
        e.send_keys(txt)

    def click_element_by_ccs_sel(self, element):
        WebDriverWait(self._driver, 3).until(EC.visibility_of_element_located((By.CSS_SELECTOR, element)))
        e = self._driver.find_element(By.CSS_SELECTOR, element)
        e.click()

    def click_element_by_xpath(self, element):
        WebDriverWait(self._driver, 3).until(EC.visibility_of_element_located((By.XPATH, element)))
        e = self._driver.find_element(By.XPATH, element)
        e.click()