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
        '''
        Este metodo se encarga de abrir la pagina de Q10
        '''
        self._driver.get(self._url)

    def __init__(self):
        '''
        Este metodo se encarga de inicializar el driver de selenium y de abrir la pagina de Q10
        '''
        self._WINDOW_SIZE = "1920,1080"
        self._chrome_options = Options()
        self._chrome_options.add_argument("--window-size=%s" % self._WINDOW_SIZE)
        # headless
        self._chrome_options.add_argument("--headless")
        self._driver = webdriver.Chrome(options=self._chrome_options)
        self._driver.maximize_window()
        self._url = "https://site2.q10.com/login?ReturnUrl=%2F&aplentId=5f0cac06-a506-459a-a7b8-364b50574728"
        xpaths_path = common.config()['base_dir']/'Scripts/fill_crm/xpaths/xpaths.yml'
        self.xpaths = common.open_yaml(xpaths_path)

    def login(self):
        '''
        Este metodo se encarga de hacer el login en la pagina de Q10
        '''
        
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

        try:
            self.click_element_by_ccs_sel(self.xpaths['BUTTON_ADMINISTRATIVOS'])
        except Exception as e:
            logger.info('No se encontró mensaje de administrativos')
            pass
        self.logger_print('Login exitoso')
        #sleep(38)

    def go_to_oportunities(self):
        '''
        Este metodo se encarga de ir a la seccion de oportunidades en el CRM
        '''
        try:
            WebDriverWait(self._driver, 3).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#mensajeMora > div > div > div > div:nth-child(2) > div.col-lg-9.col-md-9.col-sm-9.col-xs-9')))
            WebDriverWait(self._driver, 40).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, '#mensajeMora > div > div > div > div:nth-child(2) > div.col-lg-9.col-md-9.col-sm-9.col-xs-9')))
        except Exception as e:
            logger.info('No se encontró mensaje de mora en pago')
            pass
        
        try:
            self.click_element_by_ccs_sel(self.xpaths['BUTTON_NO_SPACE'])    
        except Exception as e:
            logger.info('No se encontró mensaje de espacio insuficiente')
            pass

        self.click_element_by_ccs_sel(self.xpaths['MENU_INSTITUCIONAL'])
        WebDriverWait(self._driver, 3).until(EC.visibility_of_element_located((By.CSS_SELECTOR, self.xpaths['MENU_COMERCIAL_CRM'])))
        menu_comercial_crm = self._driver.find_element(By.CSS_SELECTOR, self.xpaths['MENU_COMERCIAL_CRM'])
        actions = ActionChains(self._driver)
        actions.move_to_element(menu_comercial_crm).perform()
        self.click_element_by_ccs_sel(self.xpaths['MENU_OPORTUNITIES'])
        self.logger_print('Seccion de oportunidades abierta')
        sleep(2)

    def register_oportunitie(self, student):
        '''
        Este metodo se encarga de registrar una oportunidad en el CRM
        '''
        self.click_element_by_ccs_sel(self.xpaths['BUTTON_REGISTRAR_OPORTUNIDAD'])
        name = str(student['Nombre'])+' '+str(student['Apellido'])
        self.send_text_by_ccs_sel(self.xpaths['TXTBOX_OPOTUNIDAD_NOMBRE'], name)
        self.send_text_by_ccs_sel(self.xpaths['TXTBOX_IDENTIFICACION'], student['Cedula'])
        self.send_text_by_ccs_sel(self.xpaths['TXTBOX_EMAIL'], student['Correo'])
        self.send_text_by_ccs_sel(self.xpaths['TXTBOX_CELULAR'], student['Celular'])
        print('Datos Generales de Oportunidad Registrados')

    def register_oportunitie_detail(self, student):
        '''
        Este metodo se encarga de registrar los detalles de una oportunidad en el CRM
        '''
        print(student)
        self.click_element_by_ccs_sel(self.xpaths['BUTTON_NEXT'])
        self.send_text_by_ccs_sel(self.xpaths['TXTBOX_DIRECCION'], student['Direccion'])
        self.click_element_by_ccs_sel(self.xpaths['SELECT_COMO_SE_ENTERO'])

        #Se selecciona como se entero el estudiante
        if student['Como_Se_Entero'] == 'Facebook Empresarial':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_FB_EMPRESARIAL'])
        elif student['Como_Se_Entero'] == 'Instagram':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_INSTAGRAM'])
        elif student['Como_Se_Entero'] == 'Pagina Web':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_PAGINA_WEB'])

        # Se selecciona el medio de contacto con el estudiante
        self.click_element_by_ccs_sel(self.xpaths['SELECT_MEDIO_DE_CONTACTO'])
        if student['Medio_De_Contacto'] == 'WhatsApp':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_WHATSAPP'])
        elif student['Medio_De_Contacto'] == 'Correo Electrónico':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CORREO_ELECTRONICO'])
        elif student['Medio_De_Contacto'] == 'Visita':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_VISITA'])
        elif student['Medio_De_Contacto'] == 'Llamada':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_LLAMADA'])

        # Se selecciona la procedencia, es decir hacia donde va el estudiante
        self.click_element_by_ccs_sel(self.xpaths['SELECT_PROCEDENCIA'])
        if student['Procedencia'] == 'INSTITUTO SUPERIOR DE INGENIERIA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_ISI'])
        elif student['Procedencia'] == 'GRUPO MAKRO PANAMA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GMK_PANAMA'])
        elif student['Procedencia'] == 'GRUPO MAKRO COLOMBIA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GMK_COLOMBIA'])
        elif student['Procedencia'] == 'GRUPO MAKRO CONVENIO':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GMK_CONVENIO'])


        # Se selecciona el estado del negocio
        self.click_element_by_ccs_sel(self.xpaths['SELECT_ESTADO_DEL_NEGOCIO'])
        if student['Estado_Del_Negocio'] == 'Presentación':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_PRESENTACION'])
        elif student['Estado_Del_Negocio'] == 'En Negociación':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_EN_NEGOCIACION'])
        elif student['Estado_Del_Negocio'] == 'Cierre':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CIERRE'])

        # Se selecciona el Asesor
        self.click_element_by_ccs_sel(self.xpaths['SELECT_ASESOR'])
        if student['Asesor'] == ' Yalock Ditta':
            self.click_element_by_ccs_sel(common.config()['config']['q10user']['yalock_selector'])
        elif student['Asesor'] == 'Miriam Martínez':
            self.click_element_by_ccs_sel(common.config()['config']['q10user']['miriam_selector'])
        elif student['Asesor'] == 'Nathaly Nieto':
            self.click_element_by_ccs_sel(common.config()['config']['q10user']['nathaly_selector'])
        elif student['Asesor'] == 'Tania Gelvez':
            self.click_element_by_ccs_sel(common.config()['config']['q10user']['tania_selector'])
        elif student['Asesor'] == 'Cristian Angulo':
            self.click_element_by_ccs_sel(common.config()['config']['q10user']['cristian_selector'])

        # Se selecciona el municipio del estudiante
        self.select_option_by_ccs_sel(self.xpaths['SELECT_MUNICIPIO'], student['Municipio'])

        #Se selecciona la carrera del estudiante
        self.click_element_by_ccs_sel(self.xpaths['SELECT_PROGRAMA'])
        if student['Programa'] == 'ASISTENTE DE INGENIERIA CIVIL Y DISEÑO DE OBRAS CIVILES':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_ASISTENTE_DE_INGENIERIA_CIVIL_Y_DISENO_DE_OBRAS_CIVILES'])
        elif student['Programa'] == 'CERTIFICACIONES ESPECIALIZADAS GRUPO MAKRO':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CERTIFICACIONES_ESPECIALIZADAS_GRUPO_MAKRO'])
        elif student['Programa'] == 'CURSO DE BUCEO COMERCIAL':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CURSO_DE_BUCEO_COMERCIAL'])
        elif student['Programa'] == 'CURSO DE FRANCES JOBS - CANADA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CURSO_DE_FRANCES_JOBS_CANADA'])   
        elif student['Programa'] == 'CURSO DE INGLES ISI':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CURSO_DE_INGLES_ISI'])
        elif student['Programa'] == 'CURSO DE INGLES JOBS - CANADA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CURSO_DE_INGLES_JOBS_CANADA'])
        elif student['Programa'] == 'CURSOS ISI':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CURSOS_ISI'])
        elif student['Programa'] == 'CURSOS NIVELATORIOS ISI':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CURSOS_NIVELATORIOS_ISI'])
        elif student['Programa'] == 'CURSOS SOBRE BUCEO ISI':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CURSOS_SOBRE_BUCEO_ISI'])
        elif student['Programa'] == 'DIPLOMADO EN MAQUINARIA PESADA CON ENFASIS EN MANTENIMIENTO Y PRODUCTIVIDAD.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_DIPLOMADO_EN_MAQUINARIA_PESADA_CON_ENFASIS_EN_MANTENIMIENTO_Y_PRODUCTIVIDAD'])
        elif student['Programa'] == 'DIPLOMADO EN SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EQUIPO PESADO ZEMER':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_DIPLOMADO_EN_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EQUIPO_PESADO_ZEMER'])
        elif student['Programa'] == 'FLYCAMDRONE - CURSO BASICO DE DRONES.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_FLYCAMDRONE_CURSO_BASICO_DE_DRONES'])
        elif student['Programa'] == 'FLYCAMDRONE - CURSO PROFESIONAL DE DRONES':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_FLYCAMDRONE_CURSO_PROFESIONAL_DE_DRONES'])
        elif student['Programa'] == 'GRUPO MAKRO REENTRENAMIENTO EN SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE EXCAVADORA HIDRÁULICA.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_REENTRENAMIENTO_EN_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_EXCAVADORA_HIDRAULICA'])
        elif student['Programa'] == 'GRUPO MAKRO REENTRENAMIENTO EN SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE EXCAVADORA SOBRE ORUGAS':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_REENTRENAMIENTO_EN_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_EXCAVADORA_SOBRE_ORUGAS'])
        elif student['Programa'] == 'GRUPO MAKRO REENTRENAMIENTO EN SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE TRACTOR SOBRE ORUGAS':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_REENTRENAMIENTO_EN_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_TRACTOR_SOBRE_ORUGAS'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE CAMION RIGIDO':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_CAMION_RIGIDO'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE CAMION VOLQUETE':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_CAMION_VOLQUETE'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE EXCAVADORAS Y PALAS.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_EXCAVADORAS_Y_PALAS'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE MINICARGADOR':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_MINICARGADOR'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE MOTONIVELADORA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_MOTONIVELADORA'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD. MANTENIMIENTO Y OPERACIÓN DE MOTOTRAILLA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_MOTOTRAILLA'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE PERFORADORAS':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_PERFORADORAS'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE RETROCARGADOR':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_RETROCARGADOR'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN SCOOP Y JUMBO':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_SCOOP_Y_JUMBO'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN CAMIÓN ARTICULADO.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_CAMION_ARTICULADO'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN CAMIONETAS 4X4':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_CAMIONETAS_4X4'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN CARGADOR FRONTAL':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_CARGADOR_FRONTAL'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN EXCAVADORA HIDRÁULICA.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_EXCAVADORA_HIDRAULICA'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN GRUAS':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_GRUAS'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN MAQUINARIA AGRICOLA.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_MAQUINARIA_AGRICOLA'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN MONTACARGAS':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_MONTACARGAS'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN TRACTOR AGRICOLA.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_TRACTOR_AGRICOLA'])
        elif student['Programa'] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN VIBROCOMPACTADOR SEGURIDAD VIAL':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_VIBROCOMPACTADOR_SEGURIDAD_VIAL'])
        elif student['Programa'] == 'SEGURIDAD, MANTENIMIENTO Y OPERACION DE CAMION ARTICULADO Y EXCAVADORA HIDRAULICA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_CAMION_ARTICULADO_Y_EXCAVADORA_HIDRAULICA'])
        elif student['Programa'] == 'SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE EQUIPO PESADO':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_EQUIPO_PESADO'])
        elif student['Programa'] == 'SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE EQUIPO PESADO GMK PANAMA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_EQUIPO_PESADO_GMK_PANAMA'])
        elif student['Programa'] == 'SOLDADURA SUBACUÁTICA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_SOLDADURA_SUBACUATICA'])
        elif student['Programa'] == 'TÉCNICO SUPERIOR ASISTENTE DE INGENIERÍA CIVIL':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_ASISTENTE_DE_INGENIERIA_CIVIL'])
        elif student['Programa'] == 'TÉCNICO SUPERIOR ASISTENTE DE ODONTOLOGÍA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_ASISTENTE_DE_ODONTOLOGIA'])
        elif student['Programa'] == 'TÉCNICO SUPERIOR DISEÑO DE OBRAS CIVILES':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_DISENO_DE_OBRAS_CIVILES'])
        elif student['Programa'] == 'TÉCNICO SUPERIOR ELECTRICIDAD CON ÉNFASIS EN CENTRALES HIDROELÉCTRICAS':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_ELECTRICIDAD_CON_ENFASIS_EN_CENTRALES_HIDROELECTRICAS'])
        elif student['Programa'] == 'TÉCNICO SUPERIOR EN MEDIO AMBIENTE Y MANEJO DE CUENCAS HIDROGRÁFICAS':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_EN_MEDIO_AMBIENTE_Y_MANEJO_DE_CUENCAS_HIDROGRAFICAS'])
        elif student['Programa'] == 'TÉCNICO SUPERIOR LOGÍSTICA Y COMERCIO INTERNACIONAL':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_LOGISTICA_Y_COMERCIO_INTERNACIONAL'])
        elif student['Programa'] == 'TÉCNICO SUPERIOR MECÁNICA DE EQUIPO PESADO':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_MECANICA_DE_EQUIPO_PESADO'])
        elif student['Programa'] == 'TÉCNICO SUPERIOR TOPOGRAFÍA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_TOPOGRAFIA'])
        elif student['Programa'] == 'TRIPULANTE DE CABINA DE VUELO COMERCIAL':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TRIPULANTE_DE_CABINA_DE_VUELO_COMERCIAL'])
        
        self.logger_print('Detalles de la oportunidad registrados')
        sleep(1)


    def save_oportunitie(self):
        '''
        Este metodo se encarga de guardar una oportunidad en el CRM
        '''
        self.click_element_by_ccs_sel(self.xpaths['BUTTON_GUARDAR'])
        self.logger_print('Oportunidad guardada')
        sleep(2)

    def send_text_by_ccs_sel(self, element, txt):
        WebDriverWait(self._driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, element)))
        e = self._driver.find_element(By.CSS_SELECTOR, element)
        e.send_keys(txt)

    def click_element_by_ccs_sel(self, element):
        WebDriverWait(self._driver, 4).until(EC.visibility_of_element_located((By.CSS_SELECTOR, element)))
        e = self._driver.find_element(By.CSS_SELECTOR, element)
        e.click()

    def click_element_by_xpath(self, element):
        WebDriverWait(self._driver, 10).until(EC.visibility_of_element_located((By.XPATH, element)))
        e = self._driver.find_element(By.XPATH, element)
        e.click()

    def select_option_by_ccs_sel(self, element, option):
        WebDriverWait(self._driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, element)))
        e = self._driver.find_element(By.CSS_SELECTOR, element)
        e.send_keys(option)
        e.send_keys(Keys.ENTER)

    def logger_print(self, msg):
        logger.info(msg)
        print(msg)
