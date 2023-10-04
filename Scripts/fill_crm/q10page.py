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
        self._driver = webdriver.Chrome(options=self._chrome_options)
        self._driver.maximize_window()
        self._url = "https://site2.q10.com/login?ReturnUrl=%2F&aplentId=5f0cac06-a506-459a-a7b8-364b50574728"
        xpaths_path = common.config()['base_dir']/'Scripts/fill_crm/xpaths/xpaths.yml'
        self.xpaths = common.open_yaml(xpaths_path)
        self.student = google_methods.get_data_from_sheet()

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
        self.click_element_by_ccs_sel(self.xpaths['BUTTON_ADMINISTRATIVOS'])

        #sleep(38)

    def go_to_oportunities(self):
        '''
        Este metodo se encarga de ir a la seccion de oportunidades en el CRM
        '''
        WebDriverWait(self._driver, 3).until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#mensajeMora > div > div > div > div:nth-child(2) > div.col-lg-9.col-md-9.col-sm-9.col-xs-9')))
        WebDriverWait(self._driver, 40).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, '#mensajeMora > div > div > div > div:nth-child(2) > div.col-lg-9.col-md-9.col-sm-9.col-xs-9')))
        self.click_element_by_ccs_sel(self.xpaths['MENU_INSTITUCIONAL'])
        WebDriverWait(self._driver, 3).until(EC.visibility_of_element_located((By.CSS_SELECTOR, self.xpaths['MENU_COMERCIAL_CRM'])))
        menu_comercial_crm = self._driver.find_element(By.CSS_SELECTOR, self.xpaths['MENU_COMERCIAL_CRM'])
        actions = ActionChains(self._driver)
        actions.move_to_element(menu_comercial_crm).perform()
        self.click_element_by_ccs_sel(self.xpaths['MENU_OPORTUNITIES'])

        sleep(2)

    def register_oportunitie(self):
        '''
        Este metodo se encarga de registrar una oportunidad en el CRM
        '''
        self.click_element_by_ccs_sel(self.xpaths['BUTTON_REGISTRAR_OPORTUNIDAD'])
        name = str(self.student['Nombre'].iloc[0].iloc[0])+' '+str(self.student['Apellido'].iloc[0].iloc[0])
        self.send_text_by_ccs_sel(self.xpaths['TXTBOX_OPOTUNIDAD_NOMBRE'], name)
        self.send_text_by_ccs_sel(self.xpaths['TXTBOX_IDENTIFICACION'], self.student['Cedula'].iloc[0])
        self.send_text_by_ccs_sel(self.xpaths['TXTBOX_EMAIL'], self.student['Correo'].iloc[0])
        self.send_text_by_ccs_sel(self.xpaths['TXTBOX_CELULAR'], self.student['Celular'].iloc[0])

    def register_oportunitie_detail(self):
        '''
        Este metodo se encarga de registrar los detalles de una oportunidad en el CRM
        '''
        self.click_element_by_ccs_sel(self.xpaths['BUTTON_NEXT'])
        self.send_text_by_ccs_sel(self.xpaths['TXTBOX_DIRECCION'], self.student['Direccion'].iloc[0])
        self.click_element_by_ccs_sel(self.xpaths['SELECT_COMO_SE_ENTERO'])

        #Se selecciona como se entero el estudiante
        if self.student['Como_Se_Entero'].iloc[0].iloc[0] == 'Facebook Empresarial':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_FB_EMPRESARIAL'])
        elif self.student['Como_Se_Entero'].iloc[0] == 'Instagram':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_INSTAGRAM'])
        elif self.student['Como_Se_Entero'].iloc[0] == 'Pagina Web':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_PAGINA_WEB'])

        # Se selecciona el medio de contacto con el estudiante
        self.click_element_by_ccs_sel(self.xpaths['SELECT_MEDIO_DE_CONTACTO'])
        if self.student['Medio_De_Contacto'].iloc[0].iloc[0] == 'WhatsApp':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_WHATSAPP'])
        elif self.student['Medio_De_Contacto'].iloc[0].iloc[0] == 'Correo Electrónico':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CORREO_ELECTRONICO'])
        elif self.student['Medio_De_Contacto'].iloc[0].iloc[0] == 'Visita':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_VISITA'])
        elif self.student['Medio_De_Contacto'].iloc[0].iloc[0] == 'Llamada':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_LLAMADA'])

        # Se selecciona la procedencia, es decir hacia donde va el estudiante
        self.click_element_by_ccs_sel(self.xpaths['SELECT_PROCEDENCIA'])
        if self.student['Procedencia'].iloc[0].iloc[0] == 'INSTITUTO SUPERIOR DE INGENIERIA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_ISI'])
        elif self.student['Procedencia'].iloc[0].iloc[0] == 'GRUPO MAKRO PANAMA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GMK_PANAMA'])
        elif self.student['Procedencia'].iloc[0].iloc[0] == 'GRUPO MAKRO COLOMBIA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GMK_COLOMBIA'])
        elif self.student['Procedencia'].iloc[0].iloc[0] == 'GRUPO MAKRO CONVENIO':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GMK_CONVENIO'])


        # Se selecciona el estado del negocio
        self.click_element_by_ccs_sel(self.xpaths['SELECT_ESTADO_DEL_NEGOCIO'])
        if self.student['Estado_Del_Negocio'].iloc[0].iloc[0] == 'Presentación':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_PRESENTACION'])
        elif self.student['Estado_Del_Negocio'].iloc[0].iloc[0] == 'En Negociación':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_EN_NEGOCIACION'])
        elif self.student['Estado_Del_Negocio'].iloc[0].iloc[0] == 'Cierre':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CIERRE'])

        # Se selecciona el municipio del estudiante
        self.select_option_by_ccs_sel(self.xpaths['SELECT_MUNICIPIO'], self.student['Municipio'].iloc[0].iloc[0])

        '''Se selecciona la carrera del estudiante entre las siguientes opciones:
            ASISTENTE DE INGENIERIA CIVIL Y DISEÑO DE OBRAS CIVILES
            CERTIFICACIONES ESPECIALIZADAS GRUPO MAKRO
            CURSO DE BUCEO COMERCIAL
            CURSO DE FRANCES JOBS - CANADA
            CURSO DE INGLES ISI
            CURSO DE INGLES JOBS - CANADA
            CURSOS ISI
            CURSOS NIVELATORIOS ISI
            CURSOS SOBRE BUCEO ISI
            DIPLOMADO EN MAQUINARIA PESADA CON ENFASIS EN MANTENIMIENTO Y PRODUCTIVIDAD.
            DIPLOMADO EN SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EQUIPO PESADO ZEMER 
            FLYCAMDRONE - CURSO BASICO DE DRONES.
            FLYCAMDRONE - CURSO PROFESIONAL DE DRONES
            GRUPO MAKRO REENTRENAMIENTO EN SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE EXCAVADORA HIDRÁULICA.
            GRUPO MAKRO REENTRENAMIENTO EN SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE EXCAVADORA SOBRE ORUGAS 
            GRUPO MAKRO REENTRENAMIENTO EN SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE TRACTOR SOBRE ORUGAS 
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE CAMION RIGIDO 
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE CAMION VOLQUETE 
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE EXCAVADORAS Y PALAS.
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE MINICARGADOR 
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE MOTONIVELADORA 
            GRUPO MAKRO SEGURIDAD. MANTENIMIENTO Y OPERACIÓN DE MOTOTRAILLA
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE PERFORADORAS 
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE RETROCARGADOR 
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN SCOOP Y JUMBO 
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN CAMIÓN ARTICULADO.
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN CAMIONETAS 4X4 
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN CARGADOR FRONTAL 
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN EXCAVADORA HIDRÁULICA.
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN GRUAS 
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN MAQUINARIA AGRICOLA.
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN MONTACARGAS 
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN TRACTOR AGRICOLA.
            GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN VIBROCOMPACTADOR SEGURIDAD VIAL
            SEGURIDAD, MANTENIMIENTO Y OPERACION DE CAMION ARTICULADO Y EXCAVADORA HIDRAULICA 
            SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE EQUIPO PESADO 
            SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE EQUIPO PESADO GMK PANAMA 
            SOLDADURA SUBACUÁTICA 
            TÉCNICO SUPERIOR ASISTENTE DE INGENIERÍA CIVIL 
            TÉCNICO SUPERIOR ASISTENTE DE ODONTOLOGÍA 
            TÉCNICO SUPERIOR DISEÑO DE OBRAS CIVILES 
            TÉCNICO SUPERIOR ELECTRICIDAD CON ÉNFASIS EN CENTRALES HIDROELÉCTRICAS 
            TÉCNICO SUPERIOR EN MEDIO AMBIENTE Y MANEJO DE CUENCAS HIDROGRÁFICAS
            TÉCNICO SUPERIOR LOGÍSTICA Y COMERCIO INTERNACIONAL
            TÉCNICO SUPERIOR MECÁNICA DE EQUIPO PESADO
            TÉCNICO SUPERIOR TOPOGRAFÍA
            TRIPULANTE DE CABINA DE VUELO COMERCIAL
        
        self.click_element_by_ccs_sel(self.xpaths['SELECT_PROGRAMA'])
        if self.student['Programa'].iloc[0].iloc[0] == 'ASISTENTE DE INGENIERIA CIVIL Y DISEÑO DE OBRAS CIVILES':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_ASISTENTE_DE_INGENIERIA_CIVIL_Y_DISENO_DE_OBRAS_CIVILES'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'CERTIFICACIONES ESPECIALIZADAS GRUPO MAKRO':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CERTIFICACIONES_ESPECIALIZADAS_GRUPO_MAKRO'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'CURSO DE BUCEO COMERCIAL':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CURSO_DE_BUCEO_COMERCIAL'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'CURSO DE FRANCES JOBS - CANADA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CURSO_DE_FRANCES_JOBS_CANADA'])   
        elif self.student['Programa'].iloc[0].iloc[0] == 'CURSO DE INGLES ISI':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CURSO_DE_INGLES_ISI'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'CURSO DE INGLES JOBS - CANADA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CURSO_DE_INGLES_JOBS_CANADA'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'CURSOS ISI':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CURSOS_ISI'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'CURSOS NIVELATORIOS ISI':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CURSOS_NIVELATORIOS_ISI'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'CURSOS SOBRE BUCEO ISI':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_CURSOS_SOBRE_BUCEO_ISI'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'DIPLOMADO EN MAQUINARIA PESADA CON ENFASIS EN MANTENIMIENTO Y PRODUCTIVIDAD.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_DIPLOMADO_EN_MAQUINARIA_PESADA_CON_ENFASIS_EN_MANTENIMIENTO_Y_PRODUCTIVIDAD'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'DIPLOMADO EN SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EQUIPO PESADO ZEMER':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_DIPLOMADO_EN_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EQUIPO_PESADO_ZEMER'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'FLYCAMDRONE - CURSO BASICO DE DRONES.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_FLYCAMDRONE_CURSO_BASICO_DE_DRONES'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'FLYCAMDRONE - CURSO PROFESIONAL DE DRONES':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_FLYCAMDRONE_CURSO_PROFESIONAL_DE_DRONES'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO REENTRENAMIENTO EN SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE EXCAVADORA HIDRÁULICA.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_REENTRENAMIENTO_EN_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_EXCAVADORA_HIDRAULICA'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO REENTRENAMIENTO EN SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE EXCAVADORA SOBRE ORUGAS':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_REENTRENAMIENTO_EN_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_EXCAVADORA_SOBRE_ORUGAS'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO REENTRENAMIENTO EN SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE TRACTOR SOBRE ORUGAS':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_REENTRENAMIENTO_EN_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_TRACTOR_SOBRE_ORUGAS'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE CAMION RIGIDO':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_CAMION_RIGIDO'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE CAMION VOLQUETE':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_CAMION_VOLQUETE'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE EXCAVADORAS Y PALAS.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_EXCAVADORAS_Y_PALAS'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE MINICARGADOR':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_MINICARGADOR'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE MOTONIVELADORA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_MOTONIVELADORA'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD. MANTENIMIENTO Y OPERACIÓN DE MOTOTRAILLA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_MOTOTRAILLA'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE PERFORADORAS':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_PERFORADORAS'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE RETROCARGADOR':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_RETROCARGADOR'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN SCOOP Y JUMBO':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_SCOOP_Y_JUMBO'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN CAMIÓN ARTICULADO.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_CAMION_ARTICULADO'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN CAMIONETAS 4X4':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_CAMIONETAS_4X4'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN CARGADOR FRONTAL':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_CARGADOR_FRONTAL'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN EXCAVADORA HIDRÁULICA.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_EXCAVADORA_HIDRAULICA'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN GRUAS':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_GRUAS'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN MAQUINARIA AGRICOLA.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_MAQUINARIA_AGRICOLA'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN MONTACARGAS':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_MONTACARGAS'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN TRACTOR AGRICOLA.':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_TRACTOR_AGRICOLA'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'GRUPO MAKRO SEGURIDAD, MANTENIMIENTO Y OPERACIÓN EN VIBROCOMPACTADOR SEGURIDAD VIAL':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_GRUPO_MAKRO_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_EN_VIBROCOMPACTADOR_SEGURIDAD_VIAL'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'SEGURIDAD, MANTENIMIENTO Y OPERACION DE CAMION ARTICULADO Y EXCAVADORA HIDRAULICA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_CAMION_ARTICULADO_Y_EXCAVADORA_HIDRAULICA'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE EQUIPO PESADO':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_EQUIPO_PESADO'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'SEGURIDAD, MANTENIMIENTO Y OPERACIÓN DE EQUIPO PESADO GMK PANAMA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_SEGURIDAD_MANTENIMIENTO_Y_OPERACION_DE_EQUIPO_PESADO_GMK_PANAMA'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'SOLDADURA SUBACUÁTICA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_SOLDADURA_SUBACUATICA'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'TÉCNICO SUPERIOR ASISTENTE DE INGENIERÍA CIVIL':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_ASISTENTE_DE_INGENIERIA_CIVIL'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'TÉCNICO SUPERIOR ASISTENTE DE ODONTOLOGÍA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_ASISTENTE_DE_ODONTOLOGIA'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'TÉCNICO SUPERIOR DISEÑO DE OBRAS CIVILES':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_DISENO_DE_OBRAS_CIVILES'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'TÉCNICO SUPERIOR ELECTRICIDAD CON ÉNFASIS EN CENTRALES HIDROELÉCTRICAS':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_ELECTRICIDAD_CON_ENFASIS_EN_CENTRALES_HIDROELECTRICAS'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'TÉCNICO SUPERIOR EN MEDIO AMBIENTE Y MANEJO DE CUENCAS HIDROGRÁFICAS':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_EN_MEDIO_AMBIENTE_Y_MANEJO_DE_CUENCAS_HIDROGRAFICAS'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'TÉCNICO SUPERIOR LOGÍSTICA Y COMERCIO INTERNACIONAL':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_LOGISTICA_Y_COMERCIO_INTERNACIONAL'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'TÉCNICO SUPERIOR MECÁNICA DE EQUIPO PESADO':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_MECANICA_DE_EQUIPO_PESADO'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'TÉCNICO SUPERIOR TOPOGRAFÍA':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TECNICO_SUPERIOR_TOPOGRAFIA'])
        elif self.student['Programa'].iloc[0].iloc[0] == 'TRIPULANTE DE CABINA DE VUELO COMERCIAL':
            self.click_element_by_ccs_sel(self.xpaths['OPTION_TRIPULANTE_DE_CABINA_DE_VUELO_COMERCIAL'])
'''
        '''
        El listado de opciones es:
        OPTION_ASISTENTE_DE_INGENIERIA_CIVIL_Y_DISENO_DE_OBRAS_CIVILES
        OPTION_CERTIFICACIONES_ESPECIALIZADAS_GRUPO_MAKRO
        OPTION_CURSO_DE_BUCEO_COMERCIAL
        OPTION_CURSO_DE_FRANCES_JOBS_CANADA
        OPTION_CURSO_DE_INGLES_ISI
        OPTION_CURSO_DE_INGLES_JOBS_CANADA
        OPTION_CURSOS_ISI
        OPTION_CURSOS_NIVELATORIOS_ISI
        OPTION_CURSOS_SOBRE_BUCEO_ISI
        
        '''

        sleep(15)

    def send_text_by_ccs_sel(self, element, txt):
        WebDriverWait(self._driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, element)))
        e = self._driver.find_element(By.CSS_SELECTOR, element)
        e.send_keys(txt)

    def click_element_by_ccs_sel(self, element):
        WebDriverWait(self._driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, element)))
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