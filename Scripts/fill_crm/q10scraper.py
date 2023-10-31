import sys
from pathlib import Path
BASE_DIR=Path(__file__).resolve().parent.parent.parent
sys.path.append(str(BASE_DIR))
from Scripts import common
logger = common.config()['logger']
from Scripts.local_google import google_methods
from time import sleep
import traceback

from Scripts.fill_crm.q10page import Q10Page

def main():
        student = google_methods.get_data_from_sheet()
        for i in range(len(student)):
            # Si el campo Cargado_Q10 es None, se carga el registro
            if student.iloc[i]['Cargado_Q10'] == None:
                q10p = Q10Page()
                try:
                    q10p.open()
                    q10p.login()
                    q10p.go_to_oportunities()
                    q10p.register_oportunitie(student.iloc[i])
                    q10p.register_oportunitie_detail(student.iloc[i])
                    q10p.save_oportunitie()
                    q10p._driver.quit()
                except Exception as e:
                    logger.error(e)
                    logger.error(traceback.format_exc())
                    print(traceback.format_exc())
                    sleep(50)
                    q10p._driver.quit()
        
if __name__ == '__main__':
    main()

