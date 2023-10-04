import sys
from pathlib import Path
BASE_DIR=Path(__file__).resolve().parent.parent.parent
sys.path.append(str(BASE_DIR))
from Scripts import common
logger = common.config()['logger']
from time import sleep

from Scripts.fill_crm.q10page import Q10Page

def main():
    q10p = Q10Page()
    q10p.open()
    q10p.login()
    q10p.go_to_oportunities()
    q10p.register_oportunitie()
    q10p.register_oportunitie_detail()
    

if __name__ == '__main__':
    main()

