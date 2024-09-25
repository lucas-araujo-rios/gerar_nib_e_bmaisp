import os
import sys
from dotenv import load_dotenv
from office365_api.download_files import get_file

# carregar .env e tudo mais
load_dotenv()
ROOT = os.getenv('ROOT')
PATH_OFFICE = os.path.abspath(os.path.join(ROOT, 'office365_api'))

# Adiciona o diretório correto ao sys.path
sys.path.append(PATH_OFFICE)

# puxar planilhas do sharepoint
def puxar_planilhas():
    projetos = get_file('portfolio.xlsx', 'DWPII/srinfo', ROOT)
    projetos_empresas = get_file('projetos_empresas.xlsx', 'DWPII/srinfo', ROOT)
    empresas = get_file('info_empresas.xlsx', 'DWPII/srinfo', ROOT)
    print('Download concluído')

puxar_planilhas()


