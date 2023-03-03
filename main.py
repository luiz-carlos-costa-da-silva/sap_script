'''
Autor: Luiz Carlos Costa da Silva
Data da criação: 01/02/2023
Objetivo do projeto:
Gerar relatório logístico diário e enviar a diretoria e gerencia de maneira automatizada.
'''

from sap.script import Sap_Script

sap_script = Sap_Script()

sap_script.open()
print(sap_script.get_session_by_index(0))
sap_script.close()
