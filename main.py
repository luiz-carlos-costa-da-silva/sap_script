from sap.script import Sap_Script

sap_script = Sap_Script()

sap_script.open()
print(sap_script.get_session_by_index(0))
sap_script.close()
