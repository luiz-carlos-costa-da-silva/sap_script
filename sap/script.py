from win32com import client

class Sap_Script(object):
    '''
        Classe responsável pela conexão e gerenciamento
        de sessão do SAP GUI.
        Cada sessão é uma janela SAP aberta. A primeira 
        sessão(janela) sempre começa na posição zero.
    '''
    def __init__(self) -> None:
        self.sap = None
        self.app = None
        self.con = None

    def isCDispatch(self, obj):
        '''
            Checa se objeto passado como parâmetro é do
            tipo CDispatch.
        '''
        if type(obj) == client.CDispatch:
            return True
        return False

    def open(self):
        '''
            Faz a conexão com SAP GUI.
        '''
        try:

            self.sap = client.GetObject('SAPGUI')
            if not self.isCDispatch(self.sap):
                return self.close()

            self.app = self.sap.GetScriptingEngine
            if not self.isCDispatch(self.app):
                return self.close()

            self.con = self.app.Children(0)
            if not self.isCDispatch(self.con):
                return self.close()

        except Exception as e:
            self.close()
            raise Exception('ERRO NA CONEXÃO COM SAP: [{e}]')

    def close(self):
        '''
            Fecha a conexão com SAP GUI.
        '''
        self.con = None
        self.app = None
        self.con = None

    def __len__(self):
        '''
            Obtém a quantidade de sessão aberto no SAP GUI.
        '''
        length = 0
        if self.isCDispatch(self.con):
            try:
                length = len(self.con.Sessions)
            except:
                pass
        return length

    def get_session_by_index(self, index):
        '''
            Se o index não for maior que a quantidade de sessão
            aberto, o método retorna a sessão de acordo com a sua
            posição(index) passada. A posição da sessão começa a partir do zero.
        '''
        session = None
        try: 
            if self.isCDispatch(self.con) and index < len(self):
                session = self.con.Children(index)
        except Exception as e:
            raise Exception(f'SESSÃO NÃO ENCONTRADA: {e}')
        return session
