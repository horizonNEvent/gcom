import requests
from bs4 import BeautifulSoup
import re
import urllib3
import urllib.parse
import datetime

# Clean up warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class GComClient:
    def __init__(self, username, password, company):
        self.username = username
        self.password = password
        self.company = company
        self.session = requests.Session()
        self.session.verify = False
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })
        self.base_url = "https://www2.gcom.com.br"

    def login(self):
        try:
            # 1. Access Login Page
            login_url = f"{self.base_url}/Cloud/Home/Login.aspx"
            resp_get = self.session.get(login_url)
            resp_get.raise_for_status()
            
            soup = BeautifulSoup(resp_get.content, 'html.parser')
            payload = {}
            for inp in soup.find_all('input'):
                if inp.get('name') and inp.get('type') not in ['submit', 'button', 'image']:
                    payload[inp.get('name')] = inp.get('value', '')
                    
            # Credentials
            payload['ctl00$ContentBody$txtCompanyName'] = self.company
            payload['ctl00$ContentBody$txtUserName'] = self.username
            payload['ctl00$ContentBody$txtPassword'] = self.password
            payload['ctl00$ContentBody$btnLogin'] = 'Login'
            
            # 2. Login POST
            resp_post = self.session.post(login_url, data=payload)
            resp_post.raise_for_status()
            
            if "txtPassword" in resp_post.text:
                return False, "Credenciais inválidas ou erro na página de login."
                
            # 3. Extract Session
            sid_match = re.search(r'SessionIdGlobal=([^"\'&\s]+)', resp_post.text)
            emp_match = re.search(r'sNm_Empresa=([^"\'&\s]+)', resp_post.text)
            
            if not sid_match:
                return False, "Sessão não encontrada após login."
                
            self.sid = sid_match.group(1)
            self.emp = emp_match.group(1) if emp_match else self.company.upper()
            
            # 4. Activate Session
            self._activate_session()
            return True, "Login realizado com sucesso."
            
        except Exception as e:
            return False, f"Erro no Login: {str(e)}"

    def _activate_session(self):
        # Loader & Activate Session links
        base_loader = f"{self.base_url}/gcom/Carrega_Session.asp"
        
        loader_url = f"{base_loader}?CMD=LIMPLOGONSITE&SessionIdGlobal={self.sid}&sNm_Empresa={self.emp}&sIc_Alt_Pss=&sNm_Pgm=&sNmSelectModulo=&sNmSelectSubtitulo=&sCor=&sTipoProg=&sIc_B2B=&nId_Etb_Cmp=&nId_Ped_Cmp=&nId_Mrc=&sIc_Novo_Erp=&CompanyCode="
        self.session.get(loader_url)
        
        final_url = f"{base_loader}?CMD=AtivaAcesso&SessionIdGlobal={self.sid}&sNm_Empresa={self.emp}&sIc_Alt_Pss=&sNm_Pgm=&sNmSelectModulo=&sNmSelectSubtitulo=&sCor=&sTipoProg=&sIc_B2B=&nId_Etb_Cmp=&nId_Ped_Cmp=&nId_Mrc=&sIc_Novo_Erp=&CompanyCode="
        self.session.post(final_url)
        
        # Dashboard sequence to set cookies
        self.session.get(f"{self.base_url}/Gcom/MenuPrincipal3.asp")
        self.session.get(f"{self.base_url}/Gcom/CONTASPAGAR.ASP")
        self.session.get(f"{self.base_url}/Gcom/CONTASPAGAR/DOCEMABERT.ASP")

    def get_filial_id(self, filial_name_search="CASA TUA - CUCCINA"):
        url = f"{self.base_url}/Gcom/include/cboEmpFilUsu.asp?CMD=CarregaComboUnidade&nId_Emp_Gcom=45702&sIc_Atv_Ina=A&sTp_Nm_Fil=RDU&sIc_Pesq=S"
        resp = self.session.get(url)
        soup = BeautifulSoup(resp.text, 'html.parser')
        
        filial_id = "-1"
        filial_name = "TODAS"
        
        for opt in soup.find_all('option'):
            if filial_name_search.upper() in opt.text.upper():
                filial_id = opt['value']
                filial_name = opt.text.strip()
                break
        
        # Fallback to second option if not found
        if filial_id == "-1":
             opts = soup.find_all('option')
             if len(opts) > 1:
                 filial_id = opts[1]['value']
                 filial_name = opts[1].text.strip()
        
        return filial_id, filial_name

    def generate_report(self, dt_vencto_de, dt_vencto_ate, dt_entrada_de, dt_entrada_ate):
        
        filial_id, filial_name = self.get_filial_id()
        
        excel_url = f"{self.base_url}/Gcom/CONTASPAGAR/DocEmAbertExcel.asp"
        
        # Default value for start date if empty as per system logic
        if not dt_vencto_de:
            dt_vencto_de = "01/01/0001"
            
        params = {
            "nFilialT": filial_id,
            "dDt_De": dt_vencto_de,
            "dDt_Ate": dt_vencto_ate,
            "nTpDoctoT": "-1",
            "nId_etb": "",
            "nTpDespesaT": "-1",
            "ncboTpCobrancaT": "-1",
            "pcboOrdena": "V", 
            "nHis": "-1",
            "nNr_Doc": "",
            "sFornecedor": "",
            "sHis": "TODAS",
            "sFilialT": filial_name,
            "sEmpContabil": "TODAS", 
            "sTpDespesaT": "TODAS",
            "sTpCobrancaT": "TODAS",
            "sTpDoctoT": "TODAS",
            "sQuebra": "N",
            "sQuebra_Dsp": "N",
            "PR": "S,N", 
            "dDt_Ent_Doc_De": dt_entrada_de,
            "dDt_Ent_Doc_Ate": dt_entrada_ate,
            "sIc_Ent_Merc": "SN", 
            "sCd_Cen_Cst": "-1",
            "sNm_Cen_Cst": "", 
            "sIc_Res_Dsp": "N",
            "nVl_De": "0.00",
            "nVl_Ate": "99999999999.99",
            "sIc_Tp_Pesq_Docto": "IGUAL", 
            "sDc_Obs": "",
            "sChkSem_Int_Ban": "", 
            "nId_Arq_Reg": "-1",
            "sDc_Arq_Reg": "",
            "dDt_Inc_De": "",
            "dDt_Inc_Ate": "",
            "dDt_Comp_De": "",
            "dDt_Comp_Ate": "",
            "sIc_Doc_Com_Rat_CC": "N",
            "nId_Exc": "-1",
            "sCd_Cen_Cst_New": "-1",
            "sIc_Nao_Formata_CNPJ_CPF": "N" 
        }
        
        query_string = urllib.parse.urlencode(params)
        full_excel_url = f"{excel_url}?{query_string}"
        
        resp = self.session.get(full_excel_url)
        resp.raise_for_status()
        
        return resp.content
