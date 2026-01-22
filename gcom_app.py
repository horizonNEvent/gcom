import streamlit as st
import requests
from bs4 import BeautifulSoup
import re
import urllib3
import urllib.parse
import datetime

# Clean up warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- GCom Logic ---

def run_gcom_report(dt_vencto_de, dt_vencto_ate, dt_entrada_de, dt_entrada_ate):
    """
    Executes the full GCom login flow and generates the report with the given parameters.
    Returns: Bytes of the report file (or None if failed), and a status message.
    """
    session = requests.Session()
    session.verify = False 
    
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    })
    
    try:
        # 1. Access Login Page
        login_url = "https://www2.gcom.com.br/Cloud/Home/Login.aspx"
        resp_get = session.get(login_url)
        resp_get.raise_for_status()
        
        soup = BeautifulSoup(resp_get.content, 'html.parser')
        payload = {}
        for inp in soup.find_all('input'):
            if inp.get('name') and inp.get('type') not in ['submit', 'button', 'image']:
                payload[inp.get('name')] = inp.get('value', '')
                
        # Credentials (Hardcoded for this demo, in prod use env vars or secure storage)
        payload['ctl00$ContentBody$txtCompanyName'] = 'Casatua'
        payload['ctl00$ContentBody$txtUserName'] = 'Flopes'
        payload['ctl00$ContentBody$txtPassword'] = 'F@ril20130428'
        payload['ctl00$ContentBody$btnLogin'] = 'Login'
        
        # 2. Login POST
        resp_post = session.post(login_url, data=payload)
        resp_post.raise_for_status()
        
        if "txtPassword" in resp_post.text:
            return None, "Falha no Login: Credenciais inválidas ou erro na página."
            
        # 3. Extract Session
        sid_match = re.search(r'SessionIdGlobal=([^"\'&\s]+)', resp_post.text)
        emp_match = re.search(r'sNm_Empresa=([^"\'&\s]+)', resp_post.text)
        
        if not sid_match:
            return None, "Falha no Login: Sessão não encontrada."
            
        sid = sid_match.group(1)
        emp = emp_match.group(1) if emp_match else "CASATUA"
        
        # 4. Loader & Activate Session
        next_url = f"https://www2.gcom.com.br/gcom/Carrega_Session.asp?CMD=LIMPLOGONSITE&SessionIdGlobal={sid}&sNm_Empresa={emp}&sIc_Alt_Pss=&sNm_Pgm=&sNmSelectModulo=&sNmSelectSubtitulo=&sCor=&sTipoProg=&sIc_B2B=&nId_Etb_Cmp=&nId_Ped_Cmp=&nId_Mrc=&sIc_Novo_Erp=&CompanyCode="
        session.get(next_url)
        
        final_url = f"https://www2.gcom.com.br/gcom/Carrega_Session.asp?CMD=AtivaAcesso&SessionIdGlobal={sid}&sNm_Empresa={emp}&sIc_Alt_Pss=&sNm_Pgm=&sNmSelectModulo=&sNmSelectSubtitulo=&sCor=&sTipoProg=&sIc_B2B=&nId_Etb_Cmp=&nId_Ped_Cmp=&nId_Mrc=&sIc_Novo_Erp=&CompanyCode="
        session.post(final_url)
        
        # 5. Dashboard -> Menu -> Contas Pagar -> DocEmAbert (Chain required for cookies/session state sometimes)
        session.get("https://www2.gcom.com.br/Gcom/MenuPrincipal3.asp")
        session.get("https://www2.gcom.com.br/Gcom/CONTASPAGAR.ASP")
        resp_doc = session.get("https://www2.gcom.com.br/Gcom/CONTASPAGAR/DOCEMABERT.ASP")
        
        # 6. Fetch Filial ID
        filial_url = "https://www2.gcom.com.br/Gcom/include/cboEmpFilUsu.asp?CMD=CarregaComboUnidade&nId_Emp_Gcom=45702&sIc_Atv_Ina=A&sTp_Nm_Fil=RDU&sIc_Pesq=S"
        resp_fil = session.get(filial_url)
        
        soup_fil = BeautifulSoup(resp_fil.text, 'html.parser')
        filial_id = "-1"
        filial_name = "TODAS"
        
        # Find 'CASA TUA - CUCCINA'
        for opt in soup_fil.find_all('option'):
            if "CASA TUA - CUCCINA" in opt.text.upper():
                filial_id = opt['value']
                filial_name = opt.text.strip()
                break
        
        # Fallback
        if filial_id == "-1":
             opts = soup_fil.find_all('option')
             if len(opts) > 1:
                 filial_id = opts[1]['value']
                 filial_name = opts[1].text.strip()
                 
        # 7. Generate Excel Report
        excel_url = "https://www2.gcom.com.br/Gcom/CONTASPAGAR/DocEmAbertExcel.asp"
        
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
            "nVl_Ate": "99999999999.99", # Clean number format
            "sIc_Tp_Pesq_Docto": "IGUAL", 
            "sDc_Obs": "",
            "sChkSem_Int_Ban": "", # Both Checked
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
        
        resp_excel = session.get(full_excel_url)
        resp_excel.raise_for_status()
        
        return resp_excel.content, "Sucesso"
        
    except Exception as e:
        return None, f"Erro Fatal: {str(e)}"


# --- Streamlit UI ---

st.set_page_config(page_title="Relatório GCom", page_icon="📊")

st.title("📊 Gerador de Relatório GCom")
st.markdown("Preencha os filtros de data abaixo e clique em baixar.")

# Defaults
default_date_start = datetime.date(2026, 1, 1)
default_date_end = datetime.date(2026, 1, 22)

with st.form("filter_form"):
    st.subheader("Filtros de Data")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**📅 Vencimento**")
        v_de = st.date_input("De", value=default_date_start, key="v_de")
        v_ate = st.date_input("Até", value=default_date_end, key="v_ate")
        
    with col2:
        st.markdown("**📥 Entrada**")
        e_de = st.date_input("De", value=default_date_start, key="e_de")
        e_ate = st.date_input("Até", value=default_date_end, key="e_ate")
    
    submit_btn = st.form_submit_button("🔍 Gerar Relatório", type="primary")

if submit_btn:
    # Format dates to dd/mm/yyyy
    f_v_de = v_de.strftime("%d/%m/%Y")
    f_v_ate = v_ate.strftime("%d/%m/%Y")
    f_e_de = e_de.strftime("%d/%m/%Y")
    f_e_ate = e_ate.strftime("%d/%m/%Y")
    
    with st.spinner("Conectando ao GCom e gerando relatório..."):
        file_bytes, status = run_gcom_report(f_v_de, f_v_ate, f_e_de, f_e_ate)
        
    if file_bytes:
        st.success("Relatório gerado com sucesso!")
        st.download_button(
            label="⬇️ Baixar Relatório (.xls)",
            data=file_bytes,
            file_name="relatorio_contas_pagar.xls",
            mime="application/vnd.ms-excel"
        )
    else:
        st.error(status)
