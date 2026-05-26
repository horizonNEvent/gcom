import streamlit as st
import datetime
import os
from services.gcom_client import GComClient
from template_builder import gerar_template

DEBUG_DIR = os.path.join(os.path.dirname(__file__), "debug")

st.set_page_config(page_title="Relatório GCom", page_icon="📊")

st.title("📊 Gerador de Relatório GCom")

# --- Configuration Section ---
# Try to load secrets, otherwise show inputs (for security in public repos)
if "gcom_user" in st.secrets:
    username = st.secrets["gcom_user"]
    password = st.secrets["gcom_pass"]
    company = st.secrets["gcom_company"]
    has_creds = True
else:
    st.warning("⚠️ Credenciais não configuradas em `.streamlit/secrets.toml`.")
    with st.expander("Configurar Credenciais Manualmente"):
        company = st.text_input("Empresa", value="Casatua")
        username = st.text_input("Usuário")
        password = st.text_input("Senha", type="password")
        has_creds = username and password

st.divider()

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
    if not has_creds:
        st.error("Por favor, configure as credenciais primeiro.")
    else:
        # Format dates to dd/mm/yyyy
        f_v_de = v_de.strftime("%d/%m/%Y")
        f_v_ate = v_ate.strftime("%d/%m/%Y")
        f_e_de = e_de.strftime("%d/%m/%Y")
        f_e_ate = e_ate.strftime("%d/%m/%Y")
        
        client = GComClient(username, password, company)
        
        with st.spinner("Conectando ao GCom..."):
            success, msg = client.login()
            
        if not success:
            st.error(msg)
        else:
            with st.spinner("Gerando relatório Excel..."):
                try:
                    file_bytes = client.generate_report(f_v_de, f_v_ate, f_e_de, f_e_ate)
                    st.success("Relatório gerado com sucesso!")

                    try:
                        os.makedirs(DEBUG_DIR, exist_ok=True)
                        with open(os.path.join(DEBUG_DIR, "last_report.xls"), "wb") as _f:
                            _f.write(file_bytes)
                        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                        with open(os.path.join(DEBUG_DIR, f"report_{stamp}.xls"), "wb") as _f:
                            _f.write(file_bytes)
                    except Exception as _e:
                        st.info(f"(debug) não consegui salvar cópia local: {_e}")

                    with st.spinner("Montando template de importação..."):
                        try:
                            template_bytes = gerar_template(file_bytes)
                            template_ok = True
                        except Exception as e:
                            template_bytes = None
                            template_ok = False
                            st.warning(f"Relatório baixado, mas falhou ao montar o template: {e}")

                    col_dl1, col_dl2 = st.columns(2)
                    with col_dl1:
                        st.download_button(
                            label="⬇️ Baixar Relatório Original",
                            data=file_bytes,
                            file_name="relatorio_contas_pagar.xls",
                            mime="application/vnd.ms-excel",
                            use_container_width=True,
                        )
                    with col_dl2:
                        if template_ok:
                            st.download_button(
                                label="📋 Baixar Template Preenchido",
                                data=template_bytes,
                                file_name="titulos_a_pagar_GCOM.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                            )
                        else:
                            st.button(
                                "📋 Template indisponível",
                                disabled=True,
                                use_container_width=True,
                            )
                except Exception as e:
                    st.error(f"Erro ao gerar relatório: {str(e)}")
