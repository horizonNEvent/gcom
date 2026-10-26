"""Versão local (sem Streamlit) do gerador de relatório GCom.

Uso:
    python run_local.py --venc-ate 31/12/2026 [--venc-de 01/10/2026] [--ent-de ..] [--ent-ate ..]

Credenciais: variáveis GCOM_COMPANY / GCOM_USER / GCOM_PASS; o que faltar é perguntado
no terminal (senha via getpass, nunca por argumento de linha de comando).
Saída em debug/: report_<stamp>.xls (bruto do GCOM) e template_<stamp>.xlsx.
"""
import argparse
import datetime
import getpass
import os
import sys

from services.gcom_client import GComClient
from template_builder import _parse_relatorio, gerar_template

DEBUG_DIR = os.path.join(os.path.dirname(__file__), "debug")


def main():
    ap = argparse.ArgumentParser(description="Gera relatório GCom + template localmente")
    ap.add_argument("--venc-de", default="")
    ap.add_argument("--venc-ate", required=True)
    ap.add_argument("--ent-de", default="")
    ap.add_argument("--ent-ate", default="")
    args = ap.parse_args()

    company = os.environ.get("GCOM_COMPANY") or input("Empresa: ")
    user = os.environ.get("GCOM_USER") or input("Usuário: ")
    password = os.environ.get("GCOM_PASS") or getpass.getpass("Senha: ")

    client = GComClient(user, password, company)
    ok, msg = client.login()
    print(msg)
    if not ok:
        sys.exit(1)

    raw = client.generate_report(args.venc_de, args.venc_ate, args.ent_de, args.ent_ate)
    os.makedirs(DEBUG_DIR, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    raw_path = os.path.join(DEBUG_DIR, f"report_{stamp}.xls")
    with open(raw_path, "wb") as f:
        f.write(raw)
    print(f"Bruto:    {raw_path} ({len(raw)} bytes)")

    # Diagnóstico: como as colunas sensíveis chegam do GCOM, antes de qualquer conversão
    df = _parse_relatorio(raw)
    print(f"Linhas: {len(df)} | Colunas: {list(df.columns)}")
    cols = [c for c in df.columns if any(k in str(c).upper() for k in ("DOCTO", "VALOR", "CNPJ", "VENCTO"))]
    if cols:
        print(df[cols].head(10).to_string())
        print("dtypes:", df[cols].dtypes.to_dict())

    tpl_path = os.path.join(DEBUG_DIR, f"template_{stamp}.xlsx")
    with open(tpl_path, "wb") as f:
        f.write(gerar_template(raw))
    print(f"Template: {tpl_path}")


if __name__ == "__main__":
    main()
