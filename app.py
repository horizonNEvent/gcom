"""API do gerador de relatório GCom (FastAPI). Vercel detecta `app` neste arquivo.

Local: uvicorn app:app --reload  ->  http://127.0.0.1:8000
"""
import base64
import datetime
import os

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel

from services.gcom_client import GComClient
from template_builder import _parse_relatorio, gerar_template

INDEX_HTML = os.path.join(os.path.dirname(__file__), "static", "index.html")

app = FastAPI(title="Relatório GCom", docs_url=None, redoc_url=None)


class Pedido(BaseModel):
    # Credenciais vêm do formulário a cada pedido: nada fica guardado no servidor,
    # então quem abre a URL sem senha do GCom não consegue puxar dados.
    company: str
    user: str
    password: str
    venc_de: datetime.date
    venc_ate: datetime.date
    ent_de: datetime.date | None = None
    ent_ate: datetime.date | None = None


def _br(d: datetime.date | None) -> str:
    return d.strftime("%d/%m/%Y") if d else ""


@app.get("/")
def index():
    return FileResponse(INDEX_HTML)


@app.post("/api/gerar")
def gerar(p: Pedido):
    client = GComClient(p.user, p.password, p.company)
    ok, msg = client.login()
    if not ok:
        raise HTTPException(401, msg)

    try:
        raw = client.generate_report(_br(p.venc_de), _br(p.venc_ate), _br(p.ent_de), _br(p.ent_ate))
    except Exception as e:
        raise HTTPException(502, f"Erro ao gerar relatório no GCom: {e}")

    resp = {"relatorio": base64.b64encode(raw).decode(), "template": None, "linhas": None, "aviso": None}
    try:
        resp["linhas"] = len(_parse_relatorio(raw))
        resp["template"] = base64.b64encode(gerar_template(raw)).decode()
    except Exception as e:
        resp["aviso"] = f"Relatório baixado, mas falhou ao montar o template: {e}"
    return resp
