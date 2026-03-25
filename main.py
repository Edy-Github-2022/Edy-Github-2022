from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import os, uuid, zipfile, re, time, shutil
from datetime import datetime
from pathlib import Path
import pandas as pd
import uvicorn

# Importa a lógica do seu script original
from sms_auditado import extrair_tabela, classificar_transacao, garantir_ordem_colunas, aplicar_formatacao_excel, URL_LOGIN

app = FastAPI(title="BBTS SMS Consulta")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"]
)

# Pastas de trabalho
OUTPUTS_DIR = Path("outputs")
OUTPUTS_DIR.mkdir(exist_ok=True)

if not os.path.exists("static"):
    os.makedirs("static")

app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
def index():
    return FileResponse("static/index.html")

jobs = {}

class ConsultaRequest(BaseModel):
    numeros: str
    usuario: str
    senha: str
    url_login: str = URL_LOGIN
    t_login: int = 15
    t_token: int = 30

def run_consulta_task(job_id: str, req: ConsultaRequest):
    from playwright.sync_api import sync_playwright
    
    try:
        jobs[job_id]["status"] = "processing"
        jobs[job_id]["msg"] = "Iniciando navegador..."
        jobs[job_id]["pct"] = 10

        with sync_playwright() as p:
            # No Railway, precisamos rodar em modo headless (sem janela)
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            
            jobs[job_id]["msg"] = "Fazendo login (Aguardando tempos configurados)..."
            page.goto(req.url_login)
            
            # Preenche login se houver campos (ajuste os seletores se necessário)
            # page.fill("#user", req.usuario)
            # page.fill("#pass", req.senha)
            
            time.sleep(req.t_login)
            jobs[job_id]["pct"] = 30
            time.sleep(req.t_token)
            jobs[job_id]["pct"] = 50

            jobs[job_id]["msg"] = "Acessando Relatórios -> Consulta Celular..."
            page.locator('a[show-menu="relatorios"]').click()
            page.locator('a[ui-sref="consultaCelular"]').click()

            page.wait_for_selector("#inputTelefone", timeout=20000)
            page.locator("#inputTelefone").fill(req.numeros)
            
            jobs[job_id]["msg"] = "Consultando dados..."
            btn = page.get_by_role("button", name=re.compile(r"Consultar", re.I))
            btn.first.click()

            headers, dados = extrair_tabela(page)
            
            jobs[job_id]["msg"] = "Gerando Excel..."
            jobs[job_id]["pct"] = 80
            
            # Lógica de salvamento do Excel
            agora = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"consulta_{job_id}_{agora}.xlsx"
            excel_path = OUTPUTS_DIR / filename
            
            df = pd.DataFrame(dados) if dados else pd.DataFrame()
            if "Sms do arquivo" in df.columns:
                df["Transação Financeira"] = df["Sms do arquivo"].apply(classificar_transacao)
            
            df = garantir_ordem_colunas(df)
            df.to_excel(excel_path, index=False, engine="openpyxl")
            aplicar_formatacao_excel(str(excel_path))

            # Criar ZIP
            zip_name = f"resultado_{job_id}.zip"
            zip_path = OUTPUTS_DIR / zip_name
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                zipf.write(excel_path, arcname=filename)

            jobs[job_id]["status"] = "done"
            jobs[job_id]["pct"] = 100
            jobs[job_id]["msg"] = "Concluído!"
            jobs[job_id]["zip"] = str(zip_path)
            browser.close()

    except Exception as e:
        jobs[job_id]["status"] = "error"
        jobs[job_id]["msg"] = f"Erro: {str(e)}"

@app.post("/consultar")
def iniciar_consulta(req: ConsultaRequest, background_tasks: BackgroundTasks):
    job_id = str(uuid.uuid4())
    jobs[job_id] = {"status": "queued", "pct": 0, "msg": "Na fila...", "zip": None}
    background_tasks.add_task(run_consulta_task, job_id, req)
    return {"job_id": job_id}

@app.get("/status/{job_id}")
def get_status(job_id: str):
    return jobs.get(job_id, {"status": "error", "msg": "Não encontrado"})

@app.get("/download/{job_id}")
def download(job_id: str):
    job = jobs.get(job_id)
    if job and job["status"] == "done":
        return FileResponse(job["zip"], filename="resultado_consulta.zip")
    raise HTTPException(404, "Arquivo não pronto")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    uvicorn.run(app, host="0.0.0.0", port=port)