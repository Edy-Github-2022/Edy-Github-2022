from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

import os
import uuid
import zipfile
from pathlib import Path

from sms_auditado import executar_consulta, URL_LOGIN

app = FastAPI(title="BBTS SMS Consulta")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"]
)

OUTPUTS_DIR = Path("outputs")
OUTPUTS_DIR.mkdir(exist_ok=True)

app.mount("/static", StaticFiles(directory="static"), name="static")

jobs = {}


class ConsultaRequest(BaseModel):
    numeros: str
    usuario: str
    senha: str
    url_login: str = URL_LOGIN
    t_login: int = 15
    t_token: int = 30


@app.get("/")
def index():
    return FileResponse("static/index.html")


def run_consulta(job_id: str, req: ConsultaRequest):
    try:
        jobs[job_id]["status"] = "processing"
        jobs[job_id]["pct"] = 10
        jobs[job_id]["msg"] = "Executando consulta..."

        resultado = executar_consulta(
            numeros=req.numeros,
            usuario=req.usuario,
            senha=req.senha,
            url_login=req.url_login,
            t_login=req.t_login,
            t_token=req.t_token,
            headless=True
        )

        jobs[job_id]["pct"] = 85
        jobs[job_id]["msg"] = "Compactando arquivos..."

        zip_path = OUTPUTS_DIR / f"resultado_{job_id}.zip"

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            if resultado.get("excel_path") and Path(resultado["excel_path"]).exists():
                zipf.write(resultado["excel_path"], arcname=Path(resultado["excel_path"]).name)

            if resultado.get("screenshot_path") and Path(resultado["screenshot_path"]).exists():
                zipf.write(resultado["screenshot_path"], arcname=Path(resultado["screenshot_path"]).name)

        jobs[job_id]["status"] = "done"
        jobs[job_id]["pct"] = 100
        jobs[job_id]["msg"] = "Consulta finalizada com sucesso."
        jobs[job_id]["zip"] = str(zip_path)

    except Exception as e:
        jobs[job_id]["status"] = "error"
        jobs[job_id]["msg"] = str(e)
        jobs[job_id]["pct"] = 100


@app.post("/consultar")
def iniciar_consulta(req: ConsultaRequest, background_tasks: BackgroundTasks):
    job_id = str(uuid.uuid4())

    jobs[job_id] = {
        "status": "queued",
        "pct": 0,
        "msg": "Na fila...",
        "zip": None
    }

    background_tasks.add_task(run_consulta, job_id, req)
    return {"job_id": job_id}


@app.get("/status/{job_id}")
def status(job_id: str):
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job não encontrado")
    return job


@app.get("/download/{job_id}")
def download(job_id: str):
    job = jobs.get(job_id)

    if not job or job.get("status") != "done":
        raise HTTPException(status_code=400, detail="Job não concluído")

    zip_path = job.get("zip")
    if not zip_path or not Path(zip_path).exists():
        raise HTTPException(status_code=404, detail="Arquivo não encontrado")

    return FileResponse(
        zip_path,
        media_type="application/zip",
        filename=Path(zip_path).name
    )


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8080))
    uvicorn.run("main:app", host="0.0.0.0", port=port)