import asyncio
import os
import re
import time
import uuid
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

# Excel
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Border, Side, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle

app = FastAPI(title="BBTS SMS Consulta")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

OUTPUTS_DIR = Path("outputs")
OUTPUTS_DIR.mkdir(exist_ok=True)

# Serve static files (frontend)
app.mount("/static", StaticFiles(directory="static"), name="static")

# ── Estado dos jobs ──────────────────────────────────────────────────────────
jobs: dict[str, dict] = {}

# ── Models ───────────────────────────────────────────────────────────────────
class ConsultaRequest(BaseModel):
    numeros: str          # um ou mais, um por linha
    usuario: str
    senha: str
    url_login: str = "https://bbts.bestuse.com.br/#!/login"
    t_login: int = 15
    t_token: int = 30

# ── Helpers Excel ─────────────────────────────────────────────────────────────
def classificar_transacao(mensagem: str) -> str:
    if not isinstance(mensagem, str) or not mensagem.strip():
        return "Sem Transação"
    texto = mensagem.lower()
    tem_valor = bool(re.search(r"r\$\s*\d", texto, re.IGNORECASE))
    termos = [
        "pix","ted","doc","pagamento","pagto","transferência","transferencia",
        "agendamento","agendado","agendada","compra","boleto","fatura",
        "saque","depósito","deposito","cobrança","cobranca","crédito","credito",
        "débito","debito",
    ]
    if any(t in texto for t in termos) or tem_valor:
        return "Com Transação"
    return "Sem Transação"

def garantir_ordem_colunas(df: pd.DataFrame) -> pd.DataFrame:
    ordem = [
        "Identificador","Número","Arquivo","C.Custo","Cliente",
        "Data envio do arquivo","Início do envio","Fim do envio",
        "Status","PDU","Sms do arquivo","Transação Financeira",
    ]
    for col in ordem:
        if col not in df.columns:
            df[col] = ""
    return df.reindex(columns=ordem)

def aplicar_formatacao_excel(excel_path: str, sheet_name: str = "Dados"):
    wb = load_workbook(excel_path)
    ws = wb[sheet_name]
    last_row = ws.max_row
    last_col_letter = get_column_letter(ws.max_column)
    used = f"A1:{last_col_letter}{last_row}"

    if last_row >= 2:
        try:
            tbl = Table(displayName=f"TabelaSMS_{int(time.time())}", ref=used)
            tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
            ws.add_table(tbl)
        except Exception:
            pass

    thin = Side(border_style="thin", color="FFBFBFBF")
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws[used]:
        for cell in row:
            cell.border = brd

    for cl in ("J", "K"):
        if ws[cl + "1"].value is not None:
            for cell in ws[f"{cl}1:{cl}{last_row}"]:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    ws.freeze_panes = "A2"

    widths = {"A":34,"B":18,"C":60,"D":30,"E":40,"F":22,"G":22,"H":22,"I":16,"J":88,"K":88,"L":22}
    for c, w in widths.items():
        try:
            ws.column_dimensions[c].width = w
        except Exception:
            pass

    if last_row >= 2:
        sr = f"I2:I{last_row}"
        def add_rule(rng, texto, hex_cor):
            fill = PatternFill(start_color=hex_cor, end_color=hex_cor, fill_type="solid")
            ws.conditional_formatting.add(rng, Rule(
                type="containsText", operator="containsText",
                text=texto, dxf=DifferentialStyle(fill=fill),
            ))
        for t in ["NÃO ENTREGUE","NAO ENTREGUE","Não entregue"]: add_rule(sr, t, "FFFF0000")
        for t in ["EXPIRADO","Expirado"]:                         add_rule(sr, t, "FFFFFF00")
        for t in ["ENTREGUE","Entregue"]:                         add_rule(sr, t, "FF00B050")
        for t in ["ENVIADO","Enviado"]:                           add_rule(sr, t, "FF800080")

    wb.save(excel_path)
    wb.close()

# ── Playwright helpers ────────────────────────────────────────────────────────
def set_status(job_id: str, status: str, pct: int, msg: str):
    jobs[job_id].update({"status": status, "pct": pct, "msg": msg})

def esperar_tabela(page, sel="table.table.table-striped.table-hover", timeout_ms=45000):
    page.wait_for_selector(sel, state="attached", timeout=timeout_ms)
    page.wait_for_selector(f"{sel} thead th", state="visible", timeout=timeout_ms)
    ths = page.locator(f"{sel} thead th")
    headers = [re.sub(r"\s+", " ", ths.nth(i).inner_text().strip()) for i in range(ths.count())]
    elapsed = 0
    while elapsed < timeout_ms:
        total = page.locator(f"{sel} tbody tr").count()
        if total > 0:
            return headers, total
        page.wait_for_timeout(800)
        elapsed += 800
    return headers, 0

def capturar_pdu(row):
    try:
        row.scroll_into_view_if_needed()
    except Exception:
        pass
    pre = row.locator("pre.ng-binding, pre")
    if pre.count() > 0:
        try:
            pre.first.wait_for(state="visible", timeout=3000)
            return pre.first.inner_text().strip()
        except Exception:
            pass
    link = row.get_by_role("link", name=re.compile(r"(Ver|Ocultar)\s+PDU", re.I))
    if link.count() == 0:
        link = row.locator("a:has-text('Ver PDU'), a:has-text('Ocultar PDU')")
    if link.count() > 0:
        try:
            if "ver pdu" in link.first.inner_text().strip().lower():
                link.first.click()
                row.locator("pre").first.wait_for(state="visible", timeout=5000)
            pre = row.locator("pre.ng-binding, pre")
            if pre.count() > 0:
                return pre.first.inner_text().strip()
        except Exception:
            return ""
    return ""

def extrair_tabela(page, sel="table.table.table-striped.table-hover"):
    headers, total = esperar_tabela(page, sel)
    dados = []
    if total == 0:
        return headers, dados
    col_pdu = headers.index("PDU") if "PDU" in headers else -1
    trs = page.locator(f"{sel} tbody tr")
    for i in range(total):
        row = trs.nth(i)
        tds = row.locator("td")
        vals = [re.sub(r"\s+", " ", tds.nth(c).inner_text().strip()) for c in range(tds.count())]
        m = min(len(headers), len(vals))
        row_dict = {headers[j]: vals[j] for j in range(m)}
        for h in headers[m:]:
            row_dict[h] = ""
        if col_pdu >= 0:
            try:
                pdu = capturar_pdu(row)
                if pdu:
                    row_dict["PDU"] = pdu
            except Exception:
                pass
        dados.append(row_dict)
    return headers, dados

# ── Job principal (síncrono, roda em thread) ──────────────────────────────────
def run_consulta(job_id: str, req: ConsultaRequest):
    from playwright.sync_api import sync_playwright

    job_dir = OUTPUTS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    agora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    excel_path   = str(job_dir / f"consulta_{agora}.xlsx")
    screenshot_path = str(job_dir / f"evidencia_{agora}.png")
    zip_path     = str(job_dir / f"resultado_{agora}.zip")

    dados, headers = [], []

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=["--no-sandbox","--disable-dev-shm-usage","--disable-gpu"],
            )
            page = browser.new_page()

            # ── 1. Abrir login ──────────────────────────────────────────────
            set_status(job_id, "running", 5, "Abrindo sistema...")
            page.goto(req.url_login)
            page.wait_for_load_state("networkidle", timeout=30000)

            # ── 2. Preencher credenciais ────────────────────────────────────
            set_status(job_id, "running", 15, "Preenchendo credenciais...")
            try:
                user_sel = 'input[name="usuario"], input[type="text"], #usuario, #login'
                pass_sel = 'input[name="senha"], input[type="password"], #senha, #password'
                page.wait_for_selector(user_sel, state="visible", timeout=10000)
                page.fill(user_sel, req.usuario)
                page.fill(pass_sel, req.senha)
                page.keyboard.press("Enter")
                page.wait_for_load_state("networkidle", timeout=15000)
            except Exception as e:
                set_status(job_id, "running", 20, f"Login manual necessário ({e})")

            # ── 3. Aguardar token 2FA ───────────────────────────────────────
            set_status(job_id, "running", 25, f"Aguardando autenticação ({req.t_token}s)...")
            page.wait_for_timeout(req.t_token * 1000)

            # ── 4. Menu Relatórios → Consulta Celular ───────────────────────
            set_status(job_id, "running", 35, "Navegando para Consulta Celular...")
            page.locator('a[show-menu="relatorios"]').click()
            page.wait_for_timeout(600)
            page.locator('a[ui-sref="consultaCelular"]').click()

            # ── 5. Preencher números ────────────────────────────────────────
            set_status(job_id, "running", 45, "Preenchendo número(s)...")
            page.wait_for_selector("#inputTelefone", state="visible", timeout=20000)
            page.locator("#inputTelefone").fill("")
            page.locator("#inputTelefone").fill(req.numeros.strip())

            # ── 6. Consultar ────────────────────────────────────────────────
            set_status(job_id, "running", 50, "Consultando...")
            btn = page.get_by_role("button", name=re.compile(r"Consultar", re.I))
            if btn.count() == 0:
                btn = page.locator("button:has-text('Consultar')")
            btn.first.click()

            # ── 7. Extrair tabela ───────────────────────────────────────────
            set_status(job_id, "running", 60, "Extraindo dados da tabela...")
            headers, dados = extrair_tabela(page)
            set_status(job_id, "running", 80, f"{len(dados)} registros capturados. Gerando evidência...")

            # ── 8. Screenshot ───────────────────────────────────────────────
            page.screenshot(path=screenshot_path, full_page=True)
            browser.close()

    except Exception as e:
        jobs[job_id].update({"status": "error", "pct": 0, "msg": f"Erro: {e}"})
        return

    # ── 9. Montar Excel ─────────────────────────────────────────────────────
    try:
        set_status(job_id, "running", 88, "Gerando planilha Excel...")
        df = pd.DataFrame(dados) if dados else pd.DataFrame()
        if "Sms do arquivo" in df.columns:
            df["Transação Financeira"] = df["Sms do arquivo"].apply(classificar_transacao)
        else:
            df["Transação Financeira"] = "Sem Transação"
        df = garantir_ordem_colunas(df)
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Dados")
        aplicar_formatacao_excel(excel_path)
    except Exception as e:
        jobs[job_id].update({"status": "error", "pct": 0, "msg": f"Erro Excel: {e}"})
        return

    # ── 10. Zipar ───────────────────────────────────────────────────────────
    set_status(job_id, "running", 95, "Compactando arquivos...")
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.write(excel_path,      Path(excel_path).name)
        if Path(screenshot_path).exists():
            zf.write(screenshot_path, Path(screenshot_path).name)

    jobs[job_id].update({
        "status": "done",
        "pct": 100,
        "msg": f"Concluído! {len(dados)} registros encontrados.",
        "zip": zip_path,
        "records": len(dados),
    })

# ── Rotas API ─────────────────────────────────────────────────────────────────
@app.get("/")
def index():
    return FileResponse("static/index.html")

@app.post("/consultar")
def iniciar_consulta(req: ConsultaRequest, background_tasks: BackgroundTasks):
    job_id = str(uuid.uuid4())
    jobs[job_id] = {"status": "queued", "pct": 0, "msg": "Na fila...", "zip": None}
    background_tasks.add_task(asyncio.get_event_loop().run_in_executor, None, run_consulta, job_id, req)
    return {"job_id": job_id}

@app.get("/status/{job_id}")
def status(job_id: str):
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(404, "Job não encontrado")
    return job

@app.get("/download/{job_id}")
def download(job_id: str):
    job = jobs.get(job_id)
    if not job or job.get("status") != "done":
        raise HTTPException(400, "Job não concluído")
    zip_path = job.get("zip")
    if not zip_path or not Path(zip_path).exists():
        raise HTTPException(404, "Arquivo não encontrado")
    return FileResponse(zip_path, media_type="application/zip", filename=Path(zip_path).name)
