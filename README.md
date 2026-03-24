# 📡 BBTS · Consultar SMS e PDUs

Aplicação web para consulta automatizada de SMS e PDUs no sistema BBTS.
Preencha o número + credenciais no site, e receba um ZIP com a planilha Excel tratada + screenshot de evidência.

---

## 🚀 Deploy no Railway (recomendado)

### 1. Criar repositório no GitHub

```bash
git init
git add .
git commit -m "feat: bbts sms consulta web"
git remote add origin https://github.com/SEU-USUARIO/bbts-sms.git
git push -u origin main
```

### 2. Subir no Railway

1. Acesse [railway.app](https://railway.app) → **New Project**
2. Clique em **Deploy from GitHub repo**
3. Selecione o repositório `bbts-sms`
4. Railway detecta o `Dockerfile` automaticamente e faz o build
5. Vá em **Settings → Networking → Generate Domain** para gerar o link público

> O link gerado (ex: `https://bbts-sms.up.railway.app`) é o que você compartilha com amigos.

---

## 🏃 Rodar localmente

```bash
# Instalar dependências
pip install -r requirements.txt
playwright install chromium

# Rodar
uvicorn main:app --reload --port 8000
```

Acesse: http://localhost:8000

---

## 📦 O que é gerado

Ao concluir a consulta, o sistema entrega um arquivo `.zip` contendo:

| Arquivo | Descrição |
|---|---|
| `consulta_YYYY-MM-DD.xlsx` | Planilha Excel formatada com tabela, cores condicionais na coluna Status, wrap text no PDU |
| `evidencia_YYYY-MM-DD.png` | Screenshot da tela no momento da extração |

### Colunas da planilha

| Col | Nome | Descrição |
|---|---|---|
| A | Identificador | ID único do envio |
| B | Número | Telefone consultado |
| C | Arquivo | Nome do arquivo de envio |
| D | C.Custo | Centro de custo |
| E | Cliente | Nome do cliente |
| F | Data envio do arquivo | — |
| G | Início do envio | — |
| H | Fim do envio | — |
| I | Status | ENTREGUE / NÃO ENTREGUE / EXPIRADO / ENVIADO (com cores) |
| J | PDU | JSON completo do PDU |
| K | Sms do arquivo | Conteúdo do SMS |
| L | Transação Financeira | Com Transação / Sem Transação (classificação automática) |

---

## ⚙️ Stack

- **Backend:** FastAPI + Playwright (Chromium headless)
- **Frontend:** HTML/CSS/JS puro (sem frameworks)
- **Excel:** pandas + openpyxl
- **Deploy:** Docker → Railway
