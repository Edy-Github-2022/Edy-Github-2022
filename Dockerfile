FROM python:3.11-slim

# Instalar dependências do sistema de uma vez só
RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    ca-certificates \
    libglib2.0-0 \
    libnss3 \
    libnspr4 \
    libatk1.0-0 \
    libatk-bridge2.0-0 \
    libx11-6 \
    libxcomposite1 \
    libxdamage1 \
    libxfixes3 \
    libxrandr2 \
    libgbm1 \
    libasound2 \
    libatspi2.0-0 \
    libxkbcommon0 \
    libcups2 \
    libdrm2 \
    libxshmfence1 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Instala apenas o executável do chromium
RUN playwright install chromium

COPY . .

RUN mkdir -p outputs static

# Comando de inicialização
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-8080}"]