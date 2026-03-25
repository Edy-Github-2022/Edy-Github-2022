FROM python:3.11-slim

# Instalar dependências necessárias para o Playwright/Chromium
RUN apt-get update && apt-get install -y \
    wget gnupg ca-certificates \
    libglib2.0-0 libnss3 libnspr4 libatk1.0-0 libatk-bridge2.0-0 \
    libx11-6 libxcomposite1 libxdamage1 libxfixes3 libxrandr2 libgbm1 \
    libasound2 libatspi2.0-0 libxkbcommon0 libcups2 libdrm2 && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
RUN playwright install chromium

COPY . .

RUN mkdir -p outputs static

# ✅ O Railway usa $PORT automaticamente
EXPOSE 8080

# ✅ O mais importante: usar exec form e variável direta
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port $PORT"]
