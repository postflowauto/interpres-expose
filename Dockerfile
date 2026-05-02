FROM python:3.11-slim

# System-Abhängigkeiten:
# - poppler-utils (pdftoppm): PDF → JPG für V1-Slide-Vorschau
# - Schriftarten: Canva-Fallback + DejaVu für deutsche Sonderzeichen
# - Chromium-Runtime-Libs: für Playwright (V2-Renderer)
RUN apt-get update && apt-get install -y --no-install-recommends \
    poppler-utils \
    fonts-liberation \
    fonts-dejavu-core \
    fonts-noto-color-emoji \
    libnss3 \
    libnspr4 \
    libatk1.0-0 \
    libatk-bridge2.0-0 \
    libcups2 \
    libdrm2 \
    libxkbcommon0 \
    libxcomposite1 \
    libxdamage1 \
    libxfixes3 \
    libxrandr2 \
    libgbm1 \
    libpango-1.0-0 \
    libcairo2 \
    libasound2 \
    libatspi2.0-0 \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Playwright Chromium herunterladen (~150MB, einmalig beim Build)
RUN playwright install chromium --with-deps || playwright install chromium

COPY . .

ENV PORT=10000
EXPOSE 10000

CMD gunicorn app:app --bind 0.0.0.0:$PORT --timeout 600 --workers 1
