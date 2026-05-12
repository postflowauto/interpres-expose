FROM python:3.11-slim

# System-Abhängigkeiten:
# - poppler-utils (pdftoppm): PDF → JPG für Slide-Vorschau
# - libreoffice-impress: PPTX → PDF Konvertierung (Fallback wenn CloudConvert
#   nicht verfügbar / Credit leer). Pulls libreoffice-core mit ein (~400 MB).
# - Schriftarten: für deutsche Sonderzeichen + Canva-Fallback
RUN apt-get update && apt-get install -y --no-install-recommends \
    poppler-utils \
    libreoffice-impress \
    fonts-liberation \
    fonts-dejavu-core \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PORT=10000
EXPOSE 10000

CMD gunicorn app:app --bind 0.0.0.0:$PORT --timeout 600 --workers 1
