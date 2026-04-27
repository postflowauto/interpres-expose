FROM python:3.11-slim

# System-Abhängigkeiten:
# - poppler-utils (pdftoppm): PDF → JPG für Slide-Vorschau (~30 MB, leichtgewichtig)
# - PPTX → PDF läuft über CloudConvert API (siehe CLOUDCONVERT_KEY env),
#   spart ~500 MB RAM gegenüber lokalem LibreOffice headless
# - Schriftarten: Canva-Fallback + DejaVu für deutsche Sonderzeichen
RUN apt-get update && apt-get install -y --no-install-recommends \
    poppler-utils \
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
