FROM python:3.11-slim

# pandoc ضروري لدمج نص Markdown مع قالب Word وتوليد معادلات حقيقية
RUN apt-get update && \
    apt-get install -y --no-install-recommends pandoc && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PORT=10000
EXPOSE 10000

CMD ["gunicorn", "-w", "2", "-b", "0.0.0.0:10000", "--timeout", "300", "app:app"]
