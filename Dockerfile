FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    HOST=0.0.0.0 \
    PORT=10000 \
    DATA_DIR=/opt/app/data

WORKDIR /opt/app

RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-writer fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /opt/app/data

EXPOSE 10000

CMD ["gunicorn", "--bind", "0.0.0.0:10000", "--workers", "2", "webapp:app"]
