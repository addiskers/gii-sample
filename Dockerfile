FROM python:3.11-slim

RUN apt-get update && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
    build-essential \
    libglib2.0-0 \
    libsm6 \
    libxrender1 \
    libxext6 \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt /app/requirements.txt

RUN python -m pip install --upgrade pip setuptools wheel \
    && pip install --no-cache-dir -r /app/requirements.txt

COPY . /app

RUN mkdir -p /app/generated_ppts /app/logs /app/templates

ENV PYTHONUNBUFFERED=1
ENV PYTHONIOENCODING=utf-8

RUN touch /app/logs/app.log /app/logs/timing.log && \
    chmod 666 /app/logs/app.log /app/logs/timing.log

EXPOSE 5000

CMD ["gunicorn", "--bind", "0.0.0.0:5000", "app:app", "--workers", "2", "--threads", "2", "--timeout", "120", "--capture-output", "--enable-stdio-inheritance"]