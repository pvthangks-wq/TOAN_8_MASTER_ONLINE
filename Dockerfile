FROM python:3.12-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
RUN apt-get update && apt-get install -y --no-install-recommends fonts-dejavu-core ca-certificates gosu && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py catalog.py ./
COPY web/ ./web/
COPY deploy/start.sh /usr/local/bin/start-toan8
RUN groupadd --system -g 10001 toan8 && useradd --system --uid 10001 --gid 10001 --home /app toan8 && mkdir -p /app/instance && chown -R toan8:toan8 /app/instance && chmod 755 /usr/local/bin/start-toan8
ENV DB_PATH=/app/instance/toan8.db COOKIE_SECURE=1
EXPOSE 8000 10000
ENTRYPOINT ["/usr/local/bin/start-toan8"]
