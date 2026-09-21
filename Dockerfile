FROM python:3.12-slim AS wheel-builder

ENV PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PIP_NO_CACHE_DIR=1

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        build-essential \
        libcairo2-dev \
        libfreetype6-dev \
        pkg-config \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /build
COPY requirements.txt .
RUN python -m pip wheel --wheel-dir /wheels -r requirements.txt


FROM python:3.12-slim

ENV APP_HOST=0.0.0.0 \
    APP_PORT=8877 \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    TZ=Asia/Kolkata

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libcairo2 \
        libfreetype6 \
        fonts-liberation2 \
        libpango-1.0-0 \
        libpangocairo-1.0-0 \
        tzdata \
    && rm -rf /var/lib/apt/lists/* \
    && groupadd --system --gid 10001 sticker \
    && useradd --system --uid 10001 --gid sticker --home-dir /app sticker

WORKDIR /app
COPY --from=wheel-builder /wheels /wheels
RUN python -m pip install --no-index --find-links=/wheels /wheels/* \
    && rm -rf /wheels

COPY --chown=sticker:sticker server.py ./server.py
COPY --chown=sticker:sticker static ./static

USER sticker
EXPOSE 8877

HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
    CMD ["python", "-c", "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8877/api/health', timeout=3).read()"]

CMD ["python", "server.py"]
