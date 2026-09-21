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
RUN mkdir -p /app/static/fonts \
    && python -c "import urllib.request; urllib.request.urlretrieve('https://raw.githubusercontent.com/google/fonts/c778ad7561ec09bb553e6a6569d893b1cb52372f/ofl/playfairdisplay/PlayfairDisplay-Bold.ttf', '/app/static/fonts/PlayfairDisplay-Bold.ttf'); urllib.request.urlretrieve('https://myjiostatic.cdn.jio.com/jiostaticresources/v05/ds20Fonts/JioType/JioType-Bold.ttf', '/app/static/fonts/JioType-Bold.ttf')" \
    && echo "37fcc00a7503976693088bed271e64ce34fca5119181da9acf5b74bd59da8c7e  /app/static/fonts/PlayfairDisplay-Bold.ttf" | sha256sum -c - \
    && echo "89b5173b6c4a700ccddcee31d365e41325883749cdd04685edc969d3946a16a2  /app/static/fonts/JioType-Bold.ttf" | sha256sum -c - \
    && chmod 0644 /app/static/fonts/*.ttf

USER sticker
EXPOSE 8877

HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
    CMD ["python", "-c", "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8877/api/health', timeout=3).read()"]

CMD ["python", "server.py"]
