ARG BASE_IMAGE=docker.m.daocloud.io/library/python:3.12-slim
FROM ${BASE_IMAGE}

ARG PYPI_INDEX_URL=https://pypi.tuna.tsinghua.edu.cn/simple

ENV HOME=/app \
    PYTHONUNBUFFERED=1 \
    PYTHONUTF8=1 \
    PYTHONIOENCODING=utf-8 \
    UV_LINK_MODE=copy \
    VIRTUAL_ENV=/app/.venv \
    PATH=/app/.venv/bin:$PATH \
    OUTLOOK_CONFIG_FILE=/config/outlook-mcp.json \
    OUTLOOK_MCP_DB=/data/outlook-mcp.sqlite3 \
    OUTLOOK_MCP_UPLOAD_DIR=/data/uploads \
    OUTLOOK_MCP_HTTP_HOST=0.0.0.0 \
    OUTLOOK_MCP_HTTP_PORT=8000

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends tzdata \
    && rm -rf /var/lib/apt/lists/*

RUN PIP_INDEX_URL=${PYPI_INDEX_URL} pip install --no-cache-dir uv

COPY pyproject.toml uv.lock README.md ./
COPY src ./src

RUN uv sync --frozen --no-dev
RUN mkdir -p /config /data/uploads && chown -R 1000:1000 /config /data /app

USER 1000:1000

HEALTHCHECK --interval=30s --timeout=8s --retries=3 --start-period=10s \
    CMD python -c "import os,urllib.request; urllib.request.urlopen('http://127.0.0.1:'+os.environ.get('OUTLOOK_MCP_HTTP_PORT','8000')+'/health', timeout=8)" || exit 1

EXPOSE 8000

CMD ["outlook-mcp-http"]
