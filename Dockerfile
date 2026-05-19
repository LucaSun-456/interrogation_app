# Optimized for small VPS / China mirrors (no build-essential; pure Python wheels)
FROM python:3.11-slim-bookworm

RUN groupadd -r appuser && useradd -r -g appuser -d /app appuser

WORKDIR /app

# Use Aliyun Debian mirror when building in China (faster apt)
RUN sed -i 's|deb.debian.org|mirrors.aliyun.com|g; s|security.debian.org|mirrors.aliyun.com|g' /etc/apt/sources.list.d/debian.sources 2>/dev/null || \
    sed -i 's|deb.debian.org|mirrors.aliyun.com|g; s|security.debian.org|mirrors.aliyun.com|g' /etc/apt/sources.list 2>/dev/null || true

RUN apt-get update && apt-get install -y --no-install-recommends \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt

COPY --chown=appuser:appuser . .

RUN mkdir -p /app/data /app/logs /app/materials /app/materials/prompts && chown -R appuser:appuser /app

USER appuser

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8000/api/health')" || exit 1

CMD ["gunicorn", "-c", "gunicorn_config.py", "wsgi:app"]
