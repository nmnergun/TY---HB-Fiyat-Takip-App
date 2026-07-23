FROM mcr.microsoft.com/playwright/python:v1.60.0-noble

WORKDIR /app

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PLAYWRIGHT_BROWSER_CHANNEL="" \
    REPORT_OUTPUT_DIR=/app/generated_reports \
    PORT=8080

COPY requirements-web.txt .
RUN pip install --no-cache-dir -r requirements-web.txt

COPY hepsiburada_price.py .
COPY hb_web_panel ./hb_web_panel
COPY start.sh /app/start.sh

RUN mkdir -p /app/generated_reports \
    && chmod +x /app/start.sh

EXPOSE 8080

CMD ["/app/start.sh"]
