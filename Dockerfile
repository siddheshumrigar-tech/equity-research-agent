# Headless runtime for cloud servers / agent sandboxes.
#   docker build -t equity-agent .
#   docker run --rm -v "$PWD/output:/app/output" equity-agent --ticker TATATECH
FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 MPLBACKEND=Agg
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENTRYPOINT ["python", "generate_equity_report.py"]
