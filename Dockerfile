FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .
# Preserve the tracked baseline DB outside the mounted data volume.
RUN mkdir -p /app/data /bootstrap && cp /app/data/hr.db /bootstrap/hr.db

EXPOSE 8080

CMD ["uvicorn", "src.app:app", "--host", "0.0.0.0", "--port", "8080"]
