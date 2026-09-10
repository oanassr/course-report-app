# ── Hugging Face Spaces / Docker deployment ──────────────────────────────
# HF Spaces exposes port 7860 by default. The server auto-detects the
# HF_SPACE env-var and listens on 7860 / 0.0.0.0.

FROM python:3.11-slim

# System dependencies for PyMuPDF (libGL) and python-bidi.
RUN apt-get update && apt-get install -y --no-install-recommends \
        libgl1 \
        libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies first (layer cache).
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source.
COPY course_report_app/ ./course_report_app/

# HF_SPACE tells server.py to bind 0.0.0.0:7860.
ENV HF_SPACE=1
EXPOSE 7860

CMD ["python", "course_report_app/server.py"]
