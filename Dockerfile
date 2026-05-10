FROM python:3.11-slim

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements and install Python deps
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create required directories
RUN mkdir -p uploads outputs

# Expose port (Fly.io uses internal port)
EXPOSE 8080

# Use gunicorn for production
CMD ["sh", "-c", "python -c 'from app import init_db; init_db()' && gunicorn app:app --bind 0.0.0.0:${PORT:-8080} --workers 1 --timeout 300 --worker-class sync"]
