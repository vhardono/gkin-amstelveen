FROM python:3.11-slim

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements and install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Expose port (Fly.io sets PORT env var)
EXPOSE 8080

# Start the application
CMD exec gunicorn app:app --bind 0.0.0.0:${PORT:-8080} --workers 2 --timeout 120
