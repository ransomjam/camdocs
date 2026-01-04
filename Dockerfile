FROM python:3.11-slim

# Install system dependencies including LibreOffice and Java (required for LibreOffice)
RUN apt-get update && apt-get install -y \
    libreoffice \
    default-jre \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy backend requirements first to leverage Docker cache
COPY backend/requirements.txt backend/requirements.txt

# Install Python dependencies
# Note: docx2pdf will be skipped on Linux due to environment marker
RUN pip install --no-cache-dir -r backend/requirements.txt

# Copy the entire project
COPY . .

# Create necessary directories for uploads/outputs if they don't exist
RUN mkdir -p backend/outputs backend/uploads backend/Cover\ Pages/poorly\ formatted\ samples

# Expose the port
EXPOSE 5000

# Run with Gunicorn
# --chdir backend: switch to backend directory
# pattern_formatter_backend:app : module:variable
# --bind 0.0.0.0:5000 : listen on all interfaces
CMD ["gunicorn", "--chdir", "backend", "pattern_formatter_backend:app", "--bind", "0.0.0.0:5000", "--timeout", "120"]
