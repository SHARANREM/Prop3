FROM python:3.11-slim

# Install dependencies including LibreOffice
RUN apt-get update && \
    apt-get install -y libreoffice libreoffice-writer libreoffice-calc libreoffice-impress && \
    apt-get clean

# Set working dir
WORKDIR /app

# Copy all files
COPY . .

# Install Python deps
RUN pip install -r requirements.txt

# Expose port
EXPOSE 8000

# Start app with gunicorn
CMD ["gunicorn", "-b", "0.0.0.0:8000", "app:app"]
