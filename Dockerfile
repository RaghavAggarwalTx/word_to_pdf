FROM python:3.11-slim

# Install LibreOffice
RUN apt-get update && apt-get install -y libreoffice && rm -rf /var/lib/apt/lists/*

# Set workdir
WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app
COPY . .

# Run Flask app
CMD ["python", "app.py"]
