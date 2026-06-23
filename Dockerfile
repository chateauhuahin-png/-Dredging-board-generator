FROM python:3.10-slim

# Install system dependencies
RUN apt-get update && apt-get install -y \
    libreoffice \
    poppler-utils \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python packages
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy project files
COPY . .

# Register Thai fonts
RUN mkdir -p /usr/share/fonts/thai && \
    cp fonts/*.ttf /usr/share/fonts/thai/ && \
    fc-cache -fv

# Create upload/output directories
RUN mkdir -p uploads outputs

EXPOSE 8080

CMD ["python", "app.py"]
