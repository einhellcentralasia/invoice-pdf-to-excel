FROM python:3.11-slim

# Install system dependencies
RUN apt-get update && apt-get install -y \
    ghostscript libgl1 libglib2.0-0 poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY . /app

RUN pip install --no-cache-dir -r requirements.txt

EXPOSE 8000
CMD ["bash", "start.sh"]
