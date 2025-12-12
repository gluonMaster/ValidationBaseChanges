# Dockerfile for KindEltern Web Django application

FROM python:3.12-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Set work directory
WORKDIR /app

# Install system dependencies
# - libpq-dev, gcc: Required for psycopg2 PostgreSQL adapter
# - unixodbc, unixodbc-dev: Required for pyodbc (Access database import)
# NOTE: For Access import on Linux, you would also need mdbtools or similar.
#       Access import is primarily designed for Windows hosts with ACE driver.
RUN apt-get update && apt-get install -y \
    libpq-dev \
    gcc \
    unixodbc \
    unixodbc-dev \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy project
COPY . .

# Copy entrypoint script
COPY entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

# Expose port
EXPOSE 8000

# Use entrypoint script
ENTRYPOINT ["/entrypoint.sh"]

# Default command (can be overridden)
CMD ["runserver"]
