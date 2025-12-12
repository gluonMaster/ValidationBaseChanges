#!/bin/bash
# Entrypoint script for KindEltern Web Django application
# Handles migrations and server startup

set -e

# Wait for database to be ready
echo "Waiting for database..."
echo "DB Host: ${POSTGRES_HOST:-localhost}, DB: ${POSTGRES_DB:-kindeltern}, User: ${POSTGRES_USER:-postgres}"

while ! python -c "
import sys
import os
import psycopg2
try:
    conn = psycopg2.connect(
        dbname=os.getenv('POSTGRES_DB', 'kindeltern'),
        user=os.getenv('POSTGRES_USER', 'postgres'),
        password=os.getenv('POSTGRES_PASSWORD', 'postgres'),
        host=os.getenv('POSTGRES_HOST', 'localhost'),
        port=os.getenv('POSTGRES_PORT', '5432')
    )
    conn.close()
    sys.exit(0)
except Exception as e:
    print(f'DB connection error: {e}', file=sys.stderr)
    sys.exit(1)
"; do
    echo "Database not ready, waiting..."
    sleep 2
done
echo "Database is ready!"

# Run migrations
echo "Applying database migrations..."
python manage.py migrate --noinput

# Collect static files (for production)
if [ "$DJANGO_DEBUG" = "False" ] || [ "$DJANGO_DEBUG" = "false" ] || [ "$DJANGO_DEBUG" = "0" ]; then
    echo "Collecting static files..."
    python manage.py collectstatic --noinput
fi

# Start server based on command argument
case "$1" in
    "runserver")
        echo "Starting Django development server..."
        exec python manage.py runserver 0.0.0.0:8000
        ;;
    "gunicorn")
        echo "Starting Gunicorn production server..."
        exec gunicorn config.wsgi:application --bind 0.0.0.0:8000 --workers 3
        ;;
    *)
        # Run any other command
        exec "$@"
        ;;
esac
