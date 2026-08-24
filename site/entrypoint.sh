#!/bin/bash
set -e

python manage.py migrate --noinput || true
python manage.py collectstatic --noinput || true

exec "$@"
