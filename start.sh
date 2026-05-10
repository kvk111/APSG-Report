#!/bin/bash
# APSG (Staging Ground) Report — startup script
python -c "from app import init_db; init_db()"
exec gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 300 --worker-class sync
