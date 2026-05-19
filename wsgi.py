"""WSGI entry point for Gunicorn."""
from app import app

# app.py runs initialize_app() at import time
# The 'app' object is what Gunicorn serves
