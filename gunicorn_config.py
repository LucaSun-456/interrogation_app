import multiprocessing
import os

bind = "0.0.0.0:8000"
backlog = 2048

# Default 1 worker for small VPS (1–2 GB RAM). Set GUNICORN_WORKERS in .env to scale up.
workers = int(os.environ.get("GUNICORN_WORKERS", "1"))
worker_class = "sync"
worker_connections = 1000
timeout = 120
keepalive = 5

accesslog = "-"
errorlog = "-"
loglevel = os.environ.get("GUNICORN_LOG_LEVEL", "info")
access_log_format = '%(h)s %(l)s %(u)s %(t)s "%(r)s" %(s)s %(b)s "%(f)s" "%(a)s" %(L)s'

proc_name = "interrogation_app"


def on_starting(server):
    server.log.info("Gunicorn starting — interrogation app")


def post_fork(server, worker):
    server.log.info(f"Worker spawned (pid: {worker.pid})")
