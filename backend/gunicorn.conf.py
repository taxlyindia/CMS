# Gunicorn config for Hostinger VPS
import multiprocessing

bind        = "0.0.0.0:5000"
workers     = multiprocessing.cpu_count() * 2 + 1
worker_class = "sync"
timeout     = 120
keepalive   = 5
max_requests = 1000
max_requests_jitter = 100
preload_app = True
accesslog   = "-"
errorlog    = "-"
loglevel    = "info"
