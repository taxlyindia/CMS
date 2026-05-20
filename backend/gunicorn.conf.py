# Gunicorn config — Hostinger VPS production
import multiprocessing, os

bind            = f"0.0.0.0:{os.environ.get('PORT', '5000')}"
workers         = min(multiprocessing.cpu_count() * 2 + 1, 9)
worker_class    = "sync"
timeout         = 120
keepalive       = 5
max_requests    = 1000
max_requests_jitter = 100
preload_app     = True
forwarded_allow_ips = "*"   # trust X-Forwarded-For from Nginx
accesslog       = "-"
errorlog        = "-"
loglevel        = os.environ.get("LOG_LEVEL", "info")
