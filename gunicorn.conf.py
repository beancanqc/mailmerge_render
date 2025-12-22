# Gunicorn Configuration for Render
bind = "0.0.0.0:10000"
workers = 2
timeout = 300
keepalive = 2
max_requests = 1000
max_requests_jitter = 100