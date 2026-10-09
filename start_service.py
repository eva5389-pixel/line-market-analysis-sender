"""Run both applications behind one HTTP port and one persistent disk."""
import os
from pathlib import Path
import signal
import subprocess
import time


def main():
    for key in ('APP_PASSWORD', 'LINE_CHANNEL_ACCESS_TOKEN', 'LINE_TARGET_ID', 'LINE_CHANNEL_SECRET', 'SUBSCRIPTIONS_DB'):
        if not os.environ.get(key):
            raise RuntimeError(f'Missing required setting: {key}')
    Path(os.environ['SUBSCRIPTIONS_DB']).parent.mkdir(parents=True, exist_ok=True)
    port = int(os.environ.get('PORT', '10000'))
    config = '''pid /tmp/nginx.pid;
events { worker_connections 256; }
http {
    access_log off;
    error_log /dev/stderr warn;
    map $http_upgrade $connection_upgrade { default upgrade; '' close; }
    server {
        listen PORT_NUMBER;
        client_max_body_size 1m;
        location = /webhook { proxy_pass http://127.0.0.1:8000; }
        location = /healthz { proxy_pass http://127.0.0.1:8000; }
        location / {
            proxy_pass http://127.0.0.1:8501;
            proxy_http_version 1.1;
            proxy_set_header Upgrade $http_upgrade;
            proxy_set_header Connection $connection_upgrade;
            proxy_set_header Host $http_host;
            proxy_read_timeout 300s;
        }
    }
}
'''.replace('PORT_NUMBER', str(port))
    Path('/tmp/nginx.conf').write_text(config)
    commands = [
        ['gunicorn', '--bind', '127.0.0.1:8000', '--workers', '1', 'webhook:app'],
        ['streamlit', 'run', 'streamlit_app.py', '--server.address', '127.0.0.1', '--server.port', '8501', '--server.headless', 'true'],
        ['nginx', '-c', '/tmp/nginx.conf', '-g', 'daemon off;'],
    ]
    processes = []
    def stop(*_):
        raise SystemExit(0)
    signal.signal(signal.SIGTERM, stop)
    signal.signal(signal.SIGINT, stop)
    try:
        for command in commands:
            processes.append(subprocess.Popen(command))
        while all(p.poll() is None for p in processes):
            time.sleep(1)
        raise SystemExit(1)
    finally:
        for p in processes:
            if p.poll() is None: p.terminate()
        for p in processes:
            try: p.wait(timeout=10)
            except subprocess.TimeoutExpired: p.kill()

if __name__ == '__main__':
    main()
