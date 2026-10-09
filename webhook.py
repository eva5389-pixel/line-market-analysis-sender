"""Deploy behind HTTPS on the same persistent volume as the sender."""
import base64
import hashlib
import hmac
import os
from flask import Flask, request, abort
from subscriptions import apply_event

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024

@app.post('/webhook')
def webhook():
    secret = os.environ.get('LINE_CHANNEL_SECRET')
    if not secret: abort(503)
    body = request.get_data()
    expected = base64.b64encode(hmac.new(secret.encode(), body, hashlib.sha256).digest()).decode()
    if not hmac.compare_digest(expected, request.headers.get('X-Line-Signature', '')): abort(403)
    payload = request.get_json()
    for event in payload.get('events', []): apply_event(event)
    return '', 200
