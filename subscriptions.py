"""Persistent opt-in subscriptions and idempotent manual delivery."""
from contextlib import contextmanager
import hashlib
import os
import re
import sqlite3
import uuid
import requests
from news_links import line_messages


@contextmanager
def connect():
    path = os.environ.get('SUBSCRIPTIONS_DB')
    if not path:
        raise ValueError('請先設定持久化 SUBSCRIPTIONS_DB，再啟用好友訂閱。')
    db = sqlite3.connect(path, timeout=30)
    db.execute('CREATE TABLE IF NOT EXISTS subscribers (user_id TEXT PRIMARY KEY, frequency TEXT NOT NULL)')
    db.execute('CREATE TABLE IF NOT EXISTS events (event_id TEXT PRIMARY KEY)')
    db.execute('CREATE TABLE IF NOT EXISTS deliveries (user_id TEXT, digest TEXT, retry_key TEXT, status TEXT, PRIMARY KEY(user_id,digest))')
    db.commit()
    try:
        with db:
            yield db
    finally:
        db.close()


def apply_event(event):
    user = event.get('source', {}).get('userId', '')
    if not re.fullmatch(r'U[0-9a-fA-F]{32}', user):
        return
    kind = event.get('type')
    text = event.get('message', {}).get('text', '').strip()
    frequency = {'每日訂閱':'daily', '每週訂閱':'weekly', '暫停訂閱':'paused'}.get(text)
    if kind == 'unfollow': frequency = 'paused'
    if kind == 'follow': frequency = 'paused'  # No automatic opt-in.
    if frequency is None: return
    with connect() as db:
        event_id = event.get('webhookEventId')
        if not event_id: return
        if db.execute('INSERT OR IGNORE INTO events VALUES (?)', (event_id,)).rowcount == 0: return
        db.execute('INSERT INTO subscribers VALUES (?,?) ON CONFLICT(user_id) DO UPDATE SET frequency=excluded.frequency', (user, frequency))


def recipients(frequency):
    if frequency not in ('daily', 'weekly'): raise ValueError('訂閱頻率無效')
    with connect() as db:
        return [r[0] for r in db.execute('SELECT user_id FROM subscribers WHERE frequency=? ORDER BY user_id', (frequency,))]


def send_one(token, user, message):
    if not token or not re.fullmatch(r'U[0-9a-fA-F]{32}', user):
        raise ValueError('請設定有效的 LINE 憑證及個人 User ID。')
    payload = {'to':user, 'messages':line_messages(message)}
    response = requests.post('https://api.line.me/v2/bot/message/push', headers={'Authorization':'Bearer '+token}, json=payload, timeout=25)
    response.raise_for_status()


def send_subscribers(token, frequency, message):
    if not token: raise ValueError('缺少 LINE 憑證')
    messages = line_messages(message)
    digest = hashlib.sha256(message.encode()).hexdigest()
    sent = skipped = 0
    for user in recipients(frequency):
        # Transaction prevents two managers from claiming the same delivery.
        with connect() as db:
            db.execute('BEGIN IMMEDIATE')
            if not db.execute('SELECT 1 FROM subscribers WHERE user_id=? AND frequency=?', (user,frequency)).fetchone(): continue
            row = db.execute('SELECT retry_key,status FROM deliveries WHERE user_id=? AND digest=?', (user,digest)).fetchone()
            if row:  # Includes ambiguous network outcomes: never blindly resend.
                skipped += 1
                continue
            key = str(uuid.uuid4())
            db.execute('INSERT INTO deliveries VALUES (?,?,?,?)', (user,digest,key,'pending'))
        try:
            response = requests.post('https://api.line.me/v2/bot/message/push', headers={'Authorization':'Bearer '+token, 'X-Line-Retry-Key':key}, json={'to':user,'messages':messages}, timeout=25)
            response.raise_for_status()
        except requests.RequestException:
            raise RuntimeError('部分發送未確認，已保留紀錄防止重送；請查核 LINE 發送狀態。') from None
        with connect() as db:
            db.execute('UPDATE deliveries SET status=? WHERE user_id=? AND digest=?', ('sent',user,digest))
        sent += 1
    return sent, skipped
