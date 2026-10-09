import base64
import hashlib
import hmac
import json
import os
import tempfile
import unittest
from unittest.mock import patch, Mock
import subscriptions as s

USER='U'+'a'*32
OTHER='U'+'b'*32

class SubscriptionsTest(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory()
        self.env=patch.dict(os.environ, SUBSCRIPTIONS_DB=self.tmp.name+'/data.db',LINE_CHANNEL_SECRET='test')
        self.env.start()
    def tearDown(self):
        self.env.stop();self.tmp.cleanup()
    def event(self,user,command,event_id):
        s.apply_event(dict(type='message',webhookEventId=event_id,source={'userId':user},message={'text':command}))
    def test_segments(self):
        self.event(USER,'每日訂閱','1');self.event(OTHER,'每週訂閱','2')
        self.assertEqual(s.recipients('daily'),[USER]);self.assertEqual(s.recipients('weekly'),[OTHER])
    def test_pause(self):
        self.event(USER,'每日訂閱','1');self.event(USER,'暫停訂閱','2')
        self.assertEqual(s.recipients('daily'),[])
    def test_replay(self):
        self.event(USER,'每日訂閱','1');self.event(USER,'暫停訂閱','2');self.event(USER,'每日訂閱','1')
        self.assertEqual(s.recipients('daily'),[])
    @patch('subscriptions.requests.post')
    def test_delivery_once(self,post):
        post.return_value=Mock(status_code=200)
        self.event(USER,'每日訂閱','1')
        self.assertEqual(s.send_subscribers('test','daily','report'),(1,0))
        self.assertEqual(s.send_subscribers('test','daily','report'),(0,1))
        self.assertEqual(post.call_count,1)
    @patch('subscriptions.requests.post')
    def test_ambiguous_not_resent(self,post):
        post.side_effect=s.requests.Timeout()
        self.event(USER,'每日訂閱','1')
        with self.assertRaises(RuntimeError):s.send_subscribers('test','daily','report')
        self.assertEqual(s.send_subscribers('test','daily','report'),(0,1))
        self.assertEqual(post.call_count,1)
    def test_webhook_signature(self):
        from webhook import app
        client=app.test_client();body=json.dumps({'events':[]}).encode()
        self.assertEqual(client.post('/webhook',data=body,content_type='application/json').status_code,403)
        signature=base64.b64encode(hmac.new(b'test',body,hashlib.sha256).digest()).decode()
        self.assertEqual(client.post('/webhook',data=body,content_type='application/json',headers={'X-Line-Signature':signature}).status_code,200)
    def test_plain_links(self):
        text=('report\n'*1000)+'https://www.ftnn.com.tw/news/584972'
        messages=s.line_messages(text)
        self.assertTrue(all(m['type']=='text' for m in messages))
        self.assertEqual(''.join(m['text'] for m in messages),text)

if __name__=='__main__':unittest.main()
