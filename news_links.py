"""Resolve Google News wrappers; never present a failed wrapper as original news.
Protocol reference: https://github.com/SSujitX/google-news-url-decoder
Only public article identifiers are submitted to Google's own resolver.
"""
import json
from functools import lru_cache
from urllib.parse import urlparse, urlencode
import requests
from bs4 import BeautifulSoup


def publisher_url(url):
    try:
        p = urlparse(url)
        return p.scheme in ('http', 'https') and bool(p.hostname) and not p.username and not (p.hostname == 'google.com' or p.hostname.endswith('.google.com'))
    except (ValueError, TypeError):
        return False


@lru_cache(maxsize=256)
def resolve_article(url):
    if publisher_url(url):
        return url
    p = urlparse(url)
    if p.hostname != 'news.google.com' or p.path.split('/')[-2:-1] not in (['articles'], ['read']):
        return ''
    article = p.path.split('/')[-1]
    try:
        response = requests.get('https://news.google.com/rss/articles/' + article,
            params={'hl':'zh-TW','gl':'TW','ceid':'TW:zh-Hant'}, timeout=12, allow_redirects=False)
        response.raise_for_status()
        if response.is_redirect:
            target = response.headers.get('Location','')
            return target if publisher_url(target) else ''
        node = BeautifulSoup(response.text,'html.parser').select_one('[data-n-a-sg][data-n-a-ts]')
        if node is None:
            return ''
        context = [["X","X",["X","X"],None,None,1,1,"US:en",None,1,None,None,None,None,None,0,1],"X","X",1,[1,1,1],1,1,None,0,0,None,0]
        query = ['garturlreq',context,article,int(node['data-n-a-ts']),node['data-n-a-sg']]
        payload = [[['Fbv4je',json.dumps(query),None,'1']]]
        reply = requests.post('https://news.google.com/_/DotsSplashUi/data/batchexecute',
            data={'f.req':json.dumps(payload)},timeout=12)
        reply.raise_for_status()
        # The response may have an anti-XSSI prefix and length-delimited chunks.
        for line in reply.text.splitlines():
            if not line.lstrip().startswith('['):
                continue
            rows = json.loads(line)
            for row in rows:
                if isinstance(row,list) and len(row)>2 and row[0]=='wrb.fr' and row[1]=='Fbv4je':
                    result=json.loads(row[2])
                    if result[0]=='garturlres' and publisher_url(result[1]):
                        return result[1]
    except (requests.RequestException, ValueError, TypeError, KeyError, IndexError):
        pass
    return ''


def article_link(news):
    """Label search fallbacks explicitly; preserve publisher URLs intact."""
    original = resolve_article(news.get('url',''))
    if original:
        return '新聞原文', original
    title = news.get('title','').strip()
    return '原文連結暫時無法取得，改用標題搜尋', 'https://www.google.com/search?' + urlencode({'q':title})


def plain_line_messages(text, limit=4900):
    """Split on whole lines so a URL is never cut between messages."""
    chunks, current = [], ''
    size = lambda s: len(s.encode('utf-16-le')) // 2
    for line in text.splitlines(keepends=True):
        if size(line) > limit:
            raise ValueError('單行內容過長，請縮短後再傳送；未截斷網址。')
        if current and size(current + line) > limit:
            chunks.append(current)
            current = ''
        current += line
    if current:
        chunks.append(current)
    if len(chunks) > 5:
        raise ValueError('報告超過五則訊息容量，請縮短內容後再傳送。')
    return [{'type':'text','text':chunk} for chunk in chunks]


def line_messages(text, limit=4900):
    """Use ordinary text messages so recipients can copy and forward them."""
    return plain_line_messages(text, limit)
