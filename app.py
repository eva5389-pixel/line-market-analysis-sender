#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MoneyDJ Word Formatter v1.2 — 全自動版
======================================
雙擊 run.command 即可執行。
自動抓取 MoneyDJ 最新國際股市新聞 → 萃取內容 → 產出 Word 檔案。
"""

import csv
import os
import json
import re
import sys
import time
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime
from io import StringIO

import requests
import urllib3
from news_links import article_link, line_messages
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# MoneyDJ 的 SSL 憑證缺少 Subject Key Identifier，停用嚴格驗證
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# ============================================================================
# 全域設定
# ============================================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = SCRIPT_DIR  # 產出於腳本同目錄
LINE_CONFIG_PATH = os.path.join(SCRIPT_DIR, "line_config.json")
MARKET_CACHE_PATH = os.path.join(SCRIPT_DIR, "market_index_cache.json")

MARKET_SYMBOLS = [
    # 07:50（台北時間）報告先呈現已收盤的美國市場，再列亞洲與歐洲。
    ("美國標普500", "^GSPC", "SPY"),
    ("美國費城半導體", "^SOX", "SOXX"),
    ("美國那斯達克", "^IXIC", "QQQ"),
    ("美國道瓊", "^DJI", "DIA"),
    ("台灣加權", "^TWII", "EWT"),
    ("香港恆生", "^HSI", "EWH"),
    ("中國上證", "000001.SS", "FXI"),
    ("日本日經225", "^N225", "EWJ"),
    ("韓國KOSPI", "^KS11", "EWY"),
    ("歐洲STOXX50", "^STOXX50E", "FEZ"),
]

# LINE 個股觀察清單。只使用公開行情建立技術面觀察，不把它當成保證買賣訊號。
STOCK_WATCHLIST = [
    ("緯穎", "6669.TW"),
    ("南電", "8046.TW"),
    ("金像電", "2368.TW"),
    ("欣興", "3037.TW"),
    ("所羅門", "2359.TW"),
    ("晶豪科", "3006.TW"),
]

# 與 industry-supply-chain-dashboard 相同的研究分類摘要。僅表示同板塊，
# 不代表公司間存在直接供貨、投資或客戶關係。
SUPPLY_CHAIN_HOLDINGS = {
    "6669": {
        "sectors": ["AI伺服器"],
        "related": ["緯創", "廣達", "鴻海"],
    },
    "8046": {
        "sectors": ["博通供應鏈", "ABF載板"],
        "related": ["台積電", "日月光投控", "智邦", "啟碁", "欣興", "景碩"],
    },
    "2368": {
        "sectors": ["PCB／銅箔", "高階PCB", "CCL銅箔基板"],
        "related": ["金居", "健鼎", "臻鼎-KY", "欣興", "台光電", "聯茂"],
    },
    "3037": {
        "sectors": ["高階PCB", "ABF載板", "CCL銅箔基板"],
        "related": ["金像電", "健鼎", "景碩", "南電", "台光電", "聯茂"],
    },
    "2359": {
        "sectors": ["機器人"],
        "related": ["上銀", "台灣精銳", "和大", "達明機器人", "盟立", "鴻海"],
    },
    "3006": {
        "sectors": ["記憶體", "利基型記憶體IC"],
        "related": ["南亞科", "華邦電", "旺宏", "群聯", "威剛", "力成"],
    },
}

RELATED_TICKERS = {
    "台積電": "2330.TW", "日月光投控": "3711.TW", "智邦": "2345.TW",
    "啟碁": "6285.TW", "欣興": "3037.TW", "景碩": "3189.TW",
    "金居": "8358.TWO", "健鼎": "3044.TW", "臻鼎-KY": "4958.TW",
    "台光電": "2383.TW", "聯茂": "6213.TW", "金像電": "2368.TW",
    "南電": "8046.TW", "上銀": "2049.TW", "台灣精銳": "4583.TW",
    "和大": "1536.TW", "達明機器人": "4585.TW", "盟立": "2464.TW",
    "鴻海": "2317.TW", "緯創": "3231.TW", "廣達": "2382.TW",
    "南亞科": "2408.TW", "華邦電": "2344.TW", "旺宏": "2337.TW",
    "群聯": "8299.TWO", "威剛": "3260.TWO", "力成": "6239.TW",
}

FOREIGN_BROKER_HISTORY_URL = (
    "https://raw.githubusercontent.com/eva5389-pixel/"
    "taiwan-stock-dashboard/main/data/foreign_broker_history.csv"
)

MARKET_REFERENCE_LINKS = [
    ("Investing.com 全球主要指數", "https://www.investing.com/indices/major-indices"),
    ("玩股網全球股市", "https://www.wantgoo.com/global"),
]
UDN_MONEY_HOME = "https://money.udn.com/money/index"

# Nasdaq 官方公開報價可直接回傳這兩個指數的實際點位。
NASDAQ_INDEX_SYMBOLS = {
    "^IXIC": "COMP",
    "^SOX": "SOX",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
}

# 9 種新聞分類定義
CATEGORY_PATTERNS = [
    "《美股》", "《美債》", "《歐股》",
    "《陸股》", "《港股》", "《陸港股》",
    "《日股》", "《韓股》", "《日韓股》",
]

# 地區分組規則（合併為 4 大區塊）
REGION_MAPPING = {
    "《美股》": "美國",
    "《美債》": "美國",
    "《歐股》": "歐洲",
    "《陸股》": "陸港股",
    "《港股》": "陸港股",
    "《陸港股》": "陸港股",
    "《日股》": "亞股",
    "《韓股》": "亞股",
    "《日韓股》": "亞股",
}

REGION_ORDER = ["美國", "歐洲", "陸港股", "亞股"]

# MoneyDJ 主題代碼（用於抓取文章列表）
TOPIC_CODES = [
    "X0100014",  # 美股
    "X0100015",  # 歐股
    "X0100016",  # 日股
    "X0100017",  # 港股
    "X0100018",  # 韓股
]

# MoneyDJ 新聞列表頁（國際股市）
NEWS_LIST_URLS = [
    "https://www.moneydj.com/kmdj/news/newsreallist.aspx?a=CB010000",
    "https://www.moneydj.com/kmdj/news/newsreallist.aspx?a=CB020000",
]

# MoneyDJ 部分港股文章不一定帶有《港股》前綴，改以常見指數與市場詞補判斷。
HONG_KONG_TITLE_KEYWORDS = (
    "港股", "恆指", "恒指", "恆生", "恒生", "國企指數",
    "恒生科技", "恆生科技", "港交所", "香港股市",
)


def infer_category(title):
    """由標題前綴或市場關鍵字判斷新聞分類。"""
    for pattern in CATEGORY_PATTERNS:
        if pattern in title:
            return pattern
    if any(keyword in title for keyword in HONG_KONG_TITLE_KEYWORDS):
        return "《港股》"
    return None


# ============================================================================
# 模組一：資料索敵與驗證 (Data Scraping & Validation)
# ============================================================================

def log(msg):
    """帶時間戳的日誌輸出。"""
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}")


def fetch_page(url, timeout=15):
    """安全地抓取網頁，回傳 BeautifulSoup 物件或 None。"""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout, verify=False)
        resp.encoding = "utf-8"
        if resp.status_code == 200:
            return BeautifulSoup(resp.text, "html.parser")
    except Exception as e:
        log(f"  抓取失敗 {url}: {e}")
    return None


def discover_article_urls():
    """從 MoneyDJ 各頁面蒐集文章 URL（newsviewer.aspx?a=...）。"""
    found_urls = set()

    # 策略 1：從主題列表頁抓取
    for code in TOPIC_CODES:
        url = f"https://www.moneydj.com/kmdj/common/listnewarticles.aspx?svc=NW&a={code}"
        log(f"掃描主題頁 {code}...")
        soup = fetch_page(url)
        if soup:
            links = soup.find_all("a", href=re.compile(r"newsviewer\.aspx\?a="))
            for link in links:
                href = link.get("href", "")
                if "newsviewer.aspx?a=" in href:
                    # 取得完整 URL
                    if href.startswith("/"):
                        href = "https://www.moneydj.com" + href
                    elif not href.startswith("http"):
                        href = "https://www.moneydj.com/kmdj/news/" + href
                    found_urls.add(href)
        time.sleep(0.3)

    # 策略 2：從新聞列表頁抓取
    for list_url in NEWS_LIST_URLS:
        log(f"掃描新聞列表頁...")
        soup = fetch_page(list_url)
        if soup:
            links = soup.find_all("a", href=re.compile(r"newsviewer\.aspx\?a="))
            for link in links:
                href = link.get("href", "")
                if "newsviewer.aspx?a=" in href:
                    if href.startswith("/"):
                        href = "https://www.moneydj.com" + href
                    elif not href.startswith("http"):
                        href = "https://www.moneydj.com/kmdj/news/" + href
                    found_urls.add(href)
        time.sleep(0.3)

    # 策略 3：從 RSS Feed 抓取
    rss_urls = [
        "https://www.moneydj.com/kmdj/RssCenter.aspx?svc=NW&fno=1&arg=X0000000",
        "https://www.moneydj.com/kmdj/RssCenter.aspx?svc=NW&fno=1&arg=X0100000",
    ]
    for rss_url in rss_urls:
        log("掃描 RSS Feed...")
        try:
            resp = requests.get(rss_url, headers=HEADERS, timeout=10, verify=False)
            resp.encoding = "utf-8"
            if resp.status_code == 200:
                # 簡易 XML 解析找 link 標籤
                urls_in_rss = re.findall(
                    r'newsviewer\.aspx\?a=[0-9a-f\-]+', resp.text, re.IGNORECASE
                )
                for u in urls_in_rss:
                    full_url = "https://www.moneydj.com/kmdj/news/" + u
                    found_urls.add(full_url)
        except Exception:
            pass
        time.sleep(0.3)

    # 策略 4：從新聞首頁抓取
    log("掃描新聞首頁...")
    soup = fetch_page("https://www.moneydj.com/kmdj/news/newshome.aspx")
    if soup:
        links = soup.find_all("a", href=re.compile(r"newsviewer\.aspx\?a="))
        for link in links:
            href = link.get("href", "")
            if "newsviewer.aspx?a=" in href:
                if href.startswith("/"):
                    href = "https://www.moneydj.com" + href
                elif not href.startswith("http"):
                    href = "https://www.moneydj.com/kmdj/news/" + href
                found_urls.add(href)

    log(f"共發現 {len(found_urls)} 篇候選文章 URL")
    return list(found_urls)


def parse_article(url):
    """進入文章內頁，解析標題、發布日期、內文段落。
    回傳 dict 或 None。
    """
    soup = fetch_page(url)
    if not soup:
        return None

    # 取得頁面全文字
    page_text = soup.get_text(separator="\n")
    lines = [ln.strip() for ln in page_text.split("\n") if ln.strip()]

    # 解析標題（從 <title> 或 og:title）
    title = ""
    title_tag = soup.find("title")
    if title_tag:
        title = title_tag.get_text().strip()
        # 移除 " - MoneyDJ理財網" 後綴
        title = re.sub(r"\s*-\s*MoneyDJ.*$", "", title).strip()

    # 檢查是否符合目標分類
    category = infer_category(title)

    if not category:
        return None

    # 解析發布日期（格式：MoneyDJ新聞 2026-08-14 06:14:14 XXX 發佈）
    pub_date = None
    date_pattern = re.compile(r"MoneyDJ新聞\s+(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})")
    for line in lines:
        m = date_pattern.search(line)
        if m:
            try:
                pub_date = datetime.strptime(m.group(1), "%Y-%m-%d %H:%M:%S")
            except ValueError:
                pass
            break

    if not pub_date:
        return None

    # 解析內文（在日期行之後、相關新聞/標籤之前的文字）
    body_lines = []
    in_body = False
    for line in lines:
        if date_pattern.search(line):
            in_body = True
            continue
        if in_body:
            # 停止條件：遇到相關新聞、編者按、廣告、標籤等
            if any(stop in line for stop in [
                "＊編者按", "編者按", "圖片來源", "（圖片來源",
                "MoneyDJ理財網", "財經知識庫", "基金頻道",
                "新聞首頁", "最新頭條", "Apple", "Amazon",
                "Microsoft", "Tesla", "Alphabet", "Meta",
                "首頁", "研究報告", "財經百科", "Back To Top",
                "econsult@", "加入會員", "手機版", "iQuote",
            ]):
                break
            # 跳過過短的行或導航文字
            if len(line) > 15:
                body_lines.append(line)

    if not body_lines:
        return None

    return {
        "url": url,
        "title": title,
        "category": category,
        "region": REGION_MAPPING.get(category, "其他"),
        "pub_date": pub_date,
        "body_lines": body_lines,
    }


def select_latest_articles(all_articles, max_age_days=5, hong_kong_fallback_days=30):
    """針對每個分類，選取最新的一篇文章。
    若超過 max_age_days 天仍無結果，標示為放棄。
    """
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    cutoff = today - timedelta(days=max_age_days)

    # 按分類分組
    by_category = {}
    for art in all_articles:
        cat = art["category"]
        if cat not in by_category:
            by_category[cat] = []
        by_category[cat].append(art)

    selected = {}
    for cat in CATEGORY_PATTERNS:
        candidates = by_category.get(cat, [])
        # 篩選在時間範圍內的
        valid = [a for a in candidates if a["pub_date"] >= cutoff]
        if valid:
            # 取最新的一篇
            latest = max(valid, key=lambda x: x["pub_date"])
            selected[cat] = latest
            log(f"  {cat} → {latest['title']} ({latest['pub_date'].strftime('%Y-%m-%d %H:%M')})")
        elif cat == "《港股》" and candidates:
            # 港股專欄發文頻率有時較低；仍保留最近一篇，避免整個港股區塊消失。
            hk_cutoff = today - timedelta(days=hong_kong_fallback_days)
            fallback = [a for a in candidates if a["pub_date"] >= hk_cutoff]
            if fallback:
                latest = max(fallback, key=lambda x: x["pub_date"])
                selected[cat] = latest
                log(
                    f"  {cat} → {latest['title']} "
                    f"({latest['pub_date'].strftime('%Y-%m-%d %H:%M')}，港股備援)"
                )
            else:
                log(f"  {cat} → 過去 {hong_kong_fallback_days} 日內未找到")
        else:
            log(f"  {cat} → 過去 {max_age_days} 日內未找到")

    return selected


def build_hong_kong_market_fallback():
    """MoneyDJ 無港股文章時，以恆生指數最近兩個交易日產生港股行情摘要。"""
    source_url = "https://finance.yahoo.com/quote/%5EHSI/"
    api_urls = [
        "https://query2.finance.yahoo.com/v8/finance/chart/%5EHSI",
        "https://query1.finance.yahoo.com/v8/finance/chart/%5EHSI",
    ]
    try:
        response = None
        for api_url in api_urls:
            candidate = requests.get(
                api_url,
                params={"range": "10d", "interval": "1d"},
                headers=HEADERS,
                timeout=15,
            )
            if candidate.status_code == 200:
                response = candidate
                break
        if response is None:
            raise RuntimeError("Yahoo Finance 行情端點暫時受流量限制")
        result = response.json()["chart"]["result"][0]
        timestamps = result.get("timestamp", [])
        closes = result["indicators"]["quote"][0].get("close", [])
        observations = [
            (ts, close) for ts, close in zip(timestamps, closes)
            if close is not None
        ]
        if len(observations) < 2:
            return None

        (previous_ts, previous_close), (latest_ts, latest_close) = observations[-2:]
        change = latest_close - previous_close
        change_pct = change / previous_close * 100 if previous_close else 0.0
        direction = "上漲" if change > 0 else "下跌" if change < 0 else "持平"
        pub_date = datetime.fromtimestamp(latest_ts)
        title = f"《港股》恆生指數{direction}{abs(change_pct):.2f}%"
        paragraphs = [
            (
                f"恆生指數最新收盤為 {latest_close:,.2f} 點，較前一交易日"
                f"{direction} {abs(change):,.2f} 點，漲跌幅 {change_pct:+.2f}%。"
            ),
            "本段為 MoneyDJ 港股專欄暫缺時的行情備援摘要，以最近兩個交易日收盤價計算。",
            f"資料來源：Yahoo Finance 恆生指數（{source_url}）",
        ]
        return {
            "url": source_url,
            "title": title,
            "category": "《港股》",
            "region": "陸港股",
            "pub_date": pub_date,
            "body_lines": paragraphs,
        }
    except Exception as exc:
        log(f"  港股行情備援取得失敗：{exc}")
        return {
            "url": source_url,
            "title": "《港股》恆生指數行情待更新",
            "category": "《港股》",
            "region": "陸港股",
            "pub_date": datetime.now(),
            "body_lines": [
                "MoneyDJ 當日未刊登港股專欄，恆生指數行情端點亦暫時受流量限制。",
                "請稍後重新執行報告；程式會優先採用 MoneyDJ 文章，其次自動補入最新恆生指數行情。",
                f"行情來源：Yahoo Finance 恆生指數（{source_url}）",
            ],
        }


def scrape_all_news():
    """模組一主入口：搜尋、抓取、驗證所有新聞。"""
    log("=" * 60)
    log("模組一：資料索敵與驗證")
    log("=" * 60)

    # 1. 發現文章 URL
    urls = discover_article_urls()

    if not urls:
        log("無法從 MoneyDJ 發現任何文章 URL，請確認網路連線。")
        return {}

    # 2. 逐一解析文章
    log(f"\n開始解析 {len(urls)} 篇候選文章...")
    all_articles = []
    for i, url in enumerate(urls):
        art = parse_article(url)
        if art:
            all_articles.append(art)
        # 限制請求速率
        if (i + 1) % 10 == 0:
            log(f"  已解析 {i + 1}/{len(urls)} 篇...")
        time.sleep(0.2)

    log(f"共解析到 {len(all_articles)} 篇符合分類的文章")

    # 3. 選取每個分類的最新文章
    log("\n選取各分類最新文章：")
    selected = select_latest_articles(all_articles)

    # MoneyDJ 當日若未刊港股專欄，仍以最新恆生指數行情補齊港股內容。
    if "《港股》" not in selected and "《陸港股》" not in selected:
        hk_fallback = build_hong_kong_market_fallback()
        if hk_fallback:
            selected["《港股》"] = hk_fallback
            log(f"  《港股》 → {hk_fallback['title']}（行情備援）")

    return selected


# ============================================================================
# 每日大盤收盤與 LINE 綜合報告
# ============================================================================

def fetch_udn_hot_news(limit=5):
    """讀取經濟日報首頁「最熱」新聞；失敗時回傳空清單，不中斷早報。"""
    try:
        response = requests.get(UDN_MONEY_HOME, headers=HEADERS, timeout=20)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        items = []
        seen = set()
        for anchor in soup.find_all("a", href=True):
            href = anchor.get("href", "")
            if "from=edn_hottest_index" not in href:
                continue
            canonical = href.split("?", 1)[0]
            if canonical in seen:
                continue
            title = " ".join(anchor.get_text(" ", strip=True).split())
            title = re.sub(r"\s+20\d{2}-\d{2}-\d{2}.*$", "", title).strip()
            title = re.sub(r"\s+\d{3,}$", "", title).strip()
            if len(title) < 8:
                continue
            if canonical.startswith("/"):
                canonical = "https://money.udn.com" + canonical
            seen.add(canonical)
            items.append({"title": title, "url": canonical})
            if len(items) >= limit:
                break
        log(f"經濟日報熱門新聞取得 {len(items)} 則")
        return items
    except Exception as exc:
        log(f"經濟日報熱門新聞暫時無法取得：{exc}")
        return []

def fetch_nasdaq_proxy_close(name, proxy_symbol):
    """以 Nasdaq 公開報價取得市場代表 ETF 當日收盤，作為指數限流備援。"""
    try:
        response = requests.get(
            f"https://api.nasdaq.com/api/quote/{proxy_symbol}/info",
            params={"assetclass": "etf"},
            headers={**HEADERS, "Accept": "application/json"},
            timeout=15,
        )
        response.raise_for_status()
        primary = response.json()["data"]["primaryData"]
        close = float(re.sub(r"[^0-9.\-]", "", primary["lastSalePrice"]))
        daily_pct = float(primary["percentageChange"].replace("%", "").replace("+", ""))
        date_text = primary.get("lastTradeTimestamp", "")
        try:
            date_text = datetime.strptime(date_text, "%b %d, %Y").strftime("%Y-%m-%d")
        except ValueError:
            date_text = date_text or "—"
        return {
            "name": name,
            "symbol": proxy_symbol,
            "date": date_text,
            "close": close,
            "daily_pct": daily_pct,
            "weekly_pct": None,
            "status": f"ETF代理（{proxy_symbol}）",
            "source": "Nasdaq ETF代理",
        }
    except Exception as exc:
        log(f"  {name} Nasdaq ETF備援失敗：{exc}")
        return None


def fetch_nasdaq_index_close(name, symbol):
    """取得 Nasdaq Composite／費半實際指數點位，不使用 ETF 價格。"""
    nasdaq_symbol = NASDAQ_INDEX_SYMBOLS.get(symbol)
    if not nasdaq_symbol:
        return None
    try:
        response = requests.get(
            f"https://api.nasdaq.com/api/quote/{nasdaq_symbol}/info",
            params={"assetclass": "index"},
            headers={**HEADERS, "Accept": "application/json"},
            timeout=15,
        )
        response.raise_for_status()
        primary = response.json()["data"]["primaryData"]
        close = float(re.sub(r"[^0-9.\-]", "", primary["lastSalePrice"]))
        daily_pct = float(primary["percentageChange"].replace("%", "").replace("+", ""))
        raw_date = primary.get("lastTradeTimestamp", "")
        try:
            date_text = datetime.strptime(raw_date, "%b %d, %Y").strftime("%Y-%m-%d")
        except ValueError:
            date_text = raw_date or "—"
        return {
            "name": name,
            "symbol": symbol,
            "date": date_text,
            "close": close,
            "daily_pct": daily_pct,
            "weekly_pct": None,
            "status": "上漲" if daily_pct > 0 else "下跌" if daily_pct < 0 else "持平",
            "source": "Nasdaq實際指數",
        }
    except Exception as exc:
        log(f"  {name} Nasdaq實際指數取得失敗：{exc}")
        return None


def fetch_market_close(name, symbol, proxy_symbol):
    """取得指數最近收盤、日漲跌與近五個交易日漲跌。"""
    nasdaq_index = fetch_nasdaq_index_close(name, symbol)
    if nasdaq_index:
        return nasdaq_index
    # Yahoo spark 端點通常比 chart 穩定，且回傳的是指數本身而非 ETF。
    try:
        response = requests.get(
            "https://query1.finance.yahoo.com/v7/finance/spark",
            params={"symbols": symbol, "range": "5d", "interval": "1d"},
            headers=HEADERS,
            timeout=15,
        )
        response.raise_for_status()
        spark_result = response.json()["spark"]["result"][0]["response"][0]
        meta = spark_result.get("meta", {})
        timestamps = spark_result.get("timestamp", [])
        closes = spark_result.get("indicators", {}).get("quote", [{}])[0].get("close", [])
        observations = [(ts, float(close)) for ts, close in zip(timestamps, closes) if close is not None]
        if len(observations) >= 2:
            latest_ts, latest_close = observations[-1]
            previous_close = observations[-2][1]
            week_base = observations[0][1]
            offset = int(meta.get("gmtoffset", 0))
            daily_pct = (latest_close / previous_close - 1) * 100
            weekly_pct = (latest_close / week_base - 1) * 100
            return {
                "name": name,
                "symbol": symbol,
                "date": datetime.utcfromtimestamp(latest_ts + offset).strftime("%Y-%m-%d"),
                "close": latest_close,
                "daily_pct": daily_pct,
                "weekly_pct": weekly_pct,
                "status": "上漲" if daily_pct > 0 else "下跌" if daily_pct < 0 else "持平",
                "source": "Yahoo Finance實際指數",
            }
    except Exception as exc:
        log(f"  {name} Yahoo spark 指數取得失敗：{exc}")
    api_paths = (
        "https://query2.finance.yahoo.com/v8/finance/chart/",
        "https://query1.finance.yahoo.com/v8/finance/chart/",
    )
    last_error = None
    for api_base in api_paths:
        try:
            # Yahoo 的 range+interval 組合偶爾會被限流；range-only 會回傳
            # 近五日分鐘資料，再依交易所時區彙整成每日最後一筆。
            response = requests.get(
                api_base + requests.utils.quote(symbol, safe="") + "?range=5d",
                headers=HEADERS,
                timeout=15,
            )
            if response.status_code != 200:
                last_error = f"HTTP {response.status_code}"
                continue
            result = response.json()["chart"]["result"][0]
            timestamps = result.get("timestamp", [])
            closes = result["indicators"]["quote"][0].get("close", [])
            exchange_offset = int(result.get("meta", {}).get("gmtoffset", 0))
            daily_observations = {}
            for ts, close in zip(timestamps, closes):
                if close is None:
                    continue
                market_date = datetime.utcfromtimestamp(ts + exchange_offset).strftime("%Y-%m-%d")
                daily_observations[market_date] = (ts, float(close))
            observations = list(daily_observations.values())
            if len(observations) < 2:
                last_error = "有效收盤資料不足"
                continue
            latest_ts, latest_close = observations[-1]
            previous_close = observations[-2][1]
            week_base = observations[0][1]
            daily_pct = (latest_close / previous_close - 1) * 100
            weekly_pct = (latest_close / week_base - 1) * 100
            return {
                "name": name,
                "symbol": symbol,
                "date": datetime.utcfromtimestamp(latest_ts + exchange_offset).strftime("%Y-%m-%d"),
                "close": latest_close,
                "daily_pct": daily_pct,
                "weekly_pct": weekly_pct,
                "status": "上漲" if daily_pct > 0 else "下跌" if daily_pct < 0 else "持平",
                "source": "Yahoo Finance指數",
            }
        except Exception as exc:
            last_error = str(exc)
    log(f"  {name} 指數行情取得失敗：{last_error}；不以 ETF 價格代替指數")
    return {
        "name": name,
        "symbol": symbol,
        "date": "—",
        "close": None,
        "daily_pct": None,
        "weekly_pct": None,
        "status": "指數資料待更新",
        "source": "—",
    }


def fetch_global_market_closes():
    """單次批次抓取前一完整交易日，避免逐筆請求觸發限流。"""
    log("\n抓取每日主要市場收盤狀況：")
    batch_rows = {}
    batch_symbols = [symbol for _, symbol, _ in MARKET_SYMBOLS if symbol not in NASDAQ_INDEX_SYMBOLS]
    try:
        # Yahoo 對逗號被編碼成 %2C 的請求容易回 429；保留批次分隔逗號。
        encoded_symbols = requests.utils.quote(",".join(batch_symbols), safe=",")
        batch_url = (
            "https://query1.finance.yahoo.com/v7/finance/spark?symbols="
            f"{encoded_symbols}&range=5d&interval=1d"
        )
        response = requests.get(batch_url, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
        response.raise_for_status()
        for result in response.json().get("spark", {}).get("result", []):
            symbol = result.get("symbol")
            series = (result.get("response") or [{}])[0]
            meta = series.get("meta", {})
            timestamps = series.get("timestamp", [])
            closes = series.get("indicators", {}).get("quote", [{}])[0].get("close", [])
            offset = int(meta.get("gmtoffset", 0))
            report_date = datetime.now().strftime("%Y-%m-%d")
            observations = []
            for ts, close in zip(timestamps, closes):
                if close is None:
                    continue
                exchange_date = datetime.utcfromtimestamp(ts + offset).strftime("%Y-%m-%d")
                # 早報及白天手動執行都只使用今天以前的完整交易日。
                if exchange_date < report_date:
                    observations.append((ts, float(close), exchange_date))
            if len(observations) < 2:
                continue
            latest_ts, latest_close, latest_date = observations[-1]
            daily_pct = (latest_close / observations[-2][1] - 1) * 100
            weekly_pct = (latest_close / observations[0][1] - 1) * 100
            batch_rows[symbol] = {
                "symbol": symbol,
                "date": latest_date,
                "close": latest_close,
                "daily_pct": daily_pct,
                "weekly_pct": weekly_pct,
                "status": "上漲" if daily_pct > 0 else "下跌" if daily_pct < 0 else "持平",
                "source": "Yahoo Finance實際指數",
            }
    except Exception as exc:
        log(f"  批次指數來源暫時受限：{exc}")

    try:
        with open(MARKET_CACHE_PATH, "r", encoding="utf-8") as cache_file:
            cache = json.load(cache_file)
    except Exception:
        cache = {}

    rows = []
    for name, symbol, proxy_symbol in MARKET_SYMBOLS:
        if symbol in NASDAQ_INDEX_SYMBOLS:
            row = fetch_nasdaq_index_close(name, symbol)
        else:
            row = batch_rows.get(symbol)
            if row:
                row = {"name": name, **row}
        if not row:
            cached = cache.get(symbol)
            if (cached and cached.get("close") is not None
                    and str(cached.get("date", "")) < datetime.now().strftime("%Y-%m-%d")):
                row = {"name": name, **cached, "status": "最近成功指數資料"}
                log(f"  {name} 使用最近成功快取（{row['date']}）")
            else:
                row = fetch_market_close(name, symbol, proxy_symbol)
        rows.append(row)
        if row["close"] is not None:
            log(f"  {name} {row['close']:,.2f}（{row['daily_pct']:+.2f}%）")
        time.sleep(0.15)
    for row in rows:
        if row.get("close") is not None and row.get("symbol"):
            cache[row["symbol"]] = {key: row.get(key) for key in (
                "symbol", "date", "close", "daily_pct", "weekly_pct", "status", "source"
            )}
    try:
        with open(MARKET_CACHE_PATH, "w", encoding="utf-8") as cache_file:
            json.dump(cache, cache_file, ensure_ascii=False, indent=2)
    except Exception as exc:
        log(f"  指數快取無法寫入：{exc}")
    return rows


def build_market_analysis(market_rows):
    """將漲跌廣度、強弱市場與區域差異整理為簡短綜合判讀。"""
    valid = [row for row in market_rows if row["daily_pct"] is not None]
    if not valid:
        return "主要市場行情來源暫時無法連線，新聞摘要仍可正常使用。"

    up_count = sum(row["daily_pct"] > 0 for row in valid)
    down_count = sum(row["daily_pct"] < 0 for row in valid)
    strongest = max(valid, key=lambda row: row["daily_pct"])
    weakest = min(valid, key=lambda row: row["daily_pct"])
    avg_move = sum(row["daily_pct"] for row in valid) / len(valid)
    if up_count > down_count:
        tone = "風險偏好回升，市場廣度偏多"
    elif down_count > up_count:
        tone = "風險偏好降溫，市場廣度偏空"
    else:
        tone = "市場漲跌分歧，宜觀察區域輪動"
    return (
        f"最近完整交易日已取得 {len(valid)} 個主要市場收盤資料，上漲 {up_count} 個、"
        f"下跌 {down_count} 個，平均漲跌 {avg_move:+.2f}%。{tone}。"
        f"表現最強為{strongest['name']}（{strongest['daily_pct']:+.2f}%），"
        f"最弱為{weakest['name']}（{weakest['daily_pct']:+.2f}%）。"
        "此為收盤價與新聞的方向性整理，不等同投資建議。"
    )


def fetch_taiwan_vix():
    """取得期交所近月臺灣 VIX 收盤值，並產生可解釋的回檔風險分級。"""
    observations = []
    now = datetime.now()
    month_cursor = now.year * 12 + now.month - 1
    for offset in range(3):
        value = month_cursor - offset
        year, month_index = divmod(value, 12)
        month = month_index + 1
        url = (
            "https://www.taifex.com.tw/file/taifex/Dailydownload/vix/"
            f"log2data/{year}{month:02d}new.txt"
        )
        try:
            response = requests.get(url, headers=HEADERS, timeout=20)
            response.raise_for_status()
            content = response.content.decode("big5", errors="replace")
            for line in content.splitlines():
                fields = line.split()
                if len(fields) < 4 or not fields[0].isdigit() or len(fields[0]) != 8:
                    continue
                observations.append({
                    "date": datetime.strptime(fields[0], "%Y%m%d").strftime("%Y-%m-%d"),
                    "value": float(fields[-1]),
                })
        except Exception as exc:
            log(f"臺灣 VIX {year}-{month:02d} 取得失敗：{exc}")

    by_date = {item["date"]: item for item in observations}
    series = sorted(by_date.values(), key=lambda item: item["date"])
    if not series:
        return None

    latest = series[-1]
    previous = series[-2] if len(series) >= 2 else None
    five_days_ago = series[-6] if len(series) >= 6 else None
    daily_change = ((latest["value"] / previous["value"] - 1) * 100
                    if previous and previous["value"] else None)
    five_day_change = ((latest["value"] / five_days_ago["value"] - 1) * 100
                       if five_days_ago and five_days_ago["value"] else None)
    recent_values = [item["value"] for item in series[-20:]]
    percentile = (sum(item <= latest["value"] for item in recent_values)
                  / len(recent_values) * 100 if recent_values else None)

    level = latest["value"]
    if level >= 30:
        risk, action = "極高", "優先降槓桿、縮小持股並等待波動回落"
    elif level >= 25:
        risk, action = "高", "避免追高，分批減碼高波動持股並收緊停損"
    elif level >= 20:
        risk, action = "偏高", "保留現金、控制部位，僅分批承接強勢股"
    elif level >= 15:
        risk, action = "中性", "依趨勢持有，但仍需遵守個股停損"
    else:
        risk, action = "偏低", "市場情緒穩定，但低波動不代表不會突然回檔"
    if daily_change is not None and daily_change >= 10:
        action += "；VIX 單日急升，短線風險正在加速"

    return {
        "date": latest["date"], "value": latest["value"],
        "daily_change_pct": daily_change, "five_day_change_pct": five_day_change,
        "twenty_day_percentile": percentile, "risk": risk, "action": action,
        "source": "臺灣期貨交易所",
        "source_url": "https://www.taifex.com.tw/cht/7/vixDaily3MNew",
    }


def fetch_related_supply_chain_moves():
    """由證交所與櫃買中心官方 API 批次取得最近完整交易日漲跌。"""
    def numeric(value):
        cleaned = re.sub(r"[^0-9.\-]", "", str(value or ""))
        return float(cleaned) if re.search(r"\d", cleaned) else None

    def roc_compact_date(value):
        text = str(value or "")
        if len(text) == 7 and text.isdigit():
            return f"{int(text[:3]) + 1911:04d}-{text[3:5]}-{text[5:7]}"
        return "—"

    by_code = {}
    official_sources = (
        ("https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL",
         "Code", "Date", "ClosingPrice", "Change"),
        ("https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes",
         "SecuritiesCompanyCode", "Date", "Close", "Change"),
    )
    for url, code_field, date_field, close_field, change_field in official_sources:
        try:
            response = requests.get(url, headers=HEADERS, timeout=30)
            response.raise_for_status()
            for item in response.json():
                code = str(item.get(code_field, "")).strip()
                close = numeric(item.get(close_field))
                change = numeric(item.get(change_field))
                previous_close = close - change if close is not None and change is not None else None
                if code and close is not None and previous_close and previous_close > 0:
                    by_code[code] = {
                        "date": roc_compact_date(item.get(date_field)),
                        "close": close,
                        "daily_pct": change / previous_close * 100,
                    }
        except Exception as exc:
            log(f"官方供應鏈行情取得失敗（{url}）：{exc}")

    return {
        name: {"name": name, "symbol": symbol, **by_code[symbol.split('.')[0]]}
        for name, symbol in RELATED_TICKERS.items() if symbol.split(".")[0] in by_code
    }


def fetch_stock_watchlist():
    """以證交所正式資料計算量價、換手率及法人籌碼觀察。"""

    def number(value):
        cleaned = re.sub(r"[^0-9.\-]", "", str(value or ""))
        return float(cleaned) if cleaned not in {"", "-", "."} else 0.0

    def roc_date_to_iso(value):
        year, month, day = str(value).split("/")
        return f"{int(year) + 1911:04d}-{int(month):02d}-{int(day):02d}"

    def month_keys(count=3):
        year = datetime.now().year
        month = datetime.now().month
        keys = []
        for _ in range(count):
            keys.append(f"{year:04d}{month:02d}01")
            month -= 1
            if month == 0:
                year -= 1
                month = 12
        return keys

    def stock_history(code):
        observations = []
        for date_key in month_keys():
            response = requests.get(
                "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY",
                params={"date": date_key, "stockNo": code, "response": "json"},
                headers=HEADERS,
                timeout=20,
            )
            response.raise_for_status()
            payload = response.json()
            if payload.get("stat") != "OK":
                continue
            fields = payload.get("fields", [])
            index = {field: position for position, field in enumerate(fields)}
            for item in payload.get("data", []):
                observations.append({
                    "date": roc_date_to_iso(item[index["日期"]]),
                    "open": number(item[index["開盤價"]]),
                    "high": number(item[index["最高價"]]),
                    "low": number(item[index["最低價"]]),
                    "close": number(item[index["收盤價"]]),
                    "volume": number(item[index["成交股數"]]),
                })
            time.sleep(0.1)
        return sorted(
            {item["date"]: item for item in observations}.values(),
            key=lambda item: item["date"],
        )

    def issued_share_map():
        response = requests.get(
            "https://openapi.twse.com.tw/v1/opendata/t187ap03_L",
            headers=HEADERS,
            timeout=20,
        )
        response.raise_for_status()
        result = {}
        for company in response.json():
            code = str(company.get("公司代號", "")).strip()
            capital = number(company.get("實收資本額"))
            par_match = re.search(r"([0-9.]+)\s*元", str(company.get("普通股每股面額", "")))
            par_value = float(par_match.group(1)) if par_match else 10.0
            preferred_shares = number(company.get("特別股"))
            if code and capital > 0 and par_value > 0:
                result[code] = max(capital / par_value - preferred_shares, 0)
        return result

    names = {symbol.split(".")[0]: name for name, symbol in STOCK_WATCHLIST}
    related_market_moves = fetch_related_supply_chain_moves()
    histories = {}
    for code in names:
        try:
            histories[code] = stock_history(code)
        except Exception as exc:
            log(f"  {names[code]} 證交所行情取得失敗：{exc}")
            histories[code] = []

    try:
        share_counts = issued_share_map()
    except Exception as exc:
        log(f"  發行股數取得失敗：{exc}")
        share_counts = {}

    # 台股成本儀表板的六大外資分點逐日歷史。無資料時不使用主力成本冒充。
    foreign_history = {code: [] for code in names}
    try:
        response = requests.get(FOREIGN_BROKER_HISTORY_URL, headers=HEADERS, timeout=25)
        response.raise_for_status()
        for item in csv.DictReader(StringIO(response.text)):
            code = str(item.get("symbol", "")).replace(".0", "").zfill(4)
            if code in foreign_history:
                foreign_history[code].append(item)
    except Exception as exc:
        log(f"  六大外資分點歷史取得失敗：{exc}")

    def foreign_inventory_cost(code, history):
        raw = foreign_history.get(code, [])
        if not raw or not history:
            return None, 0, "六大外資分點歷史不足"
        available_dates = sorted({item.get("date", "") for item in raw if item.get("date")})[-30:]
        selected = [item for item in raw if item.get("date") in available_dates]
        prices = {
            item["date"]: (item["high"] + item["low"] + item["close"]) / 3
            for item in history
            if item["high"] > 0 and item["low"] > 0 and item["close"] > 0
        }
        by_broker = {}
        for item in selected:
            date_text = item.get("date", "")
            if date_text not in prices:
                continue
            broker = item.get("broker", "未命名分點")
            by_broker.setdefault(broker, []).append((date_text, item, prices[date_text]))

        inventories = []
        used_dates = set()
        for broker_rows in by_broker.values():
            inventory = 0.0
            average_cost = None
            for date_text, item, price in sorted(broker_rows, key=lambda value: value[0]):
                buy = max(number(item.get("buy_lots")), 0)
                sell = max(number(item.get("sell_lots")), 0)
                used_dates.add(date_text)
                if buy > 0:
                    total_cost = inventory * (average_cost or 0) + buy * price
                    inventory += buy
                    average_cost = total_cost / inventory
                if sell > 0:
                    if sell >= inventory:
                        inventory = 0.0
                        average_cost = None
                    else:
                        inventory -= sell
            if inventory > 0 and average_cost is not None:
                inventories.append((inventory, average_cost))
        total_inventory = sum(item[0] for item in inventories)
        if total_inventory <= 0:
            return None, len(used_dates), "區間無可追蹤剩餘庫存"
        cost = sum(quantity * price for quantity, price in inventories) / total_inventory
        return cost, len(used_dates), "六大外資分點逐日移動平均估算"

    def kd_values(history):
        k_value = None
        d_value = None
        for position in range(8, len(history)):
            window = history[position - 8:position + 1]
            high9 = max(item["high"] for item in window)
            low9 = min(item["low"] for item in window)
            rsv = ((history[position]["close"] - low9) / (high9 - low9) * 100) if high9 > low9 else 50.0
            k_value = rsv if k_value is None else (2 * k_value + rsv) / 3
            d_value = k_value if d_value is None else (2 * d_value + k_value) / 3
        return k_value, d_value

    def recent_news(name, code, limit=3):
        """由 Google 新聞 RSS 取得近7日標題、來源、日期與連結。"""
        try:
            response = requests.get(
                "https://news.google.com/rss/search",
                params={
                    "q": f"{name} {code} 股票 新聞 when:7d",
                    "hl": "zh-TW",
                    "gl": "TW",
                    "ceid": "TW:zh-Hant",
                },
                headers=HEADERS,
                timeout=20,
            )
            response.raise_for_status()
            soup = BeautifulSoup(response.content, "xml")
            results = []
            seen = set()
            non_news_phrases = (
                "個股概覽", "今日股價與討論", "法人籌碼變化", "買賣超總表",
                "股市爆料同學會", "發放緯穎", "歷史股價", "基本資料",
            )
            for item in soup.find_all("item"):
                title = " ".join((item.title.get_text(" ", strip=True) if item.title else "").split())
                link = item.link.get_text(strip=True) if item.link else ""
                source_tag = item.find("source")
                source = source_tag.get_text(" ", strip=True) if source_tag else "Google 新聞"
                if not title or not link or title in seen:
                    continue
                # 搜尋結果偶爾混入同代號或泛產業文章；至少須出現名稱或代號。
                if name not in title and code not in title:
                    continue
                if any(phrase in title for phrase in non_news_phrases):
                    continue
                published = "—"
                if item.pubDate:
                    try:
                        published = parsedate_to_datetime(item.pubDate.get_text(strip=True)).astimezone().strftime("%Y-%m-%d")
                    except Exception:
                        pass
                seen.add(title)
                results.append({
                    "stock": name,
                    "date": published,
                    "source": source,
                    "title": title,
                    "url": link,
                })
                if len(results) >= limit:
                    break
            return results
        except Exception as exc:
            log(f"  {name} 近期新聞取得失敗：{exc}")
            return []

    trading_dates = sorted(
        {item["date"] for history in histories.values() for item in history},
        reverse=True,
    )[:5]
    institutional = {
        code: {"foreign": 0.0, "domestic": 0.0, "total": 0.0}
        for code in names
    }
    latest_institutional = {
        code: {"foreign": None, "domestic": None, "date": "—"}
        for code in names
    }
    for date_text in trading_dates:
        try:
            response = requests.get(
                "https://www.twse.com.tw/rwd/zh/fund/T86",
                params={
                    "date": date_text.replace("-", ""),
                    "selectType": "ALLBUT0999",
                    "response": "json",
                },
                headers=HEADERS,
                timeout=20,
            )
            response.raise_for_status()
            payload = response.json()
            if payload.get("stat") != "OK":
                continue
            fields = payload.get("fields", [])
            index = {field: position for position, field in enumerate(fields)}
            foreign_field = next(
                field for field in fields
                if field.startswith("外陸資買賣超股數") and "不含外資自營商" in field
            )
            for item in payload.get("data", []):
                code = str(item[index["證券代號"]]).strip()
                if code not in institutional:
                    continue
                foreign = number(item[index[foreign_field]])
                trust = number(item[index["投信買賣超股數"]])
                dealer = number(item[index["自營商買賣超股數"]])
                domestic = trust + dealer
                institutional[code]["foreign"] += foreign
                institutional[code]["domestic"] += domestic
                institutional[code]["total"] += foreign + domestic
                if latest_institutional[code]["foreign"] is None:
                    latest_institutional[code] = {
                        "foreign": foreign,
                        "domestic": domestic,
                        "date": date_text,
                    }
            time.sleep(0.1)
        except Exception as exc:
            log(f"  {date_text} 法人資料取得失敗：{exc}")

    rows = []
    for code, name in names.items():
        history = histories.get(code, [])
        if len(history) < 20:
            rows.append({
                "name": name, "symbol": f"{code}.TW", "date": "—", "close": None,
                "daily_pct": None, "five_day_pct": None, "ma5": None, "ma20": None,
                "turnover_pct": None, "volume_change_pct": None, "volume_ratio": None,
                "price_volume": "資料待更新", "chip_concentration_pct": None,
                "foreign_net_lots": None, "domestic_net_lots": None,
                "institutional_date": "—", "action": "資料待更新",
                "reason": "證交所有效日線不足 20 筆", "foreign_cost": None,
                "price_vs_foreign_cost_pct": None, "foreign_cost_days": 0,
                "foreign_cost_source": "資料不足", "entry_score": None,
                "entry_assessment": "資料不足", "kd_k": None, "kd_d": None,
                "supply_chain": "未分類", "related_holdings": "—",
                "related_moves": [],
                "recent_news": recent_news(name, code),
            })
            continue

        latest = history[-1]
        previous = history[-2]
        close = latest["close"]
        daily_pct = (close / previous["close"] - 1) * 100
        five_day_pct = (close / history[-6]["close"] - 1) * 100
        ma5 = sum(item["close"] for item in history[-5:]) / 5
        ma20 = sum(item["close"] for item in history[-20:]) / 20
        volume_change_pct = (
            (latest["volume"] / previous["volume"] - 1) * 100
            if previous["volume"] else None
        )
        average_volume_20 = sum(item["volume"] for item in history[-20:]) / 20
        volume_ratio = latest["volume"] / average_volume_20 if average_volume_20 else None
        shares = share_counts.get(code)
        turnover_pct = latest["volume"] / shares * 100 if shares else None
        five_day_volume = sum(item["volume"] for item in history[-5:])
        chip_concentration = (
            institutional[code]["total"] / five_day_volume * 100
            if five_day_volume else None
        )
        latest_chip = latest_institutional[code]
        foreign_cost, foreign_cost_days, foreign_cost_source = foreign_inventory_cost(code, history)
        price_vs_foreign_cost_pct = (
            (close / foreign_cost - 1) * 100 if foreign_cost and foreign_cost > 0 else None
        )
        k_value, d_value = kd_values(history)
        raw_entry_score = (
            (50 if close > ma20 else -50)
            + (30 if k_value is not None and d_value is not None and k_value > d_value else -30)
        )
        entry_score = max(0, min(100, raw_entry_score + 40))
        entry_assessment = "分批進場觀察" if raw_entry_score > 0 else "不建議進場／觀望"
        chain = SUPPLY_CHAIN_HOLDINGS.get(code, {"sectors": [], "related": []})
        supply_chain = "、".join(chain["sectors"]) if chain["sectors"] else "現有供應鏈清單未分類"
        related_holdings = "、".join(chain["related"]) if chain["related"] else "—"
        related_moves = [
            related_market_moves[related_name]
            for related_name in chain["related"] if related_name in related_market_moves
        ]

        price_direction = "價漲" if daily_pct > 0 else "價跌" if daily_pct < 0 else "價平"
        volume_direction = (
            "量增" if volume_change_pct is not None and volume_change_pct > 0
            else "量縮" if volume_change_pct is not None and volume_change_pct < 0
            else "量平"
        )
        price_volume = f"{price_direction}{daily_pct:+.2f}%／{volume_direction}{abs(volume_change_pct or 0):.1f}%"

        if close >= ma20 and ma5 >= ma20 and (chip_concentration or 0) >= 0:
            action = "偏多續抱"
            reason = "站上20日線，且五日法人集中度未轉負"
        elif close >= ma20:
            action = "續抱觀察"
            reason = "仍守20日線，但法人籌碼偏保守"
        elif close >= ma20 * 0.97:
            action = "觀望／守支撐"
            reason = "位於20日線下方3%內，等待價量與法人轉強"
        else:
            action = "減碼觀察"
            reason = "跌破20日線逾3%，短線風險升高"

        rows.append({
            "name": name, "symbol": f"{code}.TW", "date": latest["date"],
            "close": close, "daily_pct": daily_pct, "five_day_pct": five_day_pct,
            "ma5": ma5, "ma20": ma20, "turnover_pct": turnover_pct,
            "volume_change_pct": volume_change_pct, "volume_ratio": volume_ratio,
            "price_volume": price_volume,
            "chip_concentration_pct": chip_concentration,
            "foreign_net_lots": (
                latest_chip["foreign"] / 1000 if latest_chip["foreign"] is not None else None
            ),
            "domestic_net_lots": (
                latest_chip["domestic"] / 1000 if latest_chip["domestic"] is not None else None
            ),
            "institutional_date": latest_chip["date"], "action": action, "reason": reason,
            "foreign_cost": foreign_cost,
            "price_vs_foreign_cost_pct": price_vs_foreign_cost_pct,
            "foreign_cost_days": foreign_cost_days,
            "foreign_cost_source": foreign_cost_source,
            "entry_score": entry_score,
            "entry_assessment": entry_assessment,
            "kd_k": k_value,
            "kd_d": d_value,
            "supply_chain": supply_chain,
            "related_holdings": related_holdings,
            "related_moves": related_moves,
            "recent_news": recent_news(name, code),
        })
    return rows


def build_line_message(
    market_rows,
    market_analysis,
    selected_articles,
    udn_hot_news=None,
    stock_rows=None,
    taiwan_vix=None,
):
    """建立 3 分鐘看盤日記；僅列日期，不顯示星期或 ETF 代理價。"""
    today_label = datetime.now().strftime("%m/%d").lstrip("0").replace("/0", "/")
    us_names = {"美國標普500", "美國費城半導體", "美國那斯達克", "美國道瓊"}
    us_rows = [row for row in market_rows if row["name"] in us_names]
    asia_rows = [row for row in market_rows if row["name"] not in us_names and "歐洲" not in row["name"]]
    europe_rows = [row for row in market_rows if "歐洲" in row["name"]]

    def quote_line(row):
        if row["close"] is None:
            return f"🔸️{row['name']}：—（{row['date']}，指數資料待更新）"
        icon = "📈" if row["daily_pct"] > 0 else "📉" if row["daily_pct"] < 0 else "➖"
        return f"🔸️{row['name']}：{icon}{abs(row['daily_pct']):.2f}%，收{row['close']:,.2f}（{row['date']}）"

    def chip_line(row):
        if row["daily_pct"] is None:
            return f"🔹️{row['name']}：—"
        if row["daily_pct"] >= 1:
            text = "買盤動能偏強"
        elif row["daily_pct"] > 0:
            text = "買盤略占優勢"
        elif row["daily_pct"] <= -1:
            text = "賣壓偏重"
        else:
            text = "多空拉鋸"
        return f"🔹️{row['name']}：{text}（價格動能代理 {row['daily_pct']:+.2f}%）"

    lines = [
        f"🌐〈{today_label}〉3分鐘看盤日記",
        "",
        "⭕️《重點摘要》",
        market_analysis,
        "",
        "🎯《美股盤後報價》",
    ]
    lines.extend(quote_line(row) for row in us_rows)
    lines.extend(["", "⭕️《亞洲股市》"])
    lines.extend(quote_line(row) for row in asia_rows)
    if europe_rows:
        lines.extend(["", "🎯《歐洲股市》"])
        lines.extend(quote_line(row) for row in europe_rows)

    lines.extend(["", "🎯《前一交易日籌碼／動能觀察》"])
    lines.extend(chip_line(row) for row in us_rows + asia_rows)
    lines.append("註：無跨市場一致的即時法人籌碼時，以實際指數日漲跌作價格動能代理，不等同法人買賣超。")

    if taiwan_vix:
        daily = "—" if taiwan_vix["daily_change_pct"] is None else f"{taiwan_vix['daily_change_pct']:+.2f}%"
        five_day = "—" if taiwan_vix["five_day_change_pct"] is None else f"{taiwan_vix['five_day_change_pct']:+.2f}%"
        percentile = "—" if taiwan_vix["twenty_day_percentile"] is None else f"{taiwan_vix['twenty_day_percentile']:.0f}%"
        lines.extend([
            "", "⚠️《臺股 VIX 回檔風險》",
            f"臺灣 VIX {taiwan_vix['value']:.2f}（{taiwan_vix['date']}）｜日變化{daily}｜5日變化{five_day}",
            f"近20日分位{percentile}｜回檔風險：{taiwan_vix['risk']}",
            f"判讀：{taiwan_vix['action']}",
            "註：VIX衡量未來30天預期波動，不等於股市必然下跌或精確跌幅。",
        ])
    else:
        lines.extend(["", "⚠️《臺股 VIX 回檔風險》", "期交所資料暫時無法取得，本次不推測風險等級。"])

    if stock_rows:
        lines.extend(["", "📌《台股持有觀察》"])
        for row in stock_rows:
            if row["close"] is None:
                lines.append(f"▪️{row['name']}：資料待更新")
                continue
            turnover = "—" if row["turnover_pct"] is None else f"{row['turnover_pct']:.2f}%"
            concentration = (
                "—" if row["chip_concentration_pct"] is None
                else f"{row['chip_concentration_pct']:+.2f}%"
            )
            foreign = "—" if row["foreign_net_lots"] is None else f"{row['foreign_net_lots']:+,.0f}張"
            domestic = "—" if row["domestic_net_lots"] is None else f"{row['domestic_net_lots']:+,.0f}張"
            lines.append(f"▪️{row['name']}：{row['close']:,.2f}｜{row['action']}")
            lines.append(
                f"  {row['price_volume']}；換手{turnover}；20日量比{row['volume_ratio']:.2f}倍"
            )
            lines.append(
                f"  籌碼集中度{concentration}；外資{foreign}／內資{domestic}（{row['institutional_date']}）"
            )
            foreign_cost = (
                "資料不足" if row["foreign_cost"] is None
                else f"{row['foreign_cost']:,.2f}（現價乖離{row['price_vs_foreign_cost_pct']:+.2f}%）"
            )
            lines.append(
                f"  六大外資成本{foreign_cost}；進場評估{row['entry_score'] if row['entry_score'] is not None else '—'}分／{row['entry_assessment']}"
            )
            lines.append(
                f"  供應鏈：{row['supply_chain']}；相關持股：{row['related_holdings']}"
            )
            related_moves = sorted(
                row.get("related_moves") or [], key=lambda item: item["daily_pct"], reverse=True
            )
            if related_moves:
                shown_moves = related_moves if len(related_moves) <= 3 else [related_moves[0], related_moves[-1]]
                move_text = "、".join(
                    f"{item['name']}{item['daily_pct']:+.2f}%" for item in shown_moves
                )
                lines.append(f"  相關供應鏈前一日：{move_text}（{related_moves[0]['date']}）")
            else:
                lines.append("  相關供應鏈前一日：行情資料不足")
            news_items = row.get("recent_news") or []
            if news_items:
                news = news_items[0]
                title = news["title"]
                lines.append(f"  近期新聞（{news['date']}／{news['source']}）：{title}")
                link_label, article_url = article_link(news)
                lines.append("點開看完整新聞" if link_label == "新聞原文" else "原文暫缺，點開搜尋新聞")
                lines.append(article_url)
            else:
                lines.append("  近期新聞：近7日無可驗證結果")
        lines.append("註：籌碼集中度＝近5日三大法人淨買賣超÷近5日成交量；內資＝投信＋自營商。")
        lines.append("外資成本為六大外資分點可追蹤剩餘庫存的移動平均估算，不是外資真實帳簿成本。")
        lines.append("進場評估沿用價格相對MA20與KD方向；供應鏈為研究分類，不代表直接供貨關係。")
        lines.append("個股判讀依證交所收盤、量價與法人資料產生；尚未結算時沿用最近完整交易日。")

    lines.extend(["", "🎯《市場焦點新聞》"])
    # 早報先列美股／美債焦點，其餘市場再按發布時間排列。
    latest_articles = sorted(
        selected_articles.values(),
        key=lambda article: (
            0 if article.get("region") == "美國" else 1,
            -article["pub_date"].timestamp(),
        ),
    )[:3 if stock_rows else 6]
    for article in latest_articles:
        lines.append(f"♦️{article['title']}")
        if article.get("url"):
            link_label, article_url = article_link(article)
            lines.append("點開看完整新聞" if link_label == "新聞原文" else "原文暫缺，點開搜尋新聞")
            lines.append(article_url)
    if udn_hot_news:
        lines.extend(["", "🔥《經濟日報熱門財經新聞》"])
        for article in udn_hot_news:
            lines.append(f"🔹️{article['title']}")
            link_label, article_url = article_link(article)
            lines.append("點開看完整新聞" if link_label == "新聞原文" else "原文暫缺，點開搜尋新聞")
            lines.append(article_url)
    lines.extend(["", "🔎《行情查核》"])
    for label, url in MARKET_REFERENCE_LINKS:
        lines.append(f"・{label}：{url}")
    lines.append("註：不同市場休市日不同，各列日期以最近已完成交易日為準；不使用 ETF 價格冒充指數。")
    lines.extend(["", "🌈免責聲明：依公開網頁資訊彙整，僅供參考，不構成任何投資建議。"])
    return "\n".join(lines)


def load_line_config():
    """由環境變數或 line_config.json 讀取 LINE 憑證，絕不寫入報告。"""
    config = {}
    if os.path.exists(LINE_CONFIG_PATH):
        try:
            with open(LINE_CONFIG_PATH, "r", encoding="utf-8") as config_file:
                config = json.load(config_file)
        except Exception as exc:
            log(f"LINE 設定檔無法讀取：{exc}")
    return {
        "enabled": bool(config.get("enabled", False)),
        "delivery_mode": config.get("delivery_mode", "push"),
        "channel_access_token": os.getenv("LINE_CHANNEL_ACCESS_TOKEN") or config.get("channel_access_token", ""),
        "target_id": os.getenv("LINE_TARGET_ID") or config.get("target_id", ""),
    }


def _build_stock_news_flex(stock_rows):
    """建立 LINE 原生可點擊新聞按鈕，避免 Google News RSS 長轉址顯示成亂碼。"""
    bubbles = []
    for row in stock_rows or []:
        news_items = row.get("recent_news") or []
        if not news_items:
            continue
        news = news_items[0]
        url = str(news.get("url") or "").strip()
        if not url.startswith(("https://", "http://")) or len(url) > 1000:
            continue
        title = " ".join(str(news.get("title") or "近期新聞").split())
        bubbles.append({
            "type": "bubble",
            "size": "kilo",
            "body": {
                "type": "box",
                "layout": "vertical",
                "spacing": "sm",
                "contents": [
                    {"type": "text", "text": row.get("name", "個股"), "weight": "bold", "size": "lg"},
                    {"type": "text", "text": title, "wrap": True, "size": "sm", "color": "#444444"},
                    {
                        "type": "text",
                        "text": f"{news.get('date', '—')} · {news.get('source', 'Google 新聞')}",
                        "size": "xs",
                        "color": "#888888",
                        "margin": "md",
                    },
                ],
            },
            "footer": {
                "type": "box",
                "layout": "vertical",
                "contents": [{
                    "type": "button",
                    "style": "primary",
                    "color": "#168C62",
                    "action": {"type": "uri", "label": "閱讀新聞", "uri": url},
                }],
            },
        })
    if not bubbles:
        return None
    return {
        "type": "flex",
        "altText": "六檔個股近期新聞（點擊閱讀）",
        "contents": {"type": "carousel", "contents": bubbles[:10]},
    }


def send_line_report(message, stock_rows=None):
    """透過 LINE Messaging API 推送；未設定時安全略過。"""
    config = load_line_config()
    if not config["enabled"]:
        log("LINE 推播未啟用；已產生報告但不傳送。")
        return False
    if not config["channel_access_token"]:
        log("LINE 推播已啟用，但缺少 channel_access_token。")
        return False
    if config["delivery_mode"] == "push" and not config["target_id"]:
        log("LINE 指定對象推播缺少 target_id。")
        return False
    try:
        is_broadcast = config["delivery_mode"] == "broadcast"
        endpoint = (
            "https://api.line.me/v2/bot/message/broadcast"
            if is_broadcast else
            "https://api.line.me/v2/bot/message/push"
        )
        messages = [{"type": "text", "text": message}]
        news_flex = _build_stock_news_flex(stock_rows)
        if news_flex:
            messages.append(news_flex)
        payload = {"messages": messages}
        if not is_broadcast:
            payload["to"] = config["target_id"]
        response = requests.post(
            endpoint,
            headers={
                "Authorization": f"Bearer {config['channel_access_token']}",
                "Content-Type": "application/json",
            },
            json=payload,
            timeout=20,
        )
        response.raise_for_status()
        log("LINE 市場綜合分析報告已傳送。")
        return True
    except Exception as exc:
        log(f"LINE 推播失敗：{exc}")
        return False


# ============================================================================
# 模組二：內容組織（保留原文完整內容）
# ============================================================================

def organize_for_word(selected_articles):
    """將選取的文章組織為 Word 渲染所需的資料結構。"""
    regions = {}

    for cat, article in selected_articles.items():
        region = article["region"]
        if region not in regions:
            regions[region] = {
                "region_label": region,
                "articles": [],
            }

        regions[region]["articles"].append({
            "title": article["title"],
            "date": article["pub_date"].strftime("%Y-%m-%d %H:%M:%S"),
            "paragraphs": article["body_lines"] or ["（內文待更新）"],
        })

    return regions


# ============================================================================
# 模組三：Word 底層樣式建設 (Style Initialization)
# ============================================================================

def _set_font_xml(rpr_element, font_en="Times New Roman", font_ea="標楷體"):
    """透過 XML 鎖死中英雙軌字體。"""
    rfonts = rpr_element.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = parse_xml(
            f'<w:rFonts {nsdecls("w")} '
            f'w:ascii="{font_en}" w:hAnsi="{font_en}" '
            f'w:cs="{font_en}" w:eastAsia="{font_ea}"/>'
        )
        rpr_element.insert(0, rfonts)
    else:
        rfonts.set(qn("w:ascii"), font_en)
        rfonts.set(qn("w:hAnsi"), font_en)
        rfonts.set(qn("w:cs"), font_en)
        rfonts.set(qn("w:eastAsia"), font_ea)


def _ensure_char_style(doc, name, color, bold, size_pt):
    """建立或取得 Character Style。"""
    try:
        style = doc.styles[name]
    except KeyError:
        style = doc.styles.add_style(name, 2)  # CHARACTER
        style.base_style = doc.styles["Default Paragraph Font"]

    style.font.size = Pt(size_pt)
    style.font.bold = bold
    style.font.color.rgb = color

    rpr = style.element.find(qn("w:rPr"))
    if rpr is None:
        rpr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
        style.element.append(rpr)
    _set_font_xml(rpr)
    return style


def init_styles(doc):
    """初始化所有樣式。"""
    _ensure_char_style(doc, "NewsTitle", RGBColor(0, 0, 0xFF), True, 14)
    _ensure_char_style(doc, "NewsBody", RGBColor(0, 0, 0), False, 12)

    normal = doc.styles["Normal"]
    normal.font.size = Pt(12)
    rpr = normal.element.find(qn("w:rPr"))
    if rpr is None:
        rpr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
        normal.element.append(rpr)
    _set_font_xml(rpr)


# ============================================================================
# 模組四：表格佈局與內容渲染 (Table Layout & Rendering)
# ============================================================================

def _zero_spacing(p):
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0


def _set_cell_width(cell, cm):
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="0" w:type="dxa"/>')
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(int(cm * 567)))
    tcW.set(qn("w:type"), "dxa")


def _set_cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(
        f'<w:shd {nsdecls("w")} w:fill="{hex_color}" w:val="clear"/>'
    ))


def _add_styled_run(p, text, style_name, doc):
    run = p.add_run(text)
    try:
        run.style = doc.styles[style_name]
    except KeyError:
        pass
    rpr = run._element.find(qn("w:rPr"))
    if rpr is None:
        rpr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
        run._element.insert(0, rpr)
    _set_font_xml(rpr)
    return run


def build_market_overview(doc, market_rows, market_analysis):
    """在新聞前加入每日主要市場收盤表與綜合判讀。"""
    heading = doc.add_paragraph()
    heading.paragraph_format.space_after = Pt(6)
    heading_run = _add_styled_run(heading, "每日主要市場收盤狀況", "NewsTitle", doc)
    heading_run.font.size = Pt(16)

    table = doc.add_table(rows=1, cols=6)
    table.style = doc.styles["Table Grid"]
    headers = ("市場", "日期", "收盤", "日漲跌", "近5日", "狀態")
    for index, label in enumerate(headers):
        cell = table.rows[0].cells[index]
        _set_cell_bg(cell, "92D050")
        paragraph = cell.paragraphs[0]
        _zero_spacing(paragraph)
        run = _add_styled_run(paragraph, label, "NewsBody", doc)
        run.bold = True

    for row_data in market_rows:
        row = table.add_row().cells
        values = (
            row_data["name"],
            row_data["date"],
            "—" if row_data["close"] is None else f"{row_data['close']:,.2f}",
            "—" if row_data["daily_pct"] is None else f"{row_data['daily_pct']:+.2f}%",
            "—" if row_data["weekly_pct"] is None else f"{row_data['weekly_pct']:+.2f}%",
            row_data["status"],
        )
        for index, value in enumerate(values):
            paragraph = row[index].paragraphs[0]
            _zero_spacing(paragraph)
            run = _add_styled_run(paragraph, value, "NewsBody", doc)
            if index in (3, 4) and value != "—":
                run.font.color.rgb = RGBColor(0xC0, 0, 0) if value.startswith("+") else RGBColor(0, 0x80, 0)

    source_note = doc.add_paragraph()
    source_note.paragraph_format.space_before = Pt(3)
    source_note.paragraph_format.space_after = Pt(3)
    _add_styled_run(
        source_note,
        "資料說明：只呈現實際市場指數；來源暫時受限時顯示待更新，不以 ETF 價格冒充指數。",
        "NewsBody",
        doc,
    )

    analysis_heading = doc.add_paragraph()
    analysis_heading.paragraph_format.space_before = Pt(8)
    analysis_heading.paragraph_format.space_after = Pt(2)
    run = _add_styled_run(analysis_heading, "市場綜合判讀", "NewsBody", doc)
    run.bold = True
    analysis_paragraph = doc.add_paragraph()
    analysis_paragraph.paragraph_format.space_after = Pt(10)
    _add_styled_run(analysis_paragraph, market_analysis, "NewsBody", doc)


def build_tables(doc, regions_data):
    """以單一表格渲染四大地區及其新聞。"""
    table = doc.add_table(rows=1, cols=2)
    table.style = doc.styles["Table Grid"]
    table.autofit = False

    tblPr = table._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>')
        table._tbl.insert(0, tblPr)
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    tw = tblPr.find(qn("w:tblW"))
    if tw is None:
        tw = parse_xml(f'<w:tblW {nsdecls("w")} w:w="0" w:type="dxa"/>')
        tblPr.append(tw)
    tw.set(qn("w:w"), str(int(18.6 * 567)))
    tw.set(qn("w:type"), "dxa")

    old_grid = table._tbl.find(qn("w:tblGrid"))
    if old_grid is not None:
        table._tbl.remove(old_grid)
    grid = parse_xml(
        f'<w:tblGrid {nsdecls("w")}>'
        f'<w:gridCol w:w="{int(1.6 * 567)}"/>'
        f'<w:gridCol w:w="{int(17 * 567)}"/>'
        f'</w:tblGrid>'
    )
    tblPr.addnext(grid)

    header = table.rows[0]
    trPr = header._tr.get_or_add_trPr()
    trPr.append(parse_xml(
        f'<w:tblHeader {nsdecls("w")} w:val="true"/>'
    ))
    for index, (label, width) in enumerate((("分類", 1.6), ("盤後新聞", 17))):
        cell = header.cells[index]
        _set_cell_bg(cell, "92D050")
        _set_cell_width(cell, width)
        p = cell.paragraphs[0]
        p.style = doc.styles["Normal"]
        _zero_spacing(p)
        run = _add_styled_run(p, label, "NewsBody", doc)
        run.bold = True

    for region_key in REGION_ORDER:
        rd = regions_data.get(region_key)
        if not rd:
            continue

        row = table.add_row()
        left, right = row.cells
        _set_cell_width(left, 1.6)
        _set_cell_width(right, 17)

        left_p = left.paragraphs[0]
        left_p.style = doc.styles["Normal"]
        _zero_spacing(left_p)
        _add_styled_run(left_p, region_key, "NewsBody", doc)

        first_p = right.paragraphs[0]
        is_first = True
        for art in rd["articles"]:
            if not is_first:
                separator = right.add_paragraph()
                separator.style = doc.styles["Normal"]
                _zero_spacing(separator)

            title_p = first_p if is_first else right.add_paragraph()
            is_first = False
            title_p.style = doc.styles["Normal"]
            _zero_spacing(title_p)
            _add_styled_run(title_p, art["title"], "NewsTitle", doc)

            date_p = right.add_paragraph()
            date_p.style = doc.styles["Normal"]
            _zero_spacing(date_p)
            _add_styled_run(
                date_p,
                f'MoneyDJ新聞 {art["date"]} 發佈',
                "NewsBody",
                doc,
            )

            for para_text in art["paragraphs"]:
                body_p = right.add_paragraph()
                body_p.style = doc.styles["Normal"]
                _zero_spacing(body_p)
                _add_styled_run(body_p, para_text, "NewsBody", doc)


# ============================================================================
# 主程式
# ============================================================================

def main():
    log("=" * 60)
    log("MoneyDJ Word Formatter v1.2 — 全自動版")
    log("=" * 60)

    # 模組一：抓取新聞
    selected = scrape_all_news()

    if not selected:
        log("\n未抓取到任何符合條件的文章。")
        log("可能原因：網路連線問題或 MoneyDJ 頁面結構變更。")
        input("\n按 Enter 鍵結束...")
        sys.exit(1)

    # 每日主要市場收盤與跨市場判讀
    market_rows = fetch_global_market_closes()
    market_analysis = build_market_analysis(market_rows)
    stock_rows = fetch_stock_watchlist()
    udn_hot_news = fetch_udn_hot_news()

    # 模組二：組織內容
    log("\n" + "=" * 60)
    log("模組二：內容組織")
    log("=" * 60)
    regions_data = organize_for_word(selected)
    log(f"已組織 {len(regions_data)} 個區塊")

    # 模組三 + 四：產出 Word
    log("\n" + "=" * 60)
    log("模組三/四：Word 樣式建設 + 表格渲染")
    log("=" * 60)

    doc = Document()
    for sec in doc.sections:
        sec.left_margin = Cm(1.5)
        sec.right_margin = Cm(1.5)
        sec.top_margin = Cm(1.5)
        sec.bottom_margin = Cm(1.5)

    init_styles(doc)
    build_market_overview(doc, market_rows, market_analysis)
    build_tables(doc, regions_data)

    # 儲存
    today_str = datetime.now().strftime("%Y%m%d")
    output_path = os.path.join(OUTPUT_DIR, f"MoneyDJ_盤後新聞_{today_str}.docx")
    doc.save(output_path)

    log(f"\nWord 檔案已產出：{output_path}")

    line_message = build_line_message(
        market_rows,
        market_analysis,
        selected,
        udn_hot_news,
        stock_rows,
    )
    line_text_path = os.path.join(OUTPUT_DIR, f"LINE_市場綜合分析_{today_str}.txt")
    with open(line_text_path, "w", encoding="utf-8") as line_file:
        line_file.write(line_message)
    log(f"LINE 文字版已產出：{line_text_path}")
    log("日報已產生；請在管理頁預覽後手動發送。")
    log("=" * 60)

    return output_path


if __name__ == "__main__":
    output = main()
    # macOS：自動開啟檔案
    if sys.platform == "darwin" and "--no-open" not in sys.argv:
        os.system(f'open "{output}"')
    if "--no-pause" not in sys.argv:
        input("\n按 Enter 鍵結束...")
