#!/usr/bin/env python3
"""LINE 官方帳號市場分析手動發送頁面。"""

import os
import re
from datetime import datetime

import pandas as pd
import streamlit as st

import app as report
from subscriptions import recipients, send_one, send_subscribers


st.set_page_config(page_title="LINE 市場分析發送台", page_icon="📨", layout="wide")


def load_streamlit_secrets() -> None:
    """讓 Streamlit Cloud 可用 secrets；本機仍優先沿用安全設定檔。"""
    try:
        token = st.secrets.get("LINE_CHANNEL_ACCESS_TOKEN", "")
        target_id = st.secrets.get("LINE_TARGET_ID", "")
    except Exception:
        return
    if token:
        os.environ["LINE_CHANNEL_ACCESS_TOKEN"] = str(token)
        os.environ["LINE_CONFIG_SOURCE"] = "streamlit_secrets"
    if target_id:
        os.environ["LINE_TARGET_ID"] = str(target_id)
    if st.secrets.get("SUBSCRIPTIONS_DB"):
        os.environ["SUBSCRIPTIONS_DB"] = str(st.secrets["SUBSCRIPTIONS_DB"])


def require_cloud_password() -> None:
    """雲端 LINE 發送頁必須先通過密碼，避免公開網址遭他人濫用。"""
    try:
        app_password = str(st.secrets.get("APP_PASSWORD", "")) or os.getenv("APP_PASSWORD", "")
    except Exception:
        app_password = os.getenv("APP_PASSWORD", "")
    using_cloud_secrets = bool(os.getenv("LINE_CHANNEL_ACCESS_TOKEN"))
    if using_cloud_secrets and not app_password:
        st.error("雲端版尚未設定 APP_PASSWORD；為避免他人誤發 LINE，發送功能已鎖定。")
        st.stop()
    if not app_password or st.session_state.get("authenticated"):
        return

    st.title("LINE 市場分析發送台")
    st.caption("請先輸入管理密碼，登入後才可產生並發送 LINE 市場分析。")
    entered = st.text_input("管理密碼", type="password")
    if st.button("登入", type="primary"):
        if entered == app_password:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("密碼不正確。")
    st.stop()


@st.cache_data(ttl=900, max_entries=3, show_spinner=False)
def generate_report_payload(cache_key: str):
    """每十五分鐘最多重抓一次，手動更新時以 cache_key 強制刷新。"""
    del cache_key
    market_rows = report.fetch_global_market_closes()
    market_analysis = report.build_market_analysis(market_rows)
    stock_rows = report.fetch_stock_watchlist()
    taiwan_vix = report.fetch_taiwan_vix()
    selected_articles = report.scrape_all_news()
    hot_news = report.fetch_udn_hot_news()
    message = report.build_line_message(
        market_rows,
        market_analysis,
        selected_articles,
        hot_news,
        stock_rows,
        taiwan_vix,
    )
    return message, stock_rows, taiwan_vix


load_streamlit_secrets()
require_cloud_password()
st.title("LINE 市場分析發送台")
st.caption("追蹤：緯穎、南電、金像電、欣興、所羅門、晶豪科。先預覽、發給自己確認，再手動發給訂閱好友。")

config = report.load_line_config()
if config["enabled"] and config["channel_access_token"]:
    destination = "手動發送，依每日／每週訂閱名單"
    st.success(f"LINE 已連線；傳送方式：{destination}。憑證不會顯示在頁面或報告中。")
else:
    st.warning("LINE 尚未完成設定。請先雙擊「LINE官方串接設定.command」輸入 Channel access token。")

st.session_state.setdefault("line_preview", "")
st.session_state.setdefault("stock_rows", [])
st.session_state.setdefault("taiwan_vix", None)
st.session_state.setdefault("last_updated", "")

send_clicked = st.button(
    "產生預覽（不發送）",
    type="primary",

)

if send_clicked:
    refresh_key = datetime.now().isoformat(timespec="seconds")
    with st.spinner("正在更新市場、個股與新聞資料…"):
        message, stock_rows, taiwan_vix = generate_report_payload(refresh_key)
    st.session_state.line_preview = message
    st.session_state.stock_rows = stock_rows
    st.session_state.taiwan_vix = taiwan_vix
    st.session_state.last_updated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state.self_checked = ""
    st.success("預覽已更新，尚未發送。")

if st.session_state.line_preview:
    import hashlib
    digest = hashlib.sha256(st.session_state.line_preview.encode()).hexdigest()
    with st.expander("發送前預覽", expanded=True):
        st.text(st.session_state.line_preview)
    if st.button("發送給自己"):
        try:
            send_one(config["channel_access_token"], config["target_id"], st.session_state.line_preview)
            st.session_state.self_checked = digest
            st.success("已發給自己，確認後再發給好友。")
        except Exception:
            st.error("發送未確認，請核對自己的 User ID、憑證及 LINE 收件狀態。")
    frequency = st.selectbox("接收名單", ["daily", "weekly"], format_func=lambda x: {"daily":"每日訂閱", "weekly":"每週訂閱"}[x])
    st.caption("每週訂閱只是接收頻率；目前發送的是本次預覽，不會自動彙整一週內容。")
    ready = False
    try:
        count = len(recipients(frequency))
        st.caption(f"目前訂閱人數：{count}")
        ready = count > 0
    except (ValueError, OSError):
        st.info("好友訂閱尚未啟用：請完成 Webhook 與持久化資料庫設定。")
    confirmed = st.checkbox("我已檢查本次內容，確認發送", key="confirm_" + digest + frequency)
    if st.button("發送給訂閱好友", disabled=not (ready and confirmed and st.session_state.get("self_checked") == digest)):
        try:
            sent, skipped = send_subscribers(config["channel_access_token"], frequency, st.session_state.line_preview)
            st.success(f"已發送 {sent} 人；略過已處理或待查核紀錄 {skipped} 人。")
        except Exception:
            st.error("發送未全部確認，請查核紀錄；不要重複發送相同內容。")

if st.session_state.taiwan_vix:
    vix = st.session_state.taiwan_vix
    st.subheader("臺股 VIX 回檔風險")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("臺灣 VIX", f"{vix['value']:.2f}", help=f"資料日期：{vix['date']}")
    col2.metric("單日變化", "—" if vix["daily_change_pct"] is None else f"{vix['daily_change_pct']:+.2f}%")
    col3.metric("5日變化", "—" if vix["five_day_change_pct"] is None else f"{vix['five_day_change_pct']:+.2f}%")
    col4.metric("近20日分位", "—" if vix["twenty_day_percentile"] is None else f"{vix['twenty_day_percentile']:.0f}%")
    if vix["risk"] in {"高", "極高"}:
        st.error(f"回檔風險：{vix['risk']}｜{vix['action']}")
    elif vix["risk"] == "偏高":
        st.warning(f"回檔風險：{vix['risk']}｜{vix['action']}")
    else:
        st.info(f"回檔風險：{vix['risk']}｜{vix['action']}")
    st.caption("資料來源：臺灣期貨交易所；VIX衡量未來30天預期波動，不代表股市必然下跌或精確跌幅。")

if st.session_state.stock_rows:
    st.subheader("五檔個股最新判讀")
    table = pd.DataFrame(st.session_state.stock_rows)
    table = table[[
        "name", "date", "close", "turnover_pct", "price_volume", "volume_ratio",
        "chip_concentration_pct", "foreign_net_lots", "domestic_net_lots",
        "institutional_date", "foreign_cost", "price_vs_foreign_cost_pct",
        "entry_score", "entry_assessment", "supply_chain", "related_holdings",
        "action", "reason",
    ]]
    table.columns = [
        "股票", "行情日期", "收盤價", "換手率%", "價量變化", "20日量比",
        "籌碼集中度%", "外資買賣超(張)", "內資買賣超(張)",
        "法人日期", "六大外資推估成本", "現價距成本%", "進場分數",
        "進場評估", "供應鏈板塊", "相關持股", "建議操作", "判讀依據",
    ]
    st.dataframe(
        table,
        hide_index=True,
        column_config={
            "收盤價": st.column_config.NumberColumn(format="%.2f"),
            "換手率%": st.column_config.NumberColumn(format="%.2f%%"),
            "20日量比": st.column_config.NumberColumn(format="%.2f 倍"),
            "籌碼集中度%": st.column_config.NumberColumn(format="%+.2f%%"),
            "外資買賣超(張)": st.column_config.NumberColumn(format="%+.0f"),
            "內資買賣超(張)": st.column_config.NumberColumn(format="%+.0f"),
            "六大外資推估成本": st.column_config.NumberColumn(format="%.2f"),
            "現價距成本%": st.column_config.NumberColumn(format="%+.2f%%"),
            "進場分數": st.column_config.NumberColumn(format="%.0f / 100"),
        },
    )
    st.caption("籌碼集中度＝近5日三大法人淨買賣超 ÷ 近5日成交量；內資＝投信＋自營商。")
    st.caption("六大外資成本採逐日移動平均剩餘庫存估算；進場評估依價格相對MA20與KD方向。供應鏈為研究分類，不代表直接供貨關係。")

    related_move_rows = [
        {
            "追蹤股票": stock["name"],
            "供應鏈板塊": stock["supply_chain"],
            "相關公司": item["name"],
            "行情日期": item["date"],
            "收盤價": item["close"],
            "前一日漲跌%": item["daily_pct"],
        }
        for stock in st.session_state.stock_rows
        for item in (stock.get("related_moves") or [])
    ]
    st.subheader("相關供應鏈前一交易日漲跌")
    if related_move_rows:
        related_table = pd.DataFrame(related_move_rows).sort_values(
            ["追蹤股票", "前一日漲跌%"], ascending=[True, False]
        )
        st.dataframe(
            related_table,
            hide_index=True,
            column_config={
                "收盤價": st.column_config.NumberColumn(format="%.2f"),
                "前一日漲跌%": st.column_config.NumberColumn(format="%+.2f%%"),
            },
        )
        st.caption("以各公司最近完整交易日收盤價與前一交易日收盤價計算；不同市場休市日可能不同。")
    else:
        st.info("相關供應鏈行情暫時無法取得。")

    news_rows = [
        item
        for stock in st.session_state.stock_rows
        for item in (stock.get("recent_news") or [])
    ]
    st.subheader("六檔個股近期新聞")
    if news_rows:
        news_table = pd.DataFrame(news_rows)[["stock", "date", "source", "title", "url"]]
        news_table.columns = ["股票", "日期", "來源", "新聞標題", "新聞連結"]
        st.dataframe(
            news_table,
            hide_index=True,
            column_config={
                "新聞連結": st.column_config.LinkColumn("開啟新聞", display_text="閱讀"),
            },
        )
        st.markdown("**手機可點擊新聞**")
        for stock in st.session_state.stock_rows:
            stock_news = stock.get("recent_news") or []
            if not stock_news:
                continue
            with st.container(horizontal=True, horizontal_alignment="left"):
                for item in stock_news:
                    short_title = item["title"] if len(item["title"]) <= 24 else item["title"][:23] + "…"
                    st.link_button(
                        f"{stock['name']}｜{short_title}",
                        item["url"],
                        icon=":material/open_in_new:",
                    )
        st.caption("新聞來自 Google 新聞近7日搜尋結果；請用上方『閱讀』或按鈕開啟，LINE訊息預覽框中的純文字網址本身無法點擊。")
    else:
        st.info("近7日暫無可驗證的個股新聞結果。")

if st.session_state.line_preview:
    st.subheader("本次 LINE 內容")
    st.caption(f"更新時間：{st.session_state.last_updated}")
    st.caption("預覽中的完整網址會顯示為可點擊的超連結；LINE 實際訊息仍保留原始完整網址。")
    with st.container(border=True, height=520):
        text_lines = []

        def flush_preview_text():
            if text_lines:
                st.text("\n".join(text_lines))
                text_lines.clear()

        for raw_line in st.session_state.line_preview.splitlines():
            stripped = raw_line.strip()
            url_match = re.fullmatch(r"https?://\S+", stripped)
            if not url_match:
                text_lines.append(raw_line)
                continue
            flush_preview_text()
            # 方案二：完整網址本身就是超連結，不再只顯示「開啟新聞」按鈕。
            # Streamlit link_button 會處理 URL 跳轉，避免將網址插入不安全的 HTML。
            st.link_button(stripped, stripped, icon=":material/open_in_new:")
        flush_preview_text()

st.info("訊號只做紀律化觀察，不保證報酬；法人資料尚未結算時不會自行推測。")
