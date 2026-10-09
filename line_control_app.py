#!/usr/bin/env python3
"""LINE 官方帳號市場分析手動發送頁面。"""

import os
from datetime import datetime

import pandas as pd
import streamlit as st

import app as report


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


def require_cloud_password() -> None:
    """雲端 LINE 發送頁必須先通過密碼，避免公開網址遭他人濫用。"""
    try:
        app_password = str(st.secrets.get("APP_PASSWORD", ""))
    except Exception:
        app_password = ""
    using_cloud_secrets = os.getenv("LINE_CONFIG_SOURCE") == "streamlit_secrets"
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
st.caption("追蹤：緯穎、南電、金像電、欣興、所羅門。按一次即可更新資料並發送到官方 LINE。")

config = report.load_line_config()
if config["enabled"] and config["channel_access_token"]:
    destination = "廣播給官方帳號好友" if config["delivery_mode"] == "broadcast" else "指定對象"
    st.success(f"LINE 已連線；傳送方式：{destination}。憑證不會顯示在頁面或報告中。")
else:
    st.warning("LINE 尚未完成設定。請先雙擊「LINE官方串接設定.command」輸入 Channel access token。")

st.session_state.setdefault("line_preview", "")
st.session_state.setdefault("stock_rows", [])
st.session_state.setdefault("taiwan_vix", None)
st.session_state.setdefault("last_updated", "")

send_clicked = st.button(
    "產生最新分析並發送 LINE",
    type="primary",
    disabled=not (config["enabled"] and config["channel_access_token"]),
)

if send_clicked:
    refresh_key = datetime.now().isoformat(timespec="seconds")
    with st.spinner("正在更新市場、個股與新聞資料…"):
        message, stock_rows, taiwan_vix = generate_report_payload(refresh_key)
    st.session_state.line_preview = message
    st.session_state.stock_rows = stock_rows
    st.session_state.taiwan_vix = taiwan_vix
    st.session_state.last_updated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if report.send_line_report(message):
        st.success("已成功發送到 LINE。")
    else:
        st.error("LINE 發送失敗，請檢查 Channel token、傳送模式與接收者設定。")

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
    st.subheader("五檔個股近期新聞")
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
        st.caption("新聞來自 Google 新聞近7日搜尋結果；點擊連結可回到原刊登來源查核全文。")
    else:
        st.info("近7日暫無可驗證的個股新聞結果。")

if st.session_state.line_preview:
    st.subheader("本次 LINE 內容")
    st.caption(f"更新時間：{st.session_state.last_updated}")
    st.text_area(
        "LINE 訊息預覽",
        value=st.session_state.line_preview,
        height=520,
        disabled=True,
    )

st.info("訊號只做紀律化觀察，不保證報酬；法人資料尚未結算時不會自行推測。")
