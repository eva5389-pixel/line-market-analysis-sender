# 手動免費訂閱版

請先閱讀 [部署與訂閱設定](DEPLOY_SUBSCRIPTIONS.md)。本版先預覽、發給自己確認，再手動發給每日／每週訂閱者。沒有收費功能，尚需部署 Webhook 與持久化資料庫。

以下為歷史設定參考；發送流程以部署說明為準。

# LINE 市場分析發送台

手機可用的 Streamlit 市場分析與 LINE 發送頁面。

## Streamlit Cloud

- Main file：`streamlit_app.py`
- Python dependencies：`requirements.txt`
- 在 App settings → Secrets 設定：

```toml
LINE_CHANNEL_ACCESS_TOKEN = "..."
LINE_TARGET_ID = ""
APP_PASSWORD = "..."
```

`LINE_TARGET_ID` 留空時使用官方帳號廣播；設定使用者或群組 ID 時只推送至指定對象。
真實憑證與密碼不得加入 GitHub。
