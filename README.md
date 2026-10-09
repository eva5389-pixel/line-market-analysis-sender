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
