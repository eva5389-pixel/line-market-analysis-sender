# 免費好友訂閱與手動發送

此版提供免費的每日／每週接收名單及暫停訂閱，沒有付款、VIP 或收費驗證。

## 管理頁流程

1. 登入管理頁，按「產生預覽（不發送）」。
2. 檢查內容，按「發送給自己」；LINE_TARGET_ID 必須為管理者個人 User ID。
3. 選擇每日或每週名單，核對人數，勾選確認後手動發送。

每週名單收到的是當次預覽，尚無自動一週彙整。不設定自動排程。CLI 執行只產出報告，不發送。發送為可複製、轉傳的純文字，包含新聞原文網址。

## 部署仍需完成

- 管理頁與 webhook 必須使用同一持久化磁碟上的 SQLite 檔案；設定 `SUBSCRIPTIONS_DB` 為絕對路徑。不要使用臨時磁碟，也不要把資料庫提交 GitHub。
- 建議在同一台有持久化磁碟的主機執行 Streamlit 與 `gunicorn webhook:app`，以 HTTPS 反向代理提供 `/webhook`。Streamlit Community Cloud 與另一台 webhook 主機不能各用自己的 SQLite；需遷移到同一持久化主機或另外實作共用資料庫後再啟用。
- 管理頁設定 `LINE_CHANNEL_ACCESS_TOKEN`、`LINE_TARGET_ID`、`APP_PASSWORD`；webhook 設定 `LINE_CHANNEL_SECRET`。使用環境變數／私密設定，不要提交憑證。
- Streamlit 雲端私密設定須保留 APP_PASSWORD，無密碼時不啟用發送。
- LINE Developers 的 webhook URL 指向該 HTTPS `/webhook`，通過驗證後開啟 webhook。
- 訂閱方式：好友傳送「每日訂閱」、「每週訂閱」、「暫停訂閱」。新加入好友預設暫停，必須自行選擇。可在官方帳號管理介面設定以上文字的選單與歡迎說明；本次程式不會自動建立選單或傳送歡迎訊息。

## 重複發送與查核

同一人與完全相同內容只嘗試一次。網路逾時會保留 pending 紀錄，後續略過，避免未知結果重複推播；需管理者人工查核，不會自動重試或刪除紀錄。若重新產生的內容不同，視為不同報告。備份資料庫時應採 SQLite 備份工具，保護其中 User ID。

## 驗證

`python -m unittest -v test_subscriptions.py`

測試使用臨時資料庫與模擬 API，不傳送真實 LINE 訊息。GitHub 程式更新不代表 Webhook、持久化主機或帳號選單已接通。
