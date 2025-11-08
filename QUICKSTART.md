# 🚀 快速設定 5 步驟

## 步驟 1: 修改你的 Python 檔案

在 `上市上櫃全流程.py` 中找到這個函數：

```python
def mount_google_drive():
    """掛載 Google Drive"""
    try:
        from google.colab import drive
        drive.mount('/content/drive', force_remount=False)
        base_dir = '/content/drive/MyDrive'
        print("✓ 已掛載 Google Drive\n")
        return base_dir
    except:
        base_dir = '/content'
        print("✗ 無法掛載 Google Drive，使用本地目錄\n")
        return base_dir
```

**改成：**

```python
def mount_google_drive():
    """使用本地目錄 (GitHub Actions)"""
    base_dir = os.getcwd()
    print("✓ 使用本地目錄:", base_dir, "\n")
    return base_dir
```

---

## 步驟 2: 建立 GitHub Repository

1. 登入 https://github.com
2. 點選右上角 `+` → `New repository`
3. 輸入名稱 (例如: `taiwan-stock-analysis`)
4. 點選 `Create repository`

---

## 步驟 3: 上傳檔案

在 repository 頁面，點選 `uploading an existing file`，上傳：

1. `上市上櫃全流程.py` (已修改的版本)
2. `requirements.txt`
3. `README.md`

---

## 步驟 4: 建立 Workflow 檔案

1. 點選 `Add file` → `Create new file`
2. 檔名輸入: `.github/workflows/daily_stock_analysis.yml`
3. 複製貼上 `daily_stock_analysis.yml` 的內容
4. 點選 `Commit new file`

---

## 步驟 5: 測試執行

1. 點選 `Actions` 頁籤
2. 選擇 `每日台股分析`
3. 點選 `Run workflow` → `Run workflow`
4. 等待執行完成 (約 10-30 分鐘)
5. 完成後可在 Artifacts 下載結果

---

## ⏰ 執行時間

**預設: 每天台灣時間下午 5:00**

要改時間？編輯 `.github/workflows/daily_stock_analysis.yml`：

```yaml
schedule:
  - cron: '0 9 * * *'  # 改這行
```

| 台灣時間 | 改成這個 |
|---------|----------|
| 早上 9:00 | `0 1 * * *` |
| 下午 2:00 | `0 6 * * *` |
| 下午 5:00 | `0 9 * * *` |
| 晚上 9:00 | `0 13 * * *` |

---

## 📦 檔案結構

確保你的 repository 長這樣：

```
你的repository/
├── .github/
│   └── workflows/
│       └── daily_stock_analysis.yml
├── 上市上櫃全流程.py
├── requirements.txt
└── README.md
```

---

## ❓ 常見問題

**Q: 執行失敗怎麼辦？**
A: 點選失敗的執行記錄，查看紅色 ❌ 的錯誤訊息

**Q: 如何下載結果？**
A: Actions → 點選執行記錄 → 往下捲到 Artifacts → 點選下載

**Q: 如何停止自動執行？**
A: Actions → Workflows → 點選 workflow → 右上角 `...` → Disable

**Q: 免費額度夠用嗎？**
A: 每月 2000 分鐘，每天執行一次約 300-900 分鐘，夠用！

---

✅ **完成！系統會每天自動執行並保存結果。**
