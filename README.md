# 台灣股市每日自動分析系統

此專案使用 GitHub Actions 每天自動執行台股分析，包含上市/上櫃資料爬取、分析與圖表生成。

## 📋 功能說明

- 🕐 **每天台灣時間下午 5:00** 自動執行
- 📊 爬取上市/上櫃每日交易資料
- 💰 爬取三大法人買賣超資料
- 📈 生成技術分析圖表 (HTML + PNG)
- 📊 生成 Excel 分析報告
- 💾 自動保存結果 (保留 30 天)

---

## 🚀 快速設定指南

### 步驟 1: 建立 GitHub Repository

1. 登入 GitHub
2. 點選右上角 `+` → `New repository`
3. 輸入 Repository 名稱，例如: `taiwan-stock-analysis`
4. 選擇 `Public` 或 `Private`
5. 點選 `Create repository`

---

### 步驟 2: 上傳檔案到 GitHub

#### 方法 A: 使用 GitHub 網頁介面 (推薦新手)

1. 在你的 repository 頁面，點選 `Add file` → `Upload files`

2. 上傳以下檔案：
   ```
   上市上櫃全流程.py          (你原本的 Python 檔案)
   requirements.txt         (Python 套件清單)
   README.md                (此說明文件)
   ```

3. 建立 `.github/workflows` 資料夾結構：
   - 點選 `Add file` → `Create new file`
   - 在檔名輸入: `.github/workflows/daily_stock_analysis.yml`
   - 貼上 workflow 設定內容
   - 點選 `Commit new file`

#### 方法 B: 使用 Git 指令

```bash
# 1. 初始化 Git (在你的專案資料夾)
git init

# 2. 加入所有檔案
git add .

# 3. 提交變更
git commit -m "Initial commit"

# 4. 連結到 GitHub repository
git remote add origin https://github.com/你的帳號/taiwan-stock-analysis.git

# 5. 推送到 GitHub
git branch -M main
git push -u origin main
```

---

### 步驟 3: 修改主程式檔案

⚠️ **重要：需要修改 `上市上櫃全流程.py`**

找到 `mount_google_drive()` 函數，替換成：

```python
def mount_google_drive():
    """使用本地目錄 (GitHub Actions)"""
    base_dir = os.getcwd()  # 使用當前工作目錄
    print("✓ 使用本地目錄:", base_dir)
    print()
    return base_dir
```

**原因：** GitHub Actions 沒有 Google Drive，要使用本地目錄。

---

### 步驟 4: 檢查檔案結構

確保你的 repository 有以下結構：

```
taiwan-stock-analysis/
├── .github/
│   └── workflows/
│       └── daily_stock_analysis.yml    ← workflow 設定
├── 上市上櫃全流程.py                     ← 主程式
├── requirements.txt                     ← Python 套件
└── README.md                            ← 說明文件
```

---

## ⚙️ 執行時間設定

預設執行時間：**每天台灣時間下午 5:00**

如要修改時間，編輯 `.github/workflows/daily_stock_analysis.yml`：

```yaml
schedule:
  # 每天 UTC 09:00 = 台灣時間 17:00 (下午5點)
  - cron: '0 9 * * *'
```

### 常用時間對照表 (台灣時間 → UTC)

| 台灣時間 | UTC 時間 | Cron 表達式 |
|---------|---------|------------|
| 09:00   | 01:00   | `0 1 * * *` |
| 14:00   | 06:00   | `0 6 * * *` |
| 17:00   | 09:00   | `0 9 * * *` |
| 21:00   | 13:00   | `0 13 * * *` |

### 執行頻率範例

```yaml
# 每天執行一次
- cron: '0 9 * * *'

# 每週一到週五執行
- cron: '0 9 * * 1-5'

# 每週一執行
- cron: '0 9 * * 1'

# 每 6 小時執行一次
- cron: '0 */6 * * *'
```

---

## 🧪 測試執行

### 手動觸發執行

1. 前往你的 GitHub repository
2. 點選 `Actions` 頁籤
3. 選擇 `每日台股分析` workflow
4. 點選右邊的 `Run workflow` 按鈕
5. 點選綠色的 `Run workflow` 確認

### 查看執行結果

1. 在 `Actions` 頁面可以看到執行歷史
2. 點選任一執行記錄查看詳細 log
3. 如果成功，可以下載生成的檔案

---

## 📦 下載執行結果

執行完成後，結果會自動保存為 Artifacts：

1. 進入 `Actions` 頁面
2. 點選某次執行記錄
3. 往下捲到 `Artifacts` 區塊
4. 點選檔案名稱即可下載 ZIP 壓縮檔

### Artifacts 說明

- **stock-analysis-results-XXX**: 完整執行結果 (保留 30 天)
  - 包含所有 CSV、HTML、PNG、Excel 檔案
  
- **latest-stock-data**: 最新資料 (保留 7 天)
  - 只包含 CSV 和 Excel，檔案較小

---

## 🔍 常見問題排解

### Q1: Actions 頁籤沒有出現？

**A:** 確認 `.github/workflows/daily_stock_analysis.yml` 檔案已正確上傳。

### Q2: 執行失敗怎麼辦？

**A:** 點選失敗的執行記錄，查看錯誤訊息：
- 紅色 ❌ 表示該步驟失敗
- 點選展開查看詳細錯誤 log

常見錯誤：
- `ModuleNotFoundError`: 缺少套件 → 檢查 `requirements.txt`
- `Permission denied`: 權限問題 → 檢查檔案路徑
- `mount_google_drive`: 未修改函數 → 見步驟 3

### Q3: 如何停止自動執行？

**A:** 有兩種方法：
1. 刪除 `.github/workflows/daily_stock_analysis.yml` 檔案
2. 在 `Actions` → `Workflows` 頁面，點選 workflow 右側的 `...` → `Disable workflow`

### Q4: 執行時間不準確？

**A:** GitHub Actions 執行時間可能會延遲 3-10 分鐘，這是正常的。

### Q5: 可以同時執行多個時段嗎？

**A:** 可以，在 `schedule` 下加入多個 cron：

```yaml
schedule:
  - cron: '0 1 * * *'   # 早上 9:00
  - cron: '0 9 * * *'   # 下午 5:00
```

---

## 📊 執行流程說明

```
1. 爬取上市每日交易資料 (TWSE Daily)
   ↓
2. 爬取上市三大法人買賣超 (TWSE Institutional)
   ↓
3. 爬取上櫃每日交易資料 (OTC Daily)
   ↓
4. 爬取上櫃三大法人買賣超 (OTC Institutional)
   ↓
5. 清理舊的 History 資料夾
   ↓
6. 生成 TSE 分析報告 (Excel)
   ↓
7. 生成 OTC 分析報告 (Excel)
   ↓
8. 清理舊的圖表資料夾
   ↓
9. 生成 TSE 技術分析圖表 (HTML + PNG)
   ↓
10. 生成 OTC 技術分析圖表 (HTML + PNG)
   ↓
11. 上傳結果到 Artifacts
```

---

## ⚠️ 注意事項

1. **免費額度**: GitHub Actions 免費版每月有 2000 分鐘執行時間
   - 此程式預估每次執行約 10-30 分鐘
   - 每天執行一次，每月約 300-900 分鐘

2. **資料保留**: 
   - Artifacts 最多保留 30 天
   - 需要長期保存請定期下載

3. **執行限制**:
   - 單次執行最長 6 小時
   - 超過會自動中斷

4. **時區**: 所有時間都要用 UTC 時間設定
   - 台灣時間 = UTC + 8 小時

---

## 📝 進階設定

### 加入執行通知 (Slack/Discord/Email)

可以使用 GitHub Actions 的通知套件，在執行完成/失敗時發送通知。

範例：使用 Slack 通知

在 `daily_stock_analysis.yml` 最後加入：

```yaml
- name: Slack Notification
  if: always()
  uses: 8398a7/action-slack@v3
  with:
    status: ${{ job.status }}
    webhook_url: ${{ secrets.SLACK_WEBHOOK }}
```

---

## 🎯 快速檢查清單

設定完成前，請確認：

- [ ] 已建立 GitHub repository
- [ ] 已上傳 `上市上櫃全流程.py`
- [ ] 已建立 `.github/workflows/daily_stock_analysis.yml`
- [ ] 已上傳 `requirements.txt`
- [ ] 已修改 `mount_google_drive()` 函數
- [ ] 已測試手動執行一次
- [ ] 已確認執行成功並能下載結果

---

## 💡 實用資源

- [GitHub Actions 文件](https://docs.github.com/en/actions)
- [Cron 表達式產生器](https://crontab.guru/)
- [GitHub Actions 免費額度說明](https://docs.github.com/en/billing/managing-billing-for-github-actions/about-billing-for-github-actions)

---

**祝您使用順利！有問題歡迎提出。** 🎉
