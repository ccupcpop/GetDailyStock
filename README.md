# 台股每日分析與追蹤

將 GetDailyStock 的收盤資料、StockTracker 的網頁，以及 StockTrend 的量價選股整合在同一個專案。

網站：https://ccupcpop.github.io/GetDailyStock/

## 執行流程

1. 從 TWSE / TPEx 更新上市、上櫃每日行情與三大法人資料。
2. 產生五日法人報表、Excel、個股歷史及 SQLite 資料庫。
3. 使用同一批資料執行 StockTrend 的爆量、連續上漲、法人淨買超篩選，並分析自訂追蹤股。
4. 檢查輸出、產生 `_site/`，所有步驟成功後直接發布 GitHub Pages。

已移除 StockTracker 盤中即時報價、委買委賣五檔及買賣比抓取。網站保留上市／上櫃切換、五日法人報表、歷史 K 線與 MA5/10/20/60、成交量、法人當日與累積買賣超、追蹤股分析卡，以及 StockTrend 的獨立選股圖表。

網頁顯示資料的實際日期。GitHub 排程可能延遲，但下載、選股與發布按相依順序執行，不再靠不同專案的時間表配合，也不再跨專案搬移資料。

## 本機執行

```powershell
pip install -r requirements.txt
python run_pipeline.py
python -m http.server 8000 --directory _site
```

瀏覽 `http://localhost:8000/`。如只要用現有 CSV 重算：

```powershell
python run_pipeline.py --skip-crawler
```

整合入口為 `run_pipeline.py`；`stock_workflow.py` 仍可單獨執行每日資料處理，`stock_trend.py` 執行量價分析，`build_site.py` 驗證並建立網站。

## 資料範圍與追蹤清單

- 原始 CSV：上市／上櫃市場行情與法人資料，預設起始日期為 2025-01-01。
- 下載沿用原專案的增量邏輯：從今天往回檢查，遇到已有檔案停止；較早的漏檔不會自動回補。
- 分析依 `StockInfo/tse_company_list.csv`、`otc_company_list.csv` 篩選；這些清單不會自動更新。
- 個股歷史取最近 240 個日曆天；報表統計使用最近 61 份法人日檔，詳細排行取最近 5 份。
- `stock_tse_all.db` / `stock_otc_all.db` 沿用最近 100 個資料日期中，至少一天成交量大於 5,000 張的條件。
- `stock_tse.db` / `stock_otc.db` 再取法人淨買賣超排序前 100、後 50 檔。
- 自訂追蹤股請編輯 `StockInfo/focus_stocks.csv`，股票代碼保留文字格式與前導零。追蹤股也受完整資料庫的篩選範圍限制。
- 量價引擎保留原 StockTrend 規則：成交量、連續三日收盤價上升、三日中至少兩日法人合計淨買超；參數在 `stock_trend.py` 檔案開頭。

## GitHub Actions

唯一工作流程為 `.github/workflows/daily_stock_analysis.yml`。

- 週一至週五台灣時間 18:00 排程啟動。
- 修改程式、網站或清單並推送 `main` 時會執行；也可在 Actions 手動執行。
- 建置成功才會進入發布工作；失敗時保留前次已發布網站。
- Pages 來源設為 **GitHub Actions**，使用同一個 repository 的內建 `GITHUB_TOKEN`，不需要跨專案 PAT。
- 每次結果另存為 Actions artifact，保留 30 天；Excel 可直接從網站下載。
- 網站只發布 `_site/` 中的報表、圖表和資料庫，不上傳整個原始資料倉庫。

原 StockTracker 與 StockTrend 的獨立排程在整合網站驗證成功後停用；舊專案保留供追溯。

## 驗證

```powershell
python -m unittest discover -s tests -v
```

測試涵蓋輸出資產範圍、資料庫欄位與完整性、空追蹤結果清理、驗證失敗時保留原網站，以及股票代碼前導零。
