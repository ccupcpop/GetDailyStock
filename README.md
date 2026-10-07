# 台股每日分析與追蹤

將 GetDailyStock 的收盤資料、StockTracker 的網頁，以及 StockTrend 的量價選股整合在同一個專案。

網站：https://ccupcpop.github.io/GetDailyStock/

## 執行流程

1. 從 TWSE / TPEx 更新上市、上櫃每日行情與三大法人資料。
2. 產生五日法人報表、Excel、個股歷史及 SQLite 資料庫。
3. 使用同一批資料執行 StockTrend 的爆量、連續上漲、法人淨買超篩選。
4. 分別掃描上市／上櫃原始日線，產生「爆量候選」名單與完整圖表資料。
5. 檢查輸出、產生 `_site/`，所有步驟成功後直接發布 GitHub Pages。

已移除 StockTracker 盤中即時報價、委買委賣五檔及買賣比抓取。網站保留上市／上櫃切換、五日法人報表、歷史 K 線與 MA5/10/20/60、成交量、法人當日與累積買賣超、五日爆量追蹤，以及 StockTrend 的獨立選股圖表。

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

## 爆量候選

上市、上櫃各有獨立的「爆量候選」分頁，位於歷史走勢右側。使用相同的最近 75 筆 K 線、MA5/10/20/60、成交量、法人當日與累積買賣超圖表，並顯示首次突破日期、爆量倍數、突破後漲幅及是否跌回盤整區。

`stock_breakout.py` 直接掃描 `StockTSEDaily` / `StockOTCDaily` 的四位數股票及 00 開頭 ETF（排除權證），不依賴靜態公司清單、法人前 150 檔名單或成交量 5,000 張門檻。保留最近 160 個日檔供計算與圖表使用。

篩選條件須同時成立，參數集中於 `stock_breakout.py` 的 `RULES`：

- 訊號日前 20 個交易日均量 ≤ 再往前 40 日均量的 50%。
- 前 20 日最高價 / 最低價 − 1 ≤ 15%。
- 訊號日成交量 ≥ 前 20 日均量的 3 倍，收盤較前日上漲 ≥ 3%。
- 收盤突破前 20 日最高價，且收在當日高低價區間上方 30%；一價漲停視為收在最高。
- 顯示最近 3 個市場交易日首次符合條件者；前 20 日已有完整同類訊號不重複入選。

所有比較基準排除訊號當日。缺漏、停牌、零量或新上市資料不足，不補成零量計算；最新日檔損毀會停止發布。沒有候選時產生空結果，避免沿用舊名單。法人數據另由原始法人日檔補入，缺漏留空。以上是尚未回測的價量形態篩選條件。

可使用既有原始日檔單獨重算：`python stock_breakout.py`。輸出為 `StockInfo/stock_tse_breakout.db`、`stock_otc_breakout.db`，包含候選、歷史及掃描摘要；網站建置會驗證日期與主報表一致。

「追蹤股」使用 `stock_breakout_watch.db` 彙整最近 5 個交易日的兩市候選。入選當天為第 1 天，未再爆量、跌回盤整區或暫無有效行情時仍保留並標示狀態，第 6 個交易日自動移出。每日重新更新歷史圖表、入選後漲幅及最新有效價格日期；不使用日曆天計算，也不因再次放量重置追蹤起點。原本自訂追蹤已移除。

## 資料範圍與追蹤清單

- 原始 CSV：上市／上櫃市場行情與法人資料，預設起始日期為 2025-01-01。
- 下載沿用原專案的增量邏輯：從今天往回檢查，遇到已有檔案停止；較早的漏檔不會自動回補。
- 五日法人報表與原歷史走勢依 `StockInfo/tse_company_list.csv`、`otc_company_list.csv` 篩選；這些清單不會自動更新。爆量候選與五日追蹤直接使用原始日線。
- 個股歷史取最近 240 個日曆天；報表統計使用最近 61 份法人日檔，詳細排行取最近 5 份。
- `stock_tse_all.db` / `stock_otc_all.db` 沿用最近 100 個資料日期中，至少一天成交量大於 5,000 張的條件。
- `stock_tse.db` / `stock_otc.db` 再取法人淨買賣超排序前 100、後 50 檔。
- 「追蹤股」僅彙整上市、上櫃最近 5 個交易日的爆量候選，已移除原本的自訂追蹤清單。
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

測試涵蓋篩選門檻、排除當日的量能基準、缺漏與停牌、五日追蹤及到期移除、跨市場合併、空結果清理、資料日期一致性、舊自訂清單隔離、發布資產範圍及股票代碼前導零。
