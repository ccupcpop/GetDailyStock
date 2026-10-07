# 快速開始

```powershell
pip install -r requirements.txt
python run_pipeline.py
python -m http.server 8000 --directory _site
```

開啟 `http://localhost:8000/` 查看上市、上櫃、追蹤股及量價選股圖表。

`python run_pipeline.py --skip-crawler` 可用現有 CSV 重建全部分析。

部署由「台股完整分析與網站發布」GitHub Actions 工作流程完成，下載、分析、選股、發布依序執行。自訂追蹤股位於 `StockInfo/focus_stocks.csv`。完整設定與資料篩選範圍見 [README](README.md)。
