"""Validate generated outputs and stage only public website assets for Pages."""
from __future__ import annotations

import argparse
from contextlib import closing
from datetime import datetime, timezone, timedelta
import html
import json
from pathlib import Path
import shutil
import sqlite3
import tempfile
from urllib.parse import quote


def database_summary(path: Path, table: str, allow_empty: bool = False) -> dict:
    if not path.is_file():
        raise FileNotFoundError(f"缺少資料庫：{path}")
    with closing(sqlite3.connect(path.resolve().as_uri() + '?mode=ro', uri=True)) as db:
        if db.execute('PRAGMA quick_check').fetchone()[0] != 'ok':
            raise ValueError(f"資料庫損毀：{path}")
        columns = {row[1] for row in db.execute(f'PRAGMA table_info({table})')}
        required = {'股票代碼', '股票名稱', '日期', '開盤價', '最高價', '最低價', '收盤價'}
        if table == 'hot_stocks':
            required |= {'類型', 'IS_FOCUS', '操作建議', '信號列表'}
        if not required <= columns:
            raise ValueError(f"資料庫欄位不完整：{path}: {required - columns}")
        count, stocks, latest = db.execute(
            f'SELECT COUNT(*), COUNT(DISTINCT 股票代碼), MAX(日期) FROM {table}'
        ).fetchone()
        if not count and not allow_empty:
            raise ValueError(f"資料庫沒有資料：{path}")
    return {'rows': count, 'stocks': stocks, 'latest_date': latest.replace('.', '-') if latest else None}


def build_site(base_dir: Path) -> Path:
    base_dir = base_dir.resolve()
    output = base_dir / '_site'
    if output.is_symlink() or output.resolve().parent != base_dir:
        raise ValueError('網站輸出目錄必須在專案內')
    info = base_dir / 'StockInfo'
    databases = {
        'TSE': (base_dir / 'StockTSEHistory' / 'stock_tse.db', 'stock_data'),
        'OTC': (base_dir / 'StockOTCHistory' / 'stock_otc.db', 'stock_data'),
        'WATCHLIST': (info / 'stock_hot.db', 'hot_stocks'),
    }
    summaries = {
        market: database_summary(path, table, allow_empty=(market == 'WATCHLIST'))
        for market, (path, table) in databases.items()
    }
    for market in ('tse', 'otc'):
        database_summary(info / f'stock_{market}_all.db', 'stock_data')
    required_files = ['index.html'] + [
        f'StockInfo/{market}_analysis_result{suffix}'
        for market in ('tse', 'otc') for suffix in ('.xlsx', '_complete.html')
    ]
    for filename in required_files:
        path = base_dir / filename
        if not path.is_file() or path.stat().st_size == 0:
            raise FileNotFoundError(f'缺少發布檔案：{path}')
    manifest = {
        'generated_at': datetime.now(timezone(timedelta(hours=8))).isoformat(timespec='seconds'),
        'markets': summaries,
    }
    chart_date = summaries['TSE']['latest_date'].replace('-', '.')
    charts = sorted((base_dir / 'output_charts' / chart_date).glob('*.html'))
    with tempfile.TemporaryDirectory(prefix='.site-build-', dir=base_dir) as temp:
        stage = Path(temp)
        (stage / 'StockInfo').mkdir()
        (stage / 'trends').mkdir()
        for filename in required_files:
            shutil.copy2(base_dir / filename, stage / filename)
        for path, _ in databases.values():
            shutil.copy2(path, stage / 'StockInfo' / path.name)
        for market in ('tse', 'otc'):
            path = info / f'stock_{market}_all.db'
            shutil.copy2(path, stage / 'StockInfo' / path.name)
        for path in charts:
            shutil.copy2(path, stage / 'trends' / path.name)
        links = ''.join(f'<li><a href="{quote(path.name)}">{html.escape(path.stem)}</a></li>' for path in charts)
        if not links:
            links = '<li>本次沒有符合條件的選股圖表。</li>'
        (stage / 'trends' / 'index.html').write_text(
            '<!doctype html><html lang="zh-TW"><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1">'
            '<title>量價選股圖表</title><style>body{font-family:system-ui;max-width:960px;'
            'margin:24px auto;padding:0 16px;background:#f5f7fa}li{margin:16px 0;'
            'overflow-wrap:anywhere}a{color:#245eb8}h1{font-size:1.5rem}</style>'
            f'<a href="../index.html">← 返回法人分析</a><h1>量價選股與追蹤股圖表</h1>'
            f'<p>資料日期：{html.escape(summaries["TSE"]["latest_date"])}</p><ul>{links}</ul></html>',
            encoding='utf-8',
        )
        (stage / 'StockInfo' / 'manifest.json').write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
        (stage / '.nojekyll').touch()
        if output.exists():
            shutil.rmtree(output)
        shutil.copytree(stage, output)
    print(f'網站已建立：{output}；選股圖表 {len(charts)} 張')
    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return output


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--base-dir', type=Path, default=Path(__file__).resolve().parent)
    build_site(parser.parse_args().base_dir)
