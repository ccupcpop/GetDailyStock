"""Run daily data, StockTrend analysis and website staging in dependency order."""
import argparse
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time

from build_site import build_site, database_summary


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--skip-crawler', action='store_true', help='使用已下載的 CSV 重建分析')
    parser.add_argument('--start-date', default='2025-01-01')
    args = parser.parse_args()
    base = Path(__file__).resolve().parent
    env = {**os.environ, 'STOCK_DATA_DIR': str(base), 'TZ': 'Asia/Taipei', 'PYTHONIOENCODING': 'utf-8'}
    daily = [sys.executable, str(base / 'stock_workflow.py'), '--start-date', args.start_date]
    if args.skip_crawler:
        daily.append('--skip-crawler')
    started_ns = time.time_ns()
    subprocess.run(daily, cwd=base, env=env, check=True)
    for market in ('tse', 'otc'):
        database_summary(base / 'StockInfo' / f'stock_{market}_all.db', 'stock_data')
        outputs = [base / 'StockInfo' / f'stock_{market}_all.db',
                   base / f'Stock{market.upper()}History' / f'stock_{market}.db',
                   base / 'StockInfo' / f'{market}_analysis_result.xlsx',
                   base / 'StockInfo' / f'{market}_analysis_result_complete.html']
        if any(not path.is_file() or path.stat().st_mtime_ns < started_ns for path in outputs):
            raise RuntimeError(f'{market.upper()} 分析未完成，停止發布以免使用舊結果')
    charts = base / 'output_charts'
    if charts.is_symlink() or charts.resolve().parent != base:
        raise ValueError('圖表目錄必須在專案內')
    if charts.exists():
        shutil.rmtree(charts)
    subprocess.run([sys.executable, str(base / 'stock_trend.py')], cwd=base, env=env, check=True)
    subprocess.run([sys.executable, str(base / 'stock_breakout.py')], cwd=base, env=env, check=True)
    build_site(base)


if __name__ == '__main__':
    main()
