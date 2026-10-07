"""Scan each market's raw daily bars for quiet bases followed by volume breakouts."""
from __future__ import annotations

import argparse
from contextlib import closing
import json
from pathlib import Path
import sqlite3

import numpy as np
import pandas as pd


RULES = {
    'quiet_days': 20, 'reference_days': 40, 'quiet_ratio': 0.5,
    'max_range': 0.15, 'volume_multiple': 3.0, 'min_gain': 0.03,
    'min_close_position': 0.7, 'recent_days': 3, 'watch_days': 5,
}
HISTORY_DAYS = 160  # 75 displayed bars plus enough warm-up for MA60.
PRICES = ['開盤價', '最高價', '最低價', '收盤價']
RAW_COLUMNS = ['證券代號', '證券名稱', '成交股數', *PRICES]
FLOWS = {
    '外陸資買賣超股數(不含外資自營商)': '外陸資買賣超張數',
    '投信買賣超股數': '投信買賣超張數',
    '自營商買賣超股數': '自營商買賣超張數',
}
BAR_COLUMNS = ['日期', '股票代碼', '股票名稱', *PRICES, '成交張數', *FLOWS.values()]
CANDIDATE_COLUMNS = ['code', 'name', 'signal_date', 'volume_multiple', 'quiet_ratio',
                     'base_range', 'day_gain', 'base_high', 'signal_close',
                     'return_since_signal', 'status', 'signal_age', 'latest_price_date']


def read_csv(path: Path, required: list[str]) -> pd.DataFrame:
    for encoding in ('utf-8-sig', 'cp950'):
        try:
            frame = pd.read_csv(path, encoding=encoding, dtype=str,
                                usecols=lambda col: col in required)
            if not set(required) <= set(frame.columns):
                raise ValueError(f'{path.name} 缺少欄位：{set(required) - set(frame.columns)}')
            frame['證券代號'] = frame['證券代號'].str.strip().str.replace(
                r'^="(.*)"$', r'\1', regex=True)
            return frame
        except UnicodeError:
            continue
    raise ValueError(f'{path.name} 無法解碼')


def numbers(values: pd.Series) -> pd.Series:
    return pd.to_numeric(values.str.replace(',', '', regex=False), errors='coerce')


def load_daily(base: Path, market: str) -> tuple[pd.DataFrame, list[str], list[str]]:
    files = sorted((base / f'Stock{market}Daily').glob('????-??-??.csv'))[-HISTORY_DAYS:]
    if not files:
        raise ValueError(f'{market} 沒有每日行情')
    dates = [path.stem for path in files]
    frames, warnings = [], []
    for path in files:
        try:
            frame = read_csv(path, RAW_COLUMNS)
            # Ordinary four-digit equities and listed ETFs; exclude warrants.
            frame = frame[frame['證券代號'].str.fullmatch(r'\d{4}|00\d{3,4}[A-Z]?', na=False)].copy()
            if frame.empty:
                raise ValueError(f'{path.name} 沒有股票行情')
            for col in [*PRICES, '成交股數']:
                frame[col] = numbers(frame[col])
            frame['成交張數'] = frame.pop('成交股數') / 1000
            frame['日期'] = path.stem
            frame = frame.rename(columns={'證券代號': '股票代碼', '證券名稱': '股票名稱'})
            frame['股票名稱'] = frame['股票名稱'].str.strip()
            frame = frame.drop_duplicates('股票代碼', keep='last')
            frames.append(frame)
        except (ValueError, pd.errors.ParserError) as exc:
            if path == files[-1]:
                raise ValueError(f'{market} 最新日檔無法分析，停止發布：{exc}') from exc
            # Keep the date in the market calendar: gaps must not become quiet days.
            warnings.append(str(exc))
    return pd.concat(frames, ignore_index=True), dates, warnings


def valid_bars(frame: pd.DataFrame) -> pd.Series:
    return (frame[PRICES].gt(0).all(axis=1) & frame['成交張數'].gt(0)
            & frame['最高價'].ge(frame[['開盤價', '收盤價']].max(axis=1))
            & frame['最低價'].le(frame[['開盤價', '收盤價']].min(axis=1)))


def find_candidate(bars: pd.DataFrame, dates: list[str], rules: dict = RULES,
                   recent_days: int | None = None) -> dict | None:
    """Use only preceding bars; require complete market-day windows, never zero-fill."""
    frame = bars.set_index('日期').reindex(dates)
    valid = valid_bars(frame)
    q, ref = rules['quiet_days'], rules['reference_days']
    volume = frame['成交張數'].where(valid)
    quiet = volume.shift(1).rolling(q, min_periods=q).mean()
    reference = volume.shift(q + 1).rolling(ref, min_periods=ref).mean()
    high = frame['最高價'].where(valid).shift(1).rolling(q, min_periods=q).max()
    low = frame['最低價'].where(valid).shift(1).rolling(q, min_periods=q).min()
    ratio = quiet / reference
    multiple = volume / quiet
    base_range = high / low - 1
    gain = frame['收盤價'] / frame['收盤價'].shift(1) - 1
    spread = frame['最高價'] - frame['最低價']
    # A one-price limit-up session closes at its high.
    close_position = ((frame['收盤價'] - frame['最低價']) / spread).where(spread.ne(0), 1.0)
    signal = (valid & ratio.le(rules['quiet_ratio']) & base_range.le(rules['max_range'])
              & multiple.ge(rules['volume_multiple']) & gain.ge(rules['min_gain'])
              & frame['收盤價'].gt(high) & close_position.ge(rules['min_close_position']))
    # Do not label a later surge as a first breakout in the same 20-session episode.
    first = signal & ~signal.shift(1, fill_value=False).rolling(q, min_periods=1).max().astype(bool)
    matches = np.flatnonzero(first.to_numpy())
    window = rules['recent_days'] if recent_days is None else recent_days
    recent = [i for i in matches if i >= len(dates) - window]
    if not recent:
        return None
    i = recent[0]
    row = frame.iloc[i]
    latest_index = np.flatnonzero(valid.to_numpy())[-1]
    latest = frame.iloc[latest_index]
    return {
        'code': str(latest['股票代碼']), 'name': str(latest['股票名稱']),
        'signal_date': dates[i], 'volume_multiple': float(multiple.iloc[i]),
        'quiet_ratio': float(ratio.iloc[i]), 'base_range': float(base_range.iloc[i]),
        'day_gain': float(gain.iloc[i]), 'base_high': float(high.iloc[i]),
        'signal_close': float(row['收盤價']),
        'return_since_signal': float(latest['收盤價'] / row['收盤價'] - 1),
        'status': ('最新日無有效行情' if not valid.iloc[-1] else
                   '站穩突破' if latest['收盤價'] > high.iloc[i] else '跌回盤整區'),
        'signal_age': len(dates) - i, 'latest_price_date': dates[latest_index],
    }


def add_institutional(base: Path, market: str, history: pd.DataFrame) -> pd.DataFrame:
    flows = []
    for date, day in history.groupby('日期'):
        path = base / f'Stock{market}Shares' / f'{date}.csv'
        if not path.is_file():
            continue
        frame = read_csv(path, ['證券代號', *FLOWS])
        frame = frame[frame['證券代號'].isin(day['股票代碼'])].copy()
        for source, target in FLOWS.items():
            frame[target] = numbers(frame[source]) / 1000
        frame['日期'] = date
        frame = frame.rename(columns={'證券代號': '股票代碼'})
        flows.append(frame[['日期', '股票代碼', *FLOWS.values()]].drop_duplicates(['日期', '股票代碼']))
    if flows:
        return history.merge(pd.concat(flows), on=['日期', '股票代碼'], how='left', validate='one_to_one')
    for col in FLOWS.values():
        history[col] = np.nan
    return history


def save_result(path: Path, history: pd.DataFrame, candidates: list[dict], metadata: dict,
                watch_candidates: list[dict] | None = None):
    """Replace even empty results, so yesterday's candidates cannot survive a new scan."""
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix('.tmp.db')
    try:
        with closing(sqlite3.connect(temp)) as db, db:
            history.reindex(columns=BAR_COLUMNS).to_sql('stock_data', db, if_exists='replace', index=False)
            pd.DataFrame(candidates, columns=CANDIDATE_COLUMNS).to_sql(
                'candidates', db, if_exists='replace', index=False)
            pd.DataFrame(candidates if watch_candidates is None else watch_candidates,
                         columns=CANDIDATE_COLUMNS).to_sql('watch_candidates', db, if_exists='replace', index=False)
            db.execute('DROP TABLE IF EXISTS scan_metadata')
            db.execute('CREATE TABLE scan_metadata (json TEXT NOT NULL)')
            db.execute('INSERT INTO scan_metadata VALUES (?)', (json.dumps(metadata, ensure_ascii=False),))
            db.execute('CREATE INDEX IF NOT EXISTS stock_date ON stock_data (股票代碼, 日期)')
        temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)


def scan_market(base: Path, market: str) -> dict:
    daily, dates, warnings = load_daily(base, market)
    if len(dates) < RULES['quiet_days'] + RULES['reference_days'] + 1:
        raise ValueError(f'{market} 行情不足 61 個交易日，停止發布')
    latest_codes = set(daily.loc[daily['日期'].eq(dates[-1]), '股票代碼'])
    watch_candidates, eligible = [], 0
    for code, bars in daily.groupby('股票代碼', sort=False):
        recent = bars.set_index('日期').reindex(dates[-61:])
        eligible += int(valid_bars(recent).all())
        result = find_candidate(bars, dates, recent_days=RULES['watch_days'])
        if result:
            watch_candidates.append(result)
    watch_candidates.sort(key=lambda row: (row['signal_date'], row['volume_multiple']), reverse=True)
    candidates = [row for row in watch_candidates if row['signal_age'] <= RULES['recent_days']]
    codes = {row['code'] for row in watch_candidates}
    history = daily[daily['股票代碼'].isin(codes) & valid_bars(daily)].copy()
    history = add_institutional(base, market, history)
    metadata = {'market': market, 'latest_date': dates[-1], 'scanned_stocks': len(latest_codes),
                'eligible_stocks': eligible, 'candidates': len(candidates),
                'watch_candidates': len(watch_candidates), 'rules': RULES,
                'scope': '原始日線的四位數股票與 00 開頭 ETF（不含權證）', 'warnings': warnings}
    save_result(base / 'StockInfo' / f'stock_{market.lower()}_breakout.db', history, candidates, metadata, watch_candidates)
    print(json.dumps(metadata, ensure_ascii=False))
    return metadata


def build_watchlist(base: Path) -> dict:
    """Combine only the last five market sessions' breakout signals."""
    info = base / 'StockInfo'
    history, candidates, market_dates = [], [], {}
    for market, label in [('TSE', '上市'), ('OTC', '上櫃')]:
        path = info / f'stock_{market.lower()}_breakout.db'
        with closing(sqlite3.connect(path.resolve().as_uri() + '?mode=ro', uri=True)) as db:
            bars = pd.read_sql_query('SELECT * FROM stock_data', db)
            picks = pd.read_sql_query('SELECT * FROM watch_candidates', db)
            meta = json.loads(db.execute('SELECT json FROM scan_metadata').fetchone()[0])
        market_dates[market] = meta['latest_date']
        picks['market'] = market
        candidates.append(picks)
        bars = bars.rename(columns={'成交張數': '成交量'})
        bars['類型'] = label
        bars['IS_FOCUS'] = 0
        history.append(bars)
    combined = pd.concat(history, ignore_index=True).drop_duplicates(['類型', '股票代碼', '日期'])
    for column in ('操作建議', '信號列表'):
        if column not in combined:
            combined[column] = None
    picks = pd.concat(candidates, ignore_index=True)
    meta = {'watch_days': RULES['watch_days'], 'market_dates': market_dates,
            'candidates': len(picks),
            'stocks': len(combined[['類型', '股票代碼']].drop_duplicates())}
    path = info / 'stock_breakout_watch.db'
    temp = path.with_suffix('.tmp.db')
    try:
        with closing(sqlite3.connect(temp)) as db, db:
            combined.to_sql('hot_stocks', db, if_exists='replace', index=False)
            picks.to_sql('watch_candidates', db, if_exists='replace', index=False)
            db.execute('DROP TABLE IF EXISTS scan_metadata')
            db.execute('CREATE TABLE scan_metadata (json TEXT NOT NULL)')
            db.execute('INSERT INTO scan_metadata VALUES (?)', (json.dumps(meta, ensure_ascii=False),))
            db.execute('CREATE INDEX IF NOT EXISTS stock_date ON hot_stocks (類型, 股票代碼, 日期)')
        temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)
    print(json.dumps({'watchlist': meta}, ensure_ascii=False))
    return meta


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--base-dir', type=Path, default=Path(__file__).resolve().parent)
    args = parser.parse_args()
    for market in ('TSE', 'OTC'):
        scan_market(args.base_dir, market)
    build_watchlist(args.base_dir)


if __name__ == '__main__':
    main()
