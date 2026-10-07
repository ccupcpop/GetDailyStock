import contextlib
import io
from pathlib import Path
import sqlite3
import tempfile
import unittest

import pandas as pd

from stock_breakout import (find_candidate, load_daily, save_result, scan_market, build_watchlist,
                            RAW_COLUMNS)


def fixture(count=83):
    dates = pd.bdate_range('2026-01-01', periods=count).strftime('%Y-%m-%d').tolist()
    bars = pd.DataFrame({'日期': dates, '股票代碼': '0050', '股票名稱': '測試',
                         '開盤價': 10.0, '最高價': 10.2, '最低價': 9.8,
                         '收盤價': 10.0, '成交張數': 1000.0})
    bars.loc[60:79, '成交張數'] = 100.0
    bars.loc[80, ['開盤價', '最高價', '最低價', '收盤價', '成交張數']] = [10.2, 10.7, 10.1, 10.6, 600]
    bars.loc[81:, ['開盤價', '最高價', '最低價', '收盤價', '成交張數']] = [10.6, 10.8, 10.5, 10.7, 400]
    return bars, dates


class BreakoutTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='.tmp-', dir=Path(__file__).parent)
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)

    def test_uses_prior_windows_preserves_low_volume_and_leading_zero(self):
        bars, dates = fixture()
        result = find_candidate(bars, dates)
        self.assertEqual(result['code'], '0050')
        self.assertEqual(result['signal_date'], dates[80])
        self.assertAlmostEqual(result['volume_multiple'], 6.0)
        self.assertAlmostEqual(result['quiet_ratio'], 0.1)
        self.assertAlmostEqual(result['return_since_signal'], 10.7 / 10.6 - 1)
        self.assertEqual(result['status'], '站穩突破')

    def test_requires_complete_history_and_never_treats_missing_days_as_zero(self):
        original, dates = fixture()
        for frame in (original.drop(index=75), original.iloc[30:],
                      original.assign(成交張數=original['成交張數'].mask(original.index == 75, 0)),
                      original.assign(收盤價=original['收盤價'].mask(original.index == 75, 0))):
            with self.subTest(rows=len(frame)):
                self.assertIsNone(find_candidate(frame, dates))

    def test_each_price_and_volume_gate_is_required(self):
        original, dates = fixture(81)
        changes = [('成交張數', 80, 250), ('收盤價', 80, 10.2),
                   ('最高價', 80, 12.0), ('最高價', 75, 12.0),
                   ('成交張數', slice(60, 79), 800)]
        for column, row, value in changes:
            bars = original.copy()
            bars.loc[row, column] = value
            with self.subTest(column=column, row=row):
                self.assertIsNone(find_candidate(bars, dates))

    def test_first_signal_ages_out_and_later_surge_is_not_a_new_episode(self):
        bars, dates = fixture(84)
        bars.loc[83, ['開盤價', '最高價', '最低價', '收盤價', '成交張數']] = [10.8, 11.3, 10.8, 11.2, 2000]
        self.assertIsNone(find_candidate(bars, dates))

    def test_five_trading_day_watch_window_and_expiry(self):
        bars, dates = fixture(85)
        self.assertIsNone(find_candidate(bars, dates))  # No longer in the three-day discovery tab.
        tracked = find_candidate(bars, dates, recent_days=5)
        self.assertEqual(tracked['signal_age'], 5)
        self.assertEqual(tracked['signal_date'], dates[80])
        bars, dates = fixture(86)
        self.assertIsNone(find_candidate(bars, dates, recent_days=5))

    def test_watch_keeps_halted_stock_with_last_valid_price_date(self):
        bars, dates = fixture(85)
        bars = bars.drop(index=84)
        tracked = find_candidate(bars, dates, recent_days=5)
        self.assertEqual(tracked['status'], '最新日無有效行情')
        self.assertEqual(tracked['signal_age'], 5)
        self.assertEqual(tracked['latest_price_date'], dates[83])

    def test_watch_merges_markets_without_custom_stocks_and_replaces_expired_list(self):
        bars, dates = fixture(85)
        tracked = find_candidate(bars, dates, recent_days=5)
        info = self.base / 'StockInfo'
        info.mkdir()
        # The obsolete custom-stock database must not influence the new watchlist.
        (info / 'stock_hot.db').write_text('obsolete custom stock data')
        for market in ('tse', 'otc'):
            save_result(info / f'stock_{market}_breakout.db', bars, [],
                        {'latest_date': dates[-1]}, [tracked])
        with contextlib.redirect_stdout(io.StringIO()):
            result = build_watchlist(self.base)
        self.assertEqual(result['candidates'], 2)
        self.assertEqual(result['stocks'], 2)  # Same code in different markets stays separate.
        with contextlib.closing(sqlite3.connect(info / 'stock_breakout_watch.db')) as db:
            self.assertEqual(db.execute('SELECT COUNT(*) FROM watch_candidates').fetchone()[0], 2)
            self.assertEqual(db.execute('SELECT COUNT(DISTINCT 類型) FROM hot_stocks').fetchone()[0], 2)
        for market in ('tse', 'otc'):
            save_result(info / f'stock_{market}_breakout.db', pd.DataFrame(), [],
                        {'latest_date': dates[-1]}, [])
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(build_watchlist(self.base)['stocks'], 0)
        with contextlib.closing(sqlite3.connect(info / 'stock_breakout_watch.db')) as db:
            self.assertEqual(db.execute('SELECT COUNT(*) FROM hot_stocks').fetchone()[0], 0)

    def test_marks_failed_breakout_without_using_future_prices_to_trigger_it(self):
        bars, dates = fixture()
        bars.loc[82, ['開盤價', '最高價', '最低價', '收盤價']] = [10.1, 10.2, 9.8, 10.0]
        result = find_candidate(bars, dates)
        self.assertEqual(result['signal_date'], dates[80])
        self.assertEqual(result['status'], '跌回盤整區')

    def test_one_price_limit_up_is_supported(self):
        bars, dates = fixture(81)
        bars.loc[80, ['開盤價', '最高價', '最低價', '收盤價']] = 11.0
        self.assertIsNotNone(find_candidate(bars, dates))

    def test_empty_results_replace_old_candidates_and_history(self):
        bars, dates = fixture()
        path = self.base / 'test.db'
        save_result(path, bars, [find_candidate(bars, dates)], {'candidates': 1})
        save_result(path, pd.DataFrame(), [], {'candidates': 0})
        with contextlib.closing(sqlite3.connect(path)) as db:
            self.assertEqual(db.execute('SELECT COUNT(*) FROM candidates').fetchone()[0], 0)
            self.assertEqual(db.execute('SELECT COUNT(*) FROM stock_data').fetchone()[0], 0)

    def test_raw_market_scan_is_independent_of_existing_filtered_databases(self):
        bars, dates = fixture(81)
        for market, code, encoding in [('TSE', '0050', 'cp950'), ('OTC', '1234', 'utf-8-sig')]:
            folder = self.base / f'Stock{market}Daily'
            folder.mkdir()
            for i, row in bars.iterrows():
                raw = pd.DataFrame([{
                    '證券代號': f'="{code}"' if market == 'TSE' else code, '證券名稱': '測試',
                    '成交股數': f'{row["成交張數"] * 1000:,.0f}',
                    **{col: row[col] for col in ('開盤價', '最高價', '最低價', '收盤價')},
                }], columns=RAW_COLUMNS)
                raw.to_csv(folder / f'{dates[i]}.csv', encoding=encoding, index=False)
            with contextlib.redirect_stdout(io.StringIO()):
                meta = scan_market(self.base, market)
            self.assertEqual(meta['candidates'], 1)
            with contextlib.closing(sqlite3.connect(self.base / 'StockInfo' / f'stock_{market.lower()}_breakout.db')) as db:
                self.assertEqual(db.execute('SELECT code FROM candidates').fetchone()[0], code)
                self.assertEqual(db.execute('SELECT COUNT(*) FROM stock_data').fetchone()[0], 81)
                self.assertIsNone(db.execute('SELECT 投信買賣超張數 FROM stock_data').fetchone()[0])

    def test_latest_invalid_raw_file_fails_instead_of_reusing_old_data(self):
        folder = self.base / 'StockTSEDaily'
        folder.mkdir()
        (folder / '2026-10-07.csv').write_text('broken\n1', encoding='utf-8')
        with self.assertRaisesRegex(ValueError, '最新日檔'):
            load_daily(self.base, 'TSE')


if __name__ == '__main__':
    unittest.main()
