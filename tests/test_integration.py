import contextlib
import io
import json
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import patch
import pandas as pd

from build_site import build_site, database_summary
import stock_trend
from stock_breakout import save_result, build_watchlist, RULES


class IntegrationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='.tmp-', dir=Path(__file__).parent)
        self.addCleanup(self.temp.cleanup)
        self.base = Path(self.temp.name)
        (self.base / 'StockInfo').mkdir()

    def make_db(self, path, code='0050'):
        path.parent.mkdir(parents=True, exist_ok=True)
        with contextlib.closing(sqlite3.connect(path)) as db, db:
            db.execute('CREATE TABLE stock_data (股票代碼 TEXT, 股票名稱 TEXT, 日期 TEXT, '
                       '開盤價 REAL, 最高價 REAL, 最低價 REAL, 收盤價 REAL)')
            db.execute("INSERT INTO stock_data VALUES (?, '測試', '2026-10-06', 10, 12, 9, 11)", (code,))

    def make_outputs(self):
        for market in ('tse', 'otc'):
            self.make_db(self.base / f'Stock{market.upper()}History' / f'stock_{market}.db')
            self.make_db(self.base / 'StockInfo' / f'stock_{market}_all.db')
            save_result(self.base / 'StockInfo' / f'stock_{market}_breakout.db', pd.DataFrame(), [],
                        {'market': market.upper(), 'latest_date': '2026-10-06', 'candidates': 0, 'watch_candidates': 0,
                         'scanned_stocks': 1, 'eligible_stocks': 1, 'rules': RULES})
            for suffix in ('.xlsx', '_complete.html'):
                (self.base / 'StockInfo' / f'{market}_analysis_result{suffix}').write_text('report')
        (self.base / 'index.html').write_text('<html>site</html>')
        with patch.object(stock_trend, 'BASE_PATH', self.base), contextlib.redirect_stdout(io.StringIO()):
            self.assertTrue(stock_trend.save_to_hot_db([], {}, '2026-10-06'))
            build_watchlist(self.base)

    def test_site_contains_required_data_but_no_source_or_raw_crawler_files(self):
        self.make_outputs()
        (self.base / 'private.txt').write_text('not public')
        (self.base / 'stock_analysis.py').write_text('old intraday crawler')
        charts = self.base / 'output_charts' / '2026.10.06'
        charts.mkdir(parents=True)
        (charts / '上車_0050.html').write_text('<html>chart</html>')
        with contextlib.redirect_stdout(io.StringIO()):
            output = build_site(self.base)
        manifest = json.loads((output / 'StockInfo' / 'manifest.json').read_text(encoding='utf-8'))
        self.assertEqual(manifest['markets']['TSE']['latest_date'], '2026-10-06')
        self.assertEqual(manifest['markets']['WATCHLIST']['stocks'], 0)
        self.assertEqual(manifest['breakouts']['TSE']['candidates'], 0)
        self.assertTrue((output / 'StockInfo' / 'stock_tse_breakout.db').is_file())
        self.assertTrue((output / 'StockInfo' / 'stock_otc_breakout.db').is_file())
        self.assertTrue((output / 'trends' / '上車_0050.html').exists())
        self.assertFalse((output / 'private.txt').exists())
        self.assertFalse((output / 'stock_analysis.py').exists())
        with contextlib.closing(sqlite3.connect(output / 'StockInfo' / 'stock_tse.db')) as db:
            self.assertEqual(db.execute('SELECT 股票代碼 FROM stock_data').fetchone()[0], '0050')

    def test_failed_validation_keeps_previously_built_site(self):
        self.make_outputs()
        with contextlib.redirect_stdout(io.StringIO()):
            output = build_site(self.base)
        previous = (output / 'index.html').read_bytes()
        (self.base / 'StockInfo' / 'otc_analysis_result_complete.html').unlink()
        with self.assertRaises(FileNotFoundError):
            build_site(self.base)
        self.assertEqual((output / 'index.html').read_bytes(), previous)

    def test_empty_watchlist_does_not_retain_yesterdays_selections(self):
        path = self.base / 'StockInfo' / 'stock_hot.db'
        with patch.object(stock_trend, 'BASE_PATH', self.base), contextlib.redirect_stdout(io.StringIO()):
            stock_trend.save_to_hot_db([], {}, '2026-10-05')
            with contextlib.closing(sqlite3.connect(path)) as db, db:
                db.execute("INSERT INTO hot_stocks (股票代碼, 日期) VALUES ('0050', '2026-10-05')")
            stock_trend.save_to_hot_db([], {}, '2026-10-06', is_first_stage=True)
        self.assertEqual(database_summary(path, 'hot_stocks', allow_empty=True)['rows'], 0)

    def test_rejects_stale_breakout_results(self):
        self.make_outputs()
        path = self.base / 'StockInfo' / 'stock_tse_breakout.db'
        with contextlib.closing(sqlite3.connect(path)) as db, db:
            meta = json.loads(db.execute('SELECT json FROM scan_metadata').fetchone()[0])
            meta['latest_date'] = '2026-10-05'
            db.execute('UPDATE scan_metadata SET json = ?', (json.dumps(meta),))
        with self.assertRaisesRegex(ValueError, '爆量候選'):
            build_site(self.base)

    def test_rejects_obsolete_custom_stocks_in_watchlist(self):
        self.make_outputs()
        path = self.base / 'StockInfo' / 'stock_breakout_watch.db'
        with contextlib.closing(sqlite3.connect(path)) as db, db:
            db.execute("INSERT INTO hot_stocks (股票代碼,股票名稱,類型,日期) VALUES ('9999','舊追蹤','上市','2026-10-06')")
        with self.assertRaisesRegex(ValueError, '五日追蹤'):
            build_site(self.base)

    def test_rejects_invalid_or_empty_market_database(self):
        path = self.base / 'StockInfo' / 'broken.db'
        path.write_text('bad database')
        with self.assertRaises(sqlite3.DatabaseError):
            database_summary(path, 'stock_data')
        path.unlink()
        self.make_db(path)
        with contextlib.closing(sqlite3.connect(path)) as db, db:
            db.execute('DELETE FROM stock_data')
        with self.assertRaises(ValueError):
            database_summary(path, 'stock_data')


if __name__ == '__main__':
    unittest.main()
