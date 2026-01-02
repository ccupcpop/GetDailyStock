#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試數據庫生成腳本
用於診斷為什麼 CSV 存在但數據庫沒有生成
"""

import os
import glob
import pandas as pd
import sqlite3
import traceback

# 請修改為你的 StockTSEHistory 路徑
history_folder = '.\StockTSEHistory'  # 或完整路徑如: '/path/to/StockTSEHistory'

print("="*80)
print("數據庫生成診斷測試")
print("="*80)
print(f"\n1. 檢查資料夾: {history_folder}")

# 檢查資料夾是否存在
if not os.path.exists(history_folder):
    print(f"❌ 資料夾不存在: {history_folder}")
    print("\n請修改腳本中的 history_folder 路徑")
    exit(1)
else:
    print(f"✓ 資料夾存在")

# 檢查 CSV 文件
print(f"\n2. 檢查 CSV 檔案...")
all_csv_files = glob.glob(os.path.join(history_folder, '*.csv'))
print(f"找到 {len(all_csv_files)} 個 CSV 檔案")

if len(all_csv_files) == 0:
    print(f"❌ 沒有找到 CSV 檔案")
    exit(1)

# 顯示前5個文件
print("\n前5個 CSV 檔案:")
for i, csv_file in enumerate(all_csv_files[:5], 1):
    file_size = os.path.getsize(csv_file) / 1024  # KB
    print(f"  {i}. {os.path.basename(csv_file):20s} ({file_size:.1f} KB)")

# 讀取 CSV 文件
print(f"\n3. 讀取 CSV 檔案...")
df_list = []
failed_files = []

for i, csv_file in enumerate(all_csv_files, 1):
    try:
        df = pd.read_csv(csv_file, encoding='utf-8-sig')
        df_list.append(df)
        if i <= 3:
            print(f"  ✓ {os.path.basename(csv_file):20s} ({len(df)} 筆記錄)")
    except Exception as e:
        failed_files.append((csv_file, str(e)))
        if len(failed_files) <= 3:
            print(f"  ✗ {os.path.basename(csv_file):20s} 失敗: {e}")

print(f"\n✓ 成功讀取 {len(df_list)} 個 CSV")
if failed_files:
    print(f"⚠️  讀取失敗 {len(failed_files)} 個 CSV")

if len(df_list) == 0:
    print(f"❌ 沒有可用的 CSV 數據")
    exit(1)

# 合併數據
print(f"\n4. 合併數據...")
try:
    combined_df = pd.concat(df_list, ignore_index=True)
    print(f"✓ 合併完成")
    print(f"  總筆數: {len(combined_df)}")
    print(f"  欄位: {', '.join(combined_df.columns.tolist()[:5])}...")
except Exception as e:
    print(f"❌ 合併失敗: {e}")
    traceback.print_exc()
    exit(1)

# 檢查必要欄位
print(f"\n5. 檢查必要欄位...")
required_columns = ['日期', '股票代碼', '外陸資買賣超張數', '投信買賣超張數', '自營商買賣超張數']
missing_columns = [col for col in required_columns if col not in combined_df.columns]

if missing_columns:
    print(f"❌ 缺少必要欄位: {', '.join(missing_columns)}")
    print(f"\n實際欄位:")
    for col in combined_df.columns:
        print(f"  - {col}")
    exit(1)
else:
    print(f"✓ 所有必要欄位都存在")

# 計算買賣超排序
print(f"\n6. 計算買賣超排序...")
try:
    # 找出最近5個交易日
    latest_dates = combined_df['日期'].drop_duplicates().sort_values(ascending=False).head(5)
    print(f"  最近5個交易日: {', '.join(latest_dates.tolist()[:3])}...")
    
    # 篩選最近5天的數據
    recent_df = combined_df[combined_df['日期'].isin(latest_dates)].copy()
    print(f"  最近5天數據: {len(recent_df)} 筆")
    
    # 計算買賣超
    stock_order = recent_df.groupby('股票代碼').agg({
        '外陸資買賣超張數': 'sum',
        '投信買賣超張數': 'sum',
        '自營商買賣超張數': 'sum'
    })
    
    stock_order['總買賣超'] = (
        stock_order['外陸資買賣超張數'].fillna(0) + 
        stock_order['投信買賣超張數'].fillna(0) + 
        stock_order['自營商買賣超張數'].fillna(0)
    )
    
    stock_order = stock_order.sort_values('總買賣超', ascending=False)
    
    print(f"✓ 計算完成，共 {len(stock_order)} 檔股票")
    print(f"\n前5名買超股票:")
    for i, (code, row) in enumerate(stock_order.head(5).iterrows(), 1):
        print(f"  {i}. {code:10s} 總買賣超: {int(row['總買賣超']):>8,} 張")
    
except Exception as e:
    print(f"❌ 計算失敗: {e}")
    traceback.print_exc()
    exit(1)

# 重新排列數據
print(f"\n7. 按買賣超順序重新排列數據...")
try:
    ordered_dfs = []
    for stock_code in stock_order.index:
        stock_df = combined_df[combined_df['股票代碼'] == stock_code].copy()
        if len(stock_df) > 0:
            ordered_dfs.append(stock_df)
    
    combined_df = pd.concat(ordered_dfs, ignore_index=True)
    print(f"✓ 數據已按買賣超順序排列")
    
except Exception as e:
    print(f"❌ 排列失敗: {e}")
    traceback.print_exc()
    exit(1)

# 生成數據庫
print(f"\n8. 生成數據庫...")
db_name = 'stock_tse.db'
db_path = os.path.join(history_folder, db_name)

try:
    # 刪除舊數據庫
    if os.path.exists(db_path):
        os.remove(db_path)
        print(f"  ✓ 已刪除舊數據庫")
    
    # 創建新數據庫
    conn = sqlite3.connect(db_path)
    combined_df.to_sql('stock_data', conn, if_exists='replace', index=False)
    conn.close()
    
    # 驗證數據庫
    if os.path.exists(db_path):
        db_size = os.path.getsize(db_path) / 1024 / 1024  # MB
        print(f"✓ 數據庫生成成功!")
        print(f"  路徑: {db_path}")
        print(f"  大小: {db_size:.2f} MB")
        
        # 驗證數據
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(DISTINCT 股票代碼) FROM stock_data")
        stock_count = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM stock_data")
        total_count = cursor.fetchone()[0]
        conn.close()
        
        print(f"  股票數: {stock_count} 檔")
        print(f"  總記錄: {total_count} 筆")
    else:
        print(f"❌ 數據庫文件未生成")
        
except Exception as e:
    print(f"❌ 數據庫生成失敗: {e}")
    traceback.print_exc()
    exit(1)

print(f"\n{'='*80}")
print("✅ 測試完成！數據庫已成功生成")
print("="*80)