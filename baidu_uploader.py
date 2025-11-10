#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
百度網盤自動上傳工具
功能:
1. 自動獲取 Access Token
2. 刪除舊的資料夾
3. 上傳新的分析結果
"""

import requests
import os
import glob
import json
import time
from datetime import datetime

class BaiduNetdiskUploader:
    def __init__(self, app_key, secret_key, refresh_token=None):
        self.app_key = app_key
        self.secret_key = secret_key
        self.refresh_token = refresh_token
        self.access_token = None
        self.api_url = "https://pan.baidu.com/rest/2.0/xpan/file"
        self.upload_url = "https://d.pcs.baidu.com/rest/2.0/pcs/superfile2"
        
    def get_access_token(self):
        """使用 refresh_token 獲取 access_token"""
        if not self.refresh_token:
            print("❌ 錯誤: 未提供 refresh_token")
            return None
            
        url = "https://openapi.baidu.com/oauth/2.0/token"
        params = {
            'grant_type': 'refresh_token',
            'refresh_token': self.refresh_token,
            'client_id': self.app_key,
            'client_secret': self.secret_key
        }
        
        try:
            response = requests.get(url, params=params)
            result = response.json()
            
            if 'access_token' in result:
                self.access_token = result['access_token']
                print(f"✓ 成功獲取 Access Token")
                return self.access_token
            else:
                print(f"❌ 獲取 Access Token 失敗: {result}")
                return None
        except Exception as e:
            print(f"❌ 獲取 Access Token 時發生錯誤: {str(e)}")
            return None
    
    def delete_folder(self, folder_path):
        """刪除指定資料夾"""
        if not self.access_token:
            print("❌ 請先獲取 Access Token")
            return False
            
        params = {
            'method': 'filemanager',
            'access_token': self.access_token,
            'opera': 'delete'
        }
        
        data = {
            'filelist': json.dumps([folder_path])
        }
        
        try:
            response = requests.post(self.api_url, params=params, data=data)
            result = response.json()
            
            if result.get('errno') == 0:
                print(f"✓ 成功刪除資料夾: {folder_path}")
                return True
            elif result.get('errno') == -9:
                print(f"ℹ 資料夾不存在: {folder_path} (將創建新資料夾)")
                return True
            else:
                print(f"⚠ 刪除資料夾時出現問題: {result}")
                return False
        except Exception as e:
            print(f"❌ 刪除資料夾時發生錯誤: {str(e)}")
            return False
    
    def create_folder(self, folder_path):
        """創建資料夾"""
        if not self.access_token:
            print("❌ 請先獲取 Access Token")
            return False
            
        params = {
            'method': 'create',
            'access_token': self.access_token,
            'path': folder_path,
            'isdir': 1,
            'rtype': 1
        }
        
        try:
            response = requests.post(self.api_url, params=params)
            result = response.json()
            
            if result.get('errno') in [0, -8]:  # 0=成功, -8=已存在
                print(f"✓ 資料夾已準備: {folder_path}")
                return True
            else:
                print(f"❌ 創建資料夾失敗: {result}")
                return False
        except Exception as e:
            print(f"❌ 創建資料夾時發生錯誤: {str(e)}")
            return False
    
    def upload_file(self, local_path, remote_path):
        """上傳單個檔案到百度網盤"""
        if not self.access_token:
            print("❌ 請先獲取 Access Token")
            return False
            
        try:
            file_size = os.path.getsize(local_path)
            file_name = os.path.basename(local_path)
            remote_file_path = f"{remote_path}/{file_name}"
            
            print(f"  上傳中: {file_name} ({file_size:,} bytes)")
            
            # 1. 預上傳
            precreate_params = {
                'method': 'precreate',
                'access_token': self.access_token,
                'path': remote_file_path,
                'size': file_size,
                'isdir': 0,
                'autoinit': 1,
                'rtype': 1  # 覆蓋同名文件
            }
            
            response = requests.post(self.api_url, data=precreate_params)
            result = response.json()
            
            if result.get('errno') != 0:
                print(f"  ❌ 預上傳失敗: {result}")
                return False
            
            uploadid = result.get('uploadid')
            
            # 2. 分片上傳
            with open(local_path, 'rb') as f:
                file_data = f.read()
                
            upload_params = {
                'method': 'upload',
                'access_token': self.access_token,
                'type': 'tmpfile',
                'path': remote_file_path,
                'uploadid': uploadid,
                'partseq': 0
            }
            
            files = {'file': (file_name, file_data)}
            response = requests.post(self.upload_url, params=upload_params, files=files)
            upload_result = response.json()
            
            if 'md5' not in upload_result:
                print(f"  ❌ 分片上傳失敗: {upload_result}")
                return False
            
            # 3. 創建文件
            create_params = {
                'method': 'create',
                'access_token': self.access_token,
                'path': remote_file_path,
                'size': file_size,
                'isdir': 0,
                'uploadid': uploadid,
                'block_list': json.dumps([upload_result['md5']]),
                'rtype': 1
            }
            
            response = requests.post(self.api_url, data=create_params)
            create_result = response.json()
            
            if create_result.get('errno') == 0:
                print(f"  ✓ {file_name}")
                return True
            else:
                print(f"  ❌ 創建文件失敗: {create_result}")
                return False
                
        except Exception as e:
            print(f"  ❌ 上傳 {local_path} 時發生錯誤: {str(e)}")
            return False
    
    def upload_stock_analysis(self, base_folder="/apps/股票分析數據"):
        """上傳所有股票分析檔案"""
        print("\n" + "="*60)
        print("📊 開始上傳股票分析結果到百度網盤")
        print("="*60)
        
        # 1. 刪除舊資料夾
        print(f"\n🗑️  步驟 1: 清理舊資料...")
        self.delete_folder(base_folder)
        time.sleep(1)  # 等待刪除完成
        
        # 2. 創建新資料夾
        print(f"\n📁 步驟 2: 創建資料夾...")
        if not self.create_folder(base_folder):
            print("❌ 無法創建資料夾,上傳終止")
            return False
        
        # 3. 收集要上傳的檔案
        print(f"\n📦 步驟 3: 收集檔案...")
        files_to_upload = []
        
        # Excel 報表
        excel_files = ['analysis_result.xlsx', 'otc_analysis_result.xlsx']
        for excel_file in excel_files:
            if os.path.exists(excel_file):
                files_to_upload.append(('Excel報表', excel_file))
        
        # HTML 檔案
        for html_dir in ['StockHTML', 'StockOTCHTML']:
            if os.path.exists(html_dir):
                for html_file in glob.glob(os.path.join(html_dir, '*.html')):
                    files_to_upload.append(('HTML圖表', html_file))
        
        # PNG 圖片
        for png_dir in ['StockPNG', 'StockOTCPNG']:
            if os.path.exists(png_dir):
                for png_file in glob.glob(os.path.join(png_dir, '*.png')):
                    files_to_upload.append(('PNG圖表', png_file))
        
        if not files_to_upload:
            print("❌ 沒有找到任何檔案需要上傳")
            return False
        
        print(f"\n找到 {len(files_to_upload)} 個檔案:")
        file_types = {}
        for file_type, _ in files_to_upload:
            file_types[file_type] = file_types.get(file_type, 0) + 1
        for file_type, count in file_types.items():
            print(f"  - {file_type}: {count} 個")
        
        # 4. 上傳所有檔案
        print(f"\n⬆️  步驟 4: 上傳檔案到 {base_folder}")
        print("-" * 60)
        
        success_count = 0
        fail_count = 0
        
        for file_type, file_path in files_to_upload:
            if self.upload_file(file_path, base_folder):
                success_count += 1
            else:
                fail_count += 1
        
        # 5. 顯示結果
        print("-" * 60)
        print(f"\n📈 上傳完成!")
        print(f"  ✓ 成功: {success_count} 個檔案")
        if fail_count > 0:
            print(f"  ✗ 失敗: {fail_count} 個檔案")
        print(f"\n📂 檔案位置: 百度網盤 → {base_folder}")
        print("="*60 + "\n")
        
        return fail_count == 0


def main():
    # 從環境變量讀取配置
    app_key = os.environ.get('BAIDU_APP_KEY')
    secret_key = os.environ.get('BAIDU_SECRET_KEY')
    refresh_token = os.environ.get('BAIDU_REFRESH_TOKEN')
    
    if not all([app_key, secret_key, refresh_token]):
        print("❌ 錯誤: 缺少必要的環境變量")
        print("需要設置:")
        print("  - BAIDU_APP_KEY")
        print("  - BAIDU_SECRET_KEY")
        print("  - BAIDU_REFRESH_TOKEN")
        return 1
    
    # 創建上傳器
    uploader = BaiduNetdiskUploader(app_key, secret_key, refresh_token)
    
    # 獲取 Access Token
    if not uploader.get_access_token():
        print("❌ 無法獲取 Access Token,上傳終止")
        return 1
    
    # 執行上傳
    success = uploader.upload_stock_analysis()
    
    return 0 if success else 1


if __name__ == "__main__":
    exit(main())
