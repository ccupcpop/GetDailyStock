#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
百度網盤 OAuth 認證工具
用於首次獲取 refresh_token

使用方法:
1. 在本地運行此腳本
2. 在瀏覽器中完成授權
3. 複製獲得的 refresh_token 並保存到 GitHub Secrets
"""

import requests
from urllib.parse import urlencode

# 你的百度應用資訊
APP_KEY = "bQNRgCprki9t7tqWTtI0DsW9xnQRBEWB"
SECRET_KEY = "ZaGivMbwdyDwMJBldmYOPmPo75Nyn6WV"

# 回調地址 (百度開放平台設置的)
REDIRECT_URI = "oob"  # 使用 oob 模式,適合命令行應用

def get_authorization_url():
    """生成授權 URL"""
    auth_url = "https://openapi.baidu.com/oauth/2.0/authorize"
    params = {
        'response_type': 'code',
        'client_id': APP_KEY,
        'redirect_uri': REDIRECT_URI,
        'scope': 'basic,netdisk',
        'display': 'page'
    }
    return f"{auth_url}?{urlencode(params)}"

def get_tokens(code):
    """使用 authorization code 獲取 access_token 和 refresh_token"""
    token_url = "https://openapi.baidu.com/oauth/2.0/token"
    params = {
        'grant_type': 'authorization_code',
        'code': code,
        'client_id': APP_KEY,
        'client_secret': SECRET_KEY,
        'redirect_uri': REDIRECT_URI
    }
    
    response = requests.get(token_url, params=params)
    return response.json()

def main():
    print("="*70)
    print("百度網盤 OAuth 認證")
    print("="*70)
    print()
    
    # 步驟 1: 顯示授權 URL
    auth_url = get_authorization_url()
    print("步驟 1: 在瀏覽器中打開以下 URL 進行授權:")
    print()
    print(auth_url)
    print()
    print("-"*70)
    
    # 步驟 2: 獲取 authorization code
    print()
    print("步驟 2: 授權後,你會看到一個頁面顯示 authorization code")
    print("        或者瀏覽器會跳轉到一個帶有 code 參數的 URL")
    print()
    code = input("請輸入 authorization code: ").strip()
    
    if not code:
        print("❌ 錯誤: 未輸入 authorization code")
        return
    
    # 步驟 3: 獲取 tokens
    print()
    print("正在獲取 tokens...")
    result = get_tokens(code)
    
    print()
    print("="*70)
    
    if 'access_token' in result:
        print("✓ 認證成功!")
        print()
        print("請將以下資訊保存到 GitHub Secrets:")
        print("-"*70)
        print(f"Access Token:  {result['access_token']}")
        print(f"Refresh Token: {result['refresh_token']}")
        print(f"Expires In:    {result.get('expires_in', 'N/A')} 秒")
        print("-"*70)
        print()
        print("📝 在 GitHub Repository 中設置 Secrets:")
        print("   1. Settings → Secrets and variables → Actions")
        print("   2. 點擊 'New repository secret'")
        print("   3. 添加以下 3 個 secrets:")
        print(f"      - Name: BAIDU_APP_KEY")
        print(f"        Secret: {APP_KEY}")
        print(f"      - Name: BAIDU_SECRET_KEY")
        print(f"        Secret: {SECRET_KEY}")
        print(f"      - Name: BAIDU_REFRESH_TOKEN")
        print(f"        Secret: {result['refresh_token']}")
        print()
        print("⚠️  重要: refresh_token 可以長期使用,請妥善保管!")
        print()
    else:
        print("❌ 認證失敗:")
        print(result)
    
    print("="*70)

if __name__ == "__main__":
    main()
