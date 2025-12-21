#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
直接引用 ConceptHTML 的優化版本
不需要複製檔案到 concept_frames
"""

import os


def merge_with_direct_reference(base_dir):
    """直接引用 ConceptHTML 資料夾中的檔案"""
    
    concept_html_dir = os.path.join(base_dir, 'ConceptHTML')
    stockinfo_dir = os.path.join(base_dir, 'StockInfo')
    output_file = os.path.join(stockinfo_dir, 'Concept_ALL.html')
    
    print("="*80)
    print("生成 Concept_ALL.html (直接引用版本)")
    print("="*80)
    print(f"引用資料夾: {concept_html_dir}")
    print(f"輸出檔案: {output_file}")
    print("="*80)
    
    if not os.path.exists(concept_html_dir):
        print(f"\n❌ 找不到 ConceptHTML 資料夾: {concept_html_dir}")
        return
    
    if not os.path.exists(stockinfo_dir):
        os.makedirs(stockinfo_dir, exist_ok=True)
    
    html_files = [
        'AI伺服器與資料中心.html',
        'IC載板.html',
        '功率半導體.html',
        '先進封裝CoWoS3DIC.html',
        '次世代半導體GaNSiC.html',
        '特殊應用積體電路ASIC.html',
        '國防產業.html',
        '智慧駕駛ADASV2X.html',
        '量子電腦.html',
        '機器人與智慧機械.html'
    ]
    
    print(f"\n檢查檔案:\n")
    
    concept_data = []
    for idx, filename in enumerate(html_files, 1):
        filepath = os.path.join(concept_html_dir, filename)
        
        if not os.path.exists(filepath):
            print(f"[{idx:2d}] ⚠️  找不到: {filename}")
            continue
        
        file_size = os.path.getsize(filepath)
        concept_name = filename.replace('.html', '')
        
        # 關鍵：直接引用 ConceptHTML 資料夾
        concept_data.append({
            'name': concept_name,
            'iframe_src': f'../ConceptHTML/{filename}'  # 相對路徑
        })
        
        print(f"[{idx:2d}] ✓ {concept_name:<30} ({file_size/1024:.1f} KB)")
    
    if not concept_data:
        print("\n❌ 沒有可用的概念股HTML")
        return
    
    print(f"\n{'='*80}")
    print(f"找到 {len(concept_data)}/{len(html_files)} 個概念股")
    print(f"{'='*80}")
    
    # 生成主HTML
    print("\n生成主HTML檔案...")
    merged_html = generate_html(concept_data)
    
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(merged_html)
        
        print(f"\n{'='*80}")
        print("✓ 生成完成！")
        print(f"{'='*80}")
        print(f"檔案: {output_file}")
        print(f"大小: {len(merged_html)/1024:.1f} KB")
        print(f"包含: {len(concept_data)} 個概念股")
        print(f"\n✅ 優點:")
        print("  - 不需要複製檔案")
        print("  - 不需要 concept_frames 資料夾")
        print("  - 節省磁碟空間")
        print("  - 更新 ConceptHTML 會自動反映")
        print(f"\n📁 檔案結構:")
        print("  GetStockDaily/")
        print("  ├── ConceptHTML/         (原始檔案)")
        print("  │   ├── AI伺服器與資料中心.html")
        print("  │   └── ...")
        print("  └── StockInfo/")
        print("      └── Concept_ALL.html (主檔案)")
        print(f"{'='*80}\n")
        
    except Exception as e:
        print(f"\n❌ 儲存失敗: {e}")


def generate_html(concept_data):
    """生成HTML內容"""
    
    html = """<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes">
    <title>台灣概念股產業鏈結構圖</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Microsoft JhengHei', 'PingFang TC', 'Noto Sans TC', sans-serif;
            background: #f5f7fa;
            padding: 20px;
            line-height: 1.6;
            overflow-x: hidden;
        }

        .container {
            max-width: 100%;
            margin: 0 auto;
        }

        .selector-section {
            background: white;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            margin-bottom: 20px;
            position: sticky;
            top: 20px;
            z-index: 100;
        }

        .selector-label {
            font-size: 0.95em;
            color: #666;
            margin-bottom: 12px;
            display: block;
        }

        .concept-selector {
            width: 100%;
            padding: 14px 40px 14px 16px;
            font-size: 1.1em;
            border: 1px solid #ddd;
            border-radius: 8px;
            background: white;
            cursor: pointer;
            appearance: none;
            background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%23333' d='M6 9L1 4h10z'/%3E%3C/svg%3E");
            background-repeat: no-repeat;
            background-position: right 16px center;
            font-family: inherit;
            transition: all 0.2s ease;
        }

        .concept-selector:hover {
            border-color: #4A90E2;
        }

        .concept-selector:focus {
            outline: none;
            border-color: #4A90E2;
            box-shadow: 0 0 0 3px rgba(74, 144, 226, 0.1);
        }

        .iframe-container {
            width: 100%;
            position: relative;
            min-height: 800px;
        }

        .concept-frame {
            width: 100%;
            height: calc(100vh - 180px);
            min-height: 800px;
            border: none;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            position: absolute;
            top: 0;
            left: 0;
            opacity: 0;
            visibility: hidden;
            transition: opacity 0.2s ease-in, visibility 0s linear 0.2s;
        }

        .concept-frame.active {
            opacity: 1;
            visibility: visible;
            transition: opacity 0.2s ease-in, visibility 0s linear 0s;
            z-index: 1;
        }

        .loading-overlay {
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(255, 255, 255, 0.9);
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 8px;
            z-index: 10;
            opacity: 0;
            visibility: hidden;
            transition: opacity 0.2s;
        }

        .loading-overlay.active {
            opacity: 1;
            visibility: visible;
        }

        .loading-spinner {
            width: 50px;
            height: 50px;
            border: 4px solid #f3f3f3;
            border-top: 4px solid #4A90E2;
            border-radius: 50%;
            animation: spin 1s linear infinite;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }

        @media (max-width: 768px) {
            body {
                padding: 12px;
            }

            .selector-section {
                padding: 20px;
                position: relative;
                top: 0;
            }

            .concept-selector {
                font-size: 1em;
                padding: 12px 35px 12px 14px;
            }

            .concept-frame {
                height: calc(100vh - 150px);
                min-height: 600px;
            }

            .iframe-container {
                min-height: 600px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="selector-section">
            <label class="selector-label">選擇產業概念:</label>
            <select class="concept-selector" id="conceptSelector">
"""
    
    for i, concept in enumerate(concept_data):
        html += f'                <option value="{i}">{concept["name"]}</option>\n'
    
    html += """            </select>
        </div>

        <div class="iframe-container" id="iframeContainer">
            <div class="loading-overlay" id="loadingOverlay">
                <div class="loading-spinner"></div>
            </div>
"""
    
    for i, concept in enumerate(concept_data):
        active = ' active' if i == 0 else ''
        html += f'            <iframe id="frame-{i}" class="concept-frame{active}" src="{concept["iframe_src"]}"></iframe>\n'
    
    html += """        </div>
    </div>

    <script>
        let isInitialLoad = true;
        let loadedFrames = new Set();

        function loadConcept(index) {
            document.querySelectorAll('.concept-frame').forEach(function(frame) {
                frame.classList.remove('active');
            });
            
            const targetFrame = document.getElementById('frame-' + index);
            if (targetFrame) {
                targetFrame.classList.add('active');
                
                window.scrollTo({
                    top: 0,
                    behavior: 'smooth'
                });
                
                if (!loadedFrames.has(index)) {
                    loadedFrames.add(index);
                }
            }
        }

        document.addEventListener('DOMContentLoaded', function() {
            const selector = document.getElementById('conceptSelector');
            const loadingOverlay = document.getElementById('loadingOverlay');
            
            if (selector) {
                selector.addEventListener('change', function(e) {
                    loadConcept(this.value);
                });
            }
            
            const firstFrame = document.getElementById('frame-0');
            if (firstFrame) {
                if (isInitialLoad) {
                    loadingOverlay.classList.add('active');
                }
                
                firstFrame.addEventListener('load', function() {
                    setTimeout(function() {
                        loadingOverlay.classList.remove('active');
                        isInitialLoad = false;
                        loadedFrames.add(0);
                    }, 500);
                });
            }
            
            setTimeout(function() {
                for (let i = 1; i < selector.options.length; i++) {
                    const frame = document.getElementById('frame-' + i);
                    if (frame && !loadedFrames.has(i)) {
                        loadedFrames.add(i);
                    }
                }
            }, 2000);
        });
    </script>
</body>
</html>"""
    
    return html


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        base_dir = sys.argv[1]
    else:
        base_dir = os.getcwd()
    
    current_folder = os.path.basename(base_dir)
    if current_folder in ['ConceptHTML', 'StockInfo']:
        base_dir = os.path.dirname(base_dir)
        print(f"偵測到在子資料夾內，自動切換到上層: {base_dir}\n")
    
    merge_with_direct_reference(base_dir)