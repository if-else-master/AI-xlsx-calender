# 簡單使用範例
import os
from excel import ExcelCalendarAIParser

def quick_start():
    """快速開始範例"""
    
    # 設定您的 API 金鑰（建議使用環境變數）
    GEMINI_API_KEY = ""  # 替換為您的實際 API 金鑰
    
    # Google 憑證檔案路徑
    CREDENTIALS_FILE = "credentials.json"  # 確保此檔案在同一目錄下
    
    # Excel 檔案路徑
    EXCEL_FILE = "testfile.xlsx"  # 您的 Excel 行事曆檔案
    
    print("🚀 開始 Excel 行事曆自動同步...")
    
    try:
        # 1. 建立解析器
        parser = ExcelCalendarAIParser(
            gemini_api_key=GEMINI_API_KEY,
            google_credentials_file=CREDENTIALS_FILE
        )
        
        # 2. 設定 Google Calendar API（首次使用需要瀏覽器認證）
        print("📅 設定 Google Calendar 連接...")
        parser.setup_google_calendar_api()
        
        # 3. 一鍵處理：讀取 Excel + AI 解析 + 同步到 Google Calendar
        print("🤖 使用 AI 處理複雜的合併格...")
        result = parser.process_excel_calendar(EXCEL_FILE)
        
        # 4. 顯示結果
        if result['status'] == 'success':
            sync_info = result['sync_result']
            print(f"""
✅ 同步完成！
📊 統計資訊：
   - 總事件數：{sync_info['total']}
   - 成功同步：{sync_info['success']}
   - 失敗事件：{sync_info['failed']}
            """)
            
            # 顯示解析出的事件
            print("📋 解析出的事件：")
            for i, event in enumerate(result['events'][:5], 1):  # 顯示前5個事件
                print(f"  {i}. {event['title']} - {event['start_date']} {event['start_time']}")
            
            if len(result['events']) > 5:
                print(f"  ... 還有 {len(result['events']) - 5} 個事件")
                
        else:
            print(f"❌ 處理失敗：{result.get('message', '未知錯誤')}")
            
    except FileNotFoundError as e:
        print(f"❌ 檔案未找到：{str(e)}")
        print("請確認以下檔案存在：")
        print(f"  - {EXCEL_FILE}")
        print(f"  - {CREDENTIALS_FILE}")
        
    except Exception as e:
        print(f"❌ 發生錯誤：{str(e)}")
        print("\n🔧 故障排除建議：")
        print("1. 確認 Gemini API 金鑰是否正確")
        print("2. 確認 Google 憑證檔案是否存在")
        print("3. 確認 Excel 檔案路徑是否正確")
        print("4. 檢查網路連接是否正常")


def test_ai_parsing_only():
    """僅測試 AI 解析功能（不同步到 Google Calendar）"""
    
    GEMINI_API_KEY = "your_gemini_api_key_here"
    EXCEL_FILE = "calendar.xlsx"
    
    print("🧪 測試 AI 解析功能...")
    
    try:
        # 只需要 Gemini API，不需要 Google 憑證
        parser = ExcelCalendarAIParser(gemini_api_key=GEMINI_API_KEY)
        
        # 讀取 Excel
        print("📖 讀取 Excel 檔案...")
        excel_data = parser.read_excel_with_merged_cells(EXCEL_FILE)
        print(f"✅ 成功讀取：{excel_data['max_row']} 行，{excel_data['max_col']} 列")
        print(f"🔗 合併格數量：{len(excel_data['merged_ranges'])}")
        
        # AI 解析
        print("🤖 AI 解析中...")
        events = parser.parse_calendar_with_ai(excel_data)
        
        if events:
            print(f"✅ 成功解析出 {len(events)} 個事件：\n")
            for i, event in enumerate(events, 1):
                print(f"{i}. 📅 {event['title']}")
                print(f"   ⏰ 時間：{event['start_date']} {event['start_time']} - {event['end_date']} {event['end_time']}")
                if event.get('description'):
                    print(f"   📝 描述：{event['description']}")
                if event.get('location'):
                    print(f"   📍 地點：{event['location']}")
                print()
        else:
            print("❌ 沒有解析出任何事件")
            
    except Exception as e:
        print(f"❌ 測試失敗：{str(e)}")


def check_requirements():
    """檢查必要套件是否已安裝"""
    
    required_packages = [
        'pandas',
        'openpyxl', 
        'google.generativeai',
        'google.auth',
        'googleapiclient'
    ]
    
    print("🔍 檢查必要套件...")
    
    missing_packages = []
    for package in required_packages:
        try:
            __import__(package)
            print(f"✅ {package}")
        except ImportError:
            missing_packages.append(package)
            print(f"❌ {package} (未安裝)")
    
    if missing_packages:
        print(f"\n💡 請安裝缺少的套件：")
        print("pip install pandas openpyxl google-generativeai google-auth google-auth-oauthlib google-api-python-client")
    else:
        print("\n✅ 所有必要套件已安裝！")


if __name__ == "__main__":
    print("Excel 行事曆 AI 解析器")
    print("=" * 30)
    
    # 檢查套件
    check_requirements()
    print()
    
    # 選擇執行模式
    print("請選擇執行模式：")
    print("1. 完整同步（Excel → AI 解析 → Google Calendar）")
    print("2. 僅測試 AI 解析功能")
    
    choice = input("輸入選項 (1 或 2): ").strip()
    
    if choice == "1":
        quick_start()
    elif choice == "2":
        test_ai_parsing_only()
    else:
        print("無效選項，執行完整同步...")
        quick_start()