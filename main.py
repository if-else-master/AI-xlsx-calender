# 簡單使用範例
import os
from excel import ExcelCalendarAIParser

def quick_start():
    """快速開始範例"""
    
    # 設定您的 API 金鑰（建議使用環境變數）
    GEMINI_API_KEY = "###"  # 替換為您的實際 API 金鑰
    
    # Google 憑證檔案路徑
    CREDENTIALS_FILE = "credentials.json"  # 確保此檔案在同一目錄下
    
    # Excel 檔案路徑
    EXCEL_FILE = "calendarfiles.xlsx"  # 您的 Excel 行事曆檔案
    
    print("🚀 開始 Excel 行事曆自動同步...")
    
    try:
        # 1. 建立解析器
        parser = ExcelCalendarAIParser(
            gemini_api_key=GEMINI_API_KEY,
            credentials_file=CREDENTIALS_FILE
        )
        
        # 2. 設定 Google Calendar API（首次使用需要瀏覽器認證）
        print("📅 設定 Google Calendar 連接...")
        parser.setup_google_calendar_api()
        
        # 3. 一鍵處理：讀取 Excel + AI 解析 + 同步到 Google Calendar
        print("🤖 使用 AI 處理複雜的合併格...")
        result = parser.process_calendar(EXCEL_FILE)
        
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
    """完整 AI 解析和同步功能（包含時段合併）"""
    
    GEMINI_API_KEY = "AIzaSyA_0U4ZeHCniPaFmm9tmY0keNu4rf_kSzM"  # 使用實際的API金鑰
    CREDENTIALS_FILE = "credentials.json"  # 使用實際的憑證檔案
    EXCEL_FILE = "testfile.xlsx"  # 使用實際的Excel檔案
    
    print("🧪 完整 AI 解析和同步功能（包含時段合併）...")
    
    try:
        # 1. 建立解析器
        print("\n📍 步驟 1: 建立解析器")
        parser = ExcelCalendarAIParser(
            gemini_api_key=GEMINI_API_KEY,
            credentials_file=CREDENTIALS_FILE
        )
        
        # 2. 設定 Google Calendar API
        print("\n📍 步驟 2: 設定 Google Calendar API")
        parser.setup_google_calendar_api()
        
        # 3. 讀取 Excel
        print("\n📍 步驟 3: 讀取 Excel 檔案")
        excel_data = parser.read_excel_file(EXCEL_FILE)
        print(f"✅ 成功讀取：{excel_data['max_row']} 行，{excel_data['max_col']} 列")
        print(f"🔗 合併格數量：{len(excel_data['merged_ranges'])}")
        
        # 4. AI 解析（包含時段合併）
        print("\n📍 步驟 4: AI 解析（包含時段合併）")
        events = parser.ai_parse_calendar(excel_data)
        
        if not events:
            print("❌ 沒有解析出任何事件")
            return
        
        print(f"✅ 成功解析出 {len(events)} 個事件")
        
        # 5. 顯示解析結果
        print("\n📋 解析出的事件（已包含時段合併）：")
        for i, event in enumerate(events, 1):
            print(f"  {i}. 📅 {event['title']}")
            print(f"     ⏰ 時間：{event['start_date']} {event['start_time']} - {event['end_date']} {event['end_time']}")
            if event.get('description'):
                print(f"     📝 描述：{event['description']}")
            if event.get('location'):
                print(f"     📍 地點：{event['location']}")
            print()
        
        # 6. 確認同步
        sync_confirm = input("確定要同步這些事件到 Google Calendar 嗎？(y/n): ").strip().lower()
        if sync_confirm != 'y' and sync_confirm != 'yes':
            print("❌ 同步已取消，僅完成解析測試")
            return
        
        # 7. 同步到 Google Calendar
        print("\n📍 步驟 5: 同步到 Google Calendar")
        sync_result = parser.create_calendar_events(events)
        
        # 8. 顯示結果
        print("\n" + "=" * 60)
        print("🎉 測試完成！")
        print(f"📊 同步結果：")
        print(f"   總事件數: {sync_result['total']}")
        print(f"   成功同步: {sync_result['success']}")
        print(f"   同步失敗: {sync_result['failed']}")
        
        if sync_result['failed'] > 0:
            print("\n❌ 失敗的事件：")
            for failed in sync_result['failed_events'][:3]:
                print(f"   - {failed['event']['title']}: {failed['error']}")
        
        if sync_result['success'] > 0:
            print(f"\n✅ 成功同步 {sync_result['success']} 個事件到 Google Calendar")
            print("📱 現在可以在手機和電腦上的 Google Calendar 中查看您的事件了")
            print("🔗 如果有連堂課程，時段已自動合併")
        
        print("=" * 60)
        
    except Exception as e:
        print(f"❌ 測試失敗：{str(e)}")
        print("\n🔧 故障排除建議：")
        print("1. 確認 Gemini API 金鑰是否正確")
        print("2. 確認 Google 憑證檔案是否存在")
        print("3. 確認 Excel 檔案路徑是否正確")
        print("4. 檢查網路連接是否正常")


def test_merge_consecutive_events():
    """測試連續時段合併功能"""
    print("🧪 測試連續時段合併功能...")
    
    # 建立測試用的解析器
    try:
        parser = ExcelCalendarAIParser("dummy_key", "dummy_credentials.json")
    except:
        print("⚠️ 無法建立完整解析器，使用簡化測試...")
        
        # 直接測試合併邏輯
        from excel import ExcelCalendarAIParser
        
        # 測試數據：模擬課程表中的連續時段
        test_events = [
            {
                "title": "數學",
                "start_date": "2024-12-20",
                "start_time": "08:25",
                "end_date": "2024-12-20", 
                "end_time": "09:05",
                "description": "第一節課",
                "location": "教室101"
            },
            {
                "title": "數學",
                "start_date": "2024-12-20",
                "start_time": "09:15", 
                "end_date": "2024-12-20",
                "end_time": "10:05",
                "description": "第二節課",
                "location": "教室101"
            },
            {
                "title": "英文",
                "start_date": "2024-12-20",
                "start_time": "10:15",
                "end_date": "2024-12-20",
                "end_time": "11:05", 
                "description": "單節課",
                "location": "教室102"
            },
            {
                "title": "物理",
                "start_date": "2024-12-20",
                "start_time": "13:30",
                "end_date": "2024-12-20",
                "end_time": "14:20",
                "description": "第五節課",
                "location": "實驗室"
            },
            {
                "title": "物理",
                "start_date": "2024-12-20", 
                "start_time": "14:30",
                "end_date": "2024-12-20",
                "end_time": "15:20",
                "description": "第六節課",
                "location": "實驗室"
            }
        ]
        
        print("📋 原始事件：")
        for i, event in enumerate(test_events, 1):
            print(f"  {i}. {event['title']} - {event['start_time']} 到 {event['end_time']}")
        
        # 建立一個臨時解析器實例來使用合併方法
        temp_parser = ExcelCalendarAIParser.__new__(ExcelCalendarAIParser)
        merged_events = temp_parser._merge_consecutive_events(test_events)
        
        print("\n📋 合併後事件：")
        for i, event in enumerate(merged_events, 1):
            print(f"  {i}. {event['title']} - {event['start_time']} 到 {event['end_time']}")
            if event.get('description'):
                print(f"     描述：{event['description']}")
        
        # 驗證結果
        print("\n✅ 驗證結果：")
        
        # 應該有3個事件（數學合併為1個，英文1個，物理合併為1個）
        expected_count = 3
        if len(merged_events) == expected_count:
            print(f"✅ 事件數量正確：{len(merged_events)} 個")
        else:
            print(f"❌ 事件數量錯誤：期望 {expected_count} 個，實際 {len(merged_events)} 個")
        
        # 檢查數學課是否正確合併
        math_events = [e for e in merged_events if e['title'] == '數學']
        if len(math_events) == 1:
            math_event = math_events[0]
            if math_event['start_time'] == '08:25' and math_event['end_time'] == '10:05':
                print("✅ 數學課時段合併正確：08:25-10:05")
            else:
                print(f"❌ 數學課時段合併錯誤：{math_event['start_time']}-{math_event['end_time']}")
        else:
            print(f"❌ 數學課合併錯誤：應該1個，實際{len(math_events)}個")
        
        # 檢查物理課是否正確合併  
        physics_events = [e for e in merged_events if e['title'] == '物理']
        if len(physics_events) == 1:
            physics_event = physics_events[0]
            if physics_event['start_time'] == '13:30' and physics_event['end_time'] == '15:20':
                print("✅ 物理課時段合併正確：13:30-15:20")
            else:
                print(f"❌ 物理課時段合併錯誤：{physics_event['start_time']}-{physics_event['end_time']}")
        else:
            print(f"❌ 物理課合併錯誤：應該1個，實際{len(physics_events)}個")
        
        # 檢查英文課是否保持不變
        english_events = [e for e in merged_events if e['title'] == '英文']
        if len(english_events) == 1:
            print("✅ 英文課保持獨立：10:15-11:05")
        else:
            print(f"❌ 英文課處理錯誤：應該1個，實際{len(english_events)}個")
        
        return
    
    # 如果成功建立解析器，進行完整測試
    print("✅ 解析器建立成功，進行完整測試...")



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
    print("1. 完整同步（Excel → AI 解析 → Google Calendar，包含時段合併）")
    print("2. 完整 AI 解析和同步功能（包含時段合併，可選是否同步）") 
    print("3. 測試時段合併邏輯（不需要API金鑰）")
    
    choice = input("輸入選項 (1, 2 或 3): ").strip()
    
    if choice == "1":
        quick_start()
    elif choice == "2":
        test_ai_parsing_only()
    elif choice == "3":
        test_merge_consecutive_events()
    else:
        print("無效選項，執行完整同步...")
        quick_start()