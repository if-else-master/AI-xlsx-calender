import pandas as pd
import openpyxl
from openpyxl import load_workbook
import google.generativeai as genai
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import json
import os
from datetime import datetime, timedelta
import re
import pickle

class ExcelCalendarAIParser:
    def __init__(self, gemini_api_key, google_credentials_file=None):
        """
        初始化解析器
        
        Args:
            gemini_api_key (str): Gemini API 金鑰
            google_credentials_file (str): Google 憑證檔案路徑
        """
        # 設定 Gemini API
        genai.configure(api_key=gemini_api_key)
        self.model = genai.GenerativeModel('gemini-pro')
        
        # Google Calendar API 設定
        self.SCOPES = ['https://www.googleapis.com/auth/calendar']
        self.credentials_file = google_credentials_file
        self.calendar_service = None
        
    def setup_google_calendar_api(self):
        """設定 Google Calendar API 認證"""
        creds = None
        
        # 檢查是否有已保存的憑證
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        
        # 如果沒有有效憑證，進行認證流程
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not self.credentials_file:
                    raise ValueError("需要提供 Google 憑證檔案路徑")
                
                flow = Flow.from_client_secrets_file(
                    self.credentials_file, self.SCOPES)
                flow.redirect_uri = 'http://localhost:8080/callback'
                
                # 獲取認證 URL
                auth_url, _ = flow.authorization_url(prompt='consent')
                print(f'請在瀏覽器中開啟此 URL 進行認證: {auth_url}')
                
                # 獲取認證碼
                auth_code = input('輸入認證碼: ')
                flow.fetch_token(code=auth_code)
                creds = flow.credentials
            
            # 保存憑證以供下次使用
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        
        self.calendar_service = build('calendar', 'v3', credentials=creds)
        print("Google Calendar API 認證成功！")

    def read_excel_with_merged_cells(self, file_path, sheet_name=None):
        """
        讀取包含合併格的 Excel 檔案
        
        Args:
            file_path (str): Excel 檔案路徑
            sheet_name (str): 工作表名稱，預設為第一個工作表
        
        Returns:
            dict: 包含原始數據和合併格信息的字典
        """
        try:
            # 使用 openpyxl 讀取檔案以獲得合併格信息
            workbook = load_workbook(file_path, data_only=True)
            if sheet_name:
                worksheet = workbook[sheet_name]
            else:
                worksheet = workbook.active
            
            # 獲取合併格信息
            merged_ranges = []
            for merged_range in worksheet.merged_cells.ranges:
                merged_ranges.append({
                    'range': str(merged_range),
                    'start_row': merged_range.min_row,
                    'end_row': merged_range.max_row,
                    'start_col': merged_range.min_col,
                    'end_col': merged_range.max_col
                })
            
            # 獲取所有儲存格數據
            data = []
            max_row = worksheet.max_row
            max_col = worksheet.max_column
            
            for row in range(1, max_row + 1):
                row_data = []
                for col in range(1, max_col + 1):
                    cell = worksheet.cell(row, col)
                    row_data.append(cell.value)
                data.append(row_data)
            
            return {
                'data': data,
                'merged_ranges': merged_ranges,
                'max_row': max_row,
                'max_col': max_col
            }
            
        except Exception as e:
            raise Exception(f"讀取 Excel 檔案時發生錯誤: {str(e)}")

    def create_ai_prompt_for_calendar(self, excel_data):
        """
        為 AI 建立解析行事曆的提示
        
        Args:
            excel_data (dict): Excel 資料
        
        Returns:
            str: AI 提示文字
        """
        # 將資料轉換為字符串格式
        data_str = ""
        for i, row in enumerate(excel_data['data']):
            data_str += f"第{i+1}列: {row}\n"
        
        merged_str = ""
        for merged in excel_data['merged_ranges']:
            merged_str += f"合併格: {merged['range']} (行 {merged['start_row']}-{merged['end_row']}, 列 {merged['start_col']}-{merged['end_col']})\n"
        
        prompt = f"""
你是一個專業的行事曆資料解析專家。請分析以下來自 Excel 檔案的行事曆數據，這個檔案包含很多合併的儲存格。

Excel 數據：
{data_str}

合併格信息：
{merged_str}

請從這些數據中提取行事曆事件，並以 JSON 格式回傳。每個事件應包含：
1. title: 事件標題
2. start_date: 開始日期 (YYYY-MM-DD)
3. start_time: 開始時間 (HH:MM)，如果沒有具體時間則設為 "09:00"
4. end_date: 結束日期 (YYYY-MM-DD)
5. end_time: 結束時間 (HH:MM)，如果沒有具體時間則設為 "18:00"
6. description: 事件描述（可選）
7. location: 地點（可選）

請特別注意：
- 合併的儲存格可能表示跨多天的事件
- 時間格式可能多樣化，請盡量解析
- 如果某些信息不完整，請根據上下文做合理推測
- 忽略空白或無意義的數據

請只回傳有效的 JSON 陣列，不要包含其他說明文字。

範例格式：
[
  {
    "title": "會議",
    "start_date": "2024-01-15",
    "start_time": "10:00",
    "end_date": "2024-01-15", 
    "end_time": "12:00",
    "description": "重要會議",
    "location": "會議室A"
  }
]
"""
        return prompt

    def parse_calendar_with_ai(self, excel_data):
        """
        使用 AI 解析行事曆數據
        
        Args:
            excel_data (dict): Excel 資料
        
        Returns:
            list: 解析後的行事曆事件列表
        """
        try:
            prompt = self.create_ai_prompt_for_calendar(excel_data)
            
            print("正在使用 AI 解析行事曆數據...")
            response = self.model.generate_content(prompt)
            
            # 嘗試解析 JSON 回應
            response_text = response.text.strip()
            
            # 移除可能的 markdown 格式標記
            if response_text.startswith('```json'):
                response_text = response_text[7:]
            if response_text.endswith('```'):
                response_text = response_text[:-3]
            
            response_text = response_text.strip()
            
            # 解析 JSON
            events = json.loads(response_text)
            
            print(f"成功解析出 {len(events)} 個事件")
            return events
            
        except json.JSONDecodeError as e:
            print(f"JSON 解析錯誤: {str(e)}")
            print(f"AI 回應: {response.text}")
            return []
        except Exception as e:
            print(f"AI 解析錯誤: {str(e)}")
            return []

    def create_google_calendar_event(self, event_data):
        """
        在 Google Calendar 中建立事件
        
        Args:
            event_data (dict): 事件數據
        
        Returns:
            dict: 建立的事件信息
        """
        try:
            # 建立事件對象
            event = {
                'summary': event_data['title'],
                'start': {
                    'dateTime': f"{event_data['start_date']}T{event_data['start_time']}:00",
                    'timeZone': 'Asia/Taipei',
                },
                'end': {
                    'dateTime': f"{event_data['end_date']}T{event_data['end_time']}:00",
                    'timeZone': 'Asia/Taipei',
                },
            }
            
            # 添加描述（如果有）
            if event_data.get('description'):
                event['description'] = event_data['description']
            
            # 添加地點（如果有）
            if event_data.get('location'):
                event['location'] = event_data['location']
            
            # 建立事件
            created_event = self.calendar_service.events().insert(
                calendarId='primary',
                body=event
            ).execute()
            
            return created_event
            
        except HttpError as e:
            print(f"建立 Google Calendar 事件時發生錯誤: {str(e)}")
            return None
        except Exception as e:
            print(f"建立事件時發生錯誤: {str(e)}")
            return None

    def sync_to_google_calendar(self, events):
        """
        將事件同步到 Google Calendar
        
        Args:
            events (list): 事件列表
        
        Returns:
            dict: 同步結果統計
        """
        if not self.calendar_service:
            print("請先設定 Google Calendar API")
            return {'success': 0, 'failed': 0}
        
        success_count = 0
        failed_count = 0
        
        print(f"開始同步 {len(events)} 個事件到 Google Calendar...")
        
        for i, event in enumerate(events, 1):
            print(f"正在建立事件 {i}/{len(events)}: {event['title']}")
            
            result = self.create_google_calendar_event(event)
            if result:
                success_count += 1
                print(f"✓ 成功建立事件: {event['title']}")
            else:
                failed_count += 1
                print(f"✗ 建立事件失敗: {event['title']}")
        
        return {
            'success': success_count,
            'failed': failed_count,
            'total': len(events)
        }

    def process_excel_calendar(self, excel_file_path, sheet_name=None):
        """
        處理整個流程：讀取 Excel -> AI 解析 -> 同步到 Google Calendar
        
        Args:
            excel_file_path (str): Excel 檔案路徑
            sheet_name (str): 工作表名稱
        
        Returns:
            dict: 處理結果
        """
        try:
            print("=" * 50)
            print("開始處理 Excel 行事曆...")
            
            # 1. 讀取 Excel 檔案
            print("步驟 1: 讀取 Excel 檔案...")
            excel_data = self.read_excel_with_merged_cells(excel_file_path, sheet_name)
            print(f"✓ 成功讀取 Excel，包含 {excel_data['max_row']} 行 {excel_data['max_col']} 列")
            print(f"✓ 發現 {len(excel_data['merged_ranges'])} 個合併格")
            
            # 2. 使用 AI 解析
            print("\n步驟 2: 使用 AI 解析行事曆數據...")
            events = self.parse_calendar_with_ai(excel_data)
            
            if not events:
                return {
                    'status': 'failed',
                    'message': '無法從 Excel 中解析出有效的行事曆事件',
                    'events': []
                }
            
            # 3. 同步到 Google Calendar
            print("\n步驟 3: 同步到 Google Calendar...")
            sync_result = self.sync_to_google_calendar(events)
            
            print("\n" + "=" * 50)
            print("處理完成！")
            print(f"總事件數: {sync_result['total']}")
            print(f"成功同步: {sync_result['success']}")
            print(f"同步失敗: {sync_result['failed']}")
            
            return {
                'status': 'success',
                'events': events,
                'sync_result': sync_result
            }
            
        except Exception as e:
            print(f"處理過程中發生錯誤: {str(e)}")
            return {
                'status': 'error',
                'message': str(e)
            }


# 使用範例
def main():
    # 設定 API 金鑰和憑證檔案
    GEMINI_API_KEY = "your_gemini_api_key_here"  # 請替換為您的 Gemini API 金鑰
    GOOGLE_CREDENTIALS_FILE = "credentials.json"  # Google 憑證檔案路徑
    
    # 建立解析器
    parser = ExcelCalendarAIParser(
        gemini_api_key=GEMINI_API_KEY,
        google_credentials_file=GOOGLE_CREDENTIALS_FILE
    )
    
    try:
        # 設定 Google Calendar API
        parser.setup_google_calendar_api()
        
        # 處理 Excel 行事曆
        excel_file = "calendar.xlsx"  # 請替換為您的 Excel 檔案路徑
        result = parser.process_excel_calendar(excel_file)
        
        if result['status'] == 'success':
            print("行事曆同步成功完成！")
        else:
            print(f"處理失敗: {result.get('message', '未知錯誤')}")
            
    except Exception as e:
        print(f"主程式執行錯誤: {str(e)}")


if __name__ == "__main__":
    main()