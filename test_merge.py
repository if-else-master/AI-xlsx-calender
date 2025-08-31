#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試時段合併功能
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from excel import ExcelCalendarAIParser

def test_merge_consecutive_events():
    """測試連續時段合併功能"""
    print("🧪 測試連續時段合併功能...")
    
    # 建立測試用的解析器（不需要真正的API金鑰和憑證檔案進行合併測試）
    try:
        parser = ExcelCalendarAIParser("dummy_key", "dummy_credentials.json")
    except:
        # 直接建立一個簡化的測試對象
        class MockParser:
            def _merge_consecutive_events(self, events):
                return ExcelCalendarAIParser._merge_consecutive_events(None, events)
            
            def _time_to_minutes(self, time_str):
                return ExcelCalendarAIParser._time_to_minutes(None, time_str)
        
        parser = MockParser()
    
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
    
    # 執行合併
    merged_events = parser._merge_consecutive_events(test_events)
    
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

if __name__ == "__main__":
    test_merge_consecutive_events()
