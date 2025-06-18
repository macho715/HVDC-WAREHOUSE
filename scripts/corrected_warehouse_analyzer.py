#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
정확한 Case별 이벤트 기반 창고 분석기
- 각 Case의 실제 이동 순서를 추적
- 창고 간 이동 시 이중 카운트 방지
- 정확한 월별 입출고/재고 산출
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

class CorrectedWarehouseAnalyzer:
    def __init__(self, excel_path, sheet_name='CASE LIST'):
        """
        정확한 창고 분석기 초기화
        
        Args:
            excel_path (str): 엑셀 파일 경로
            sheet_name (str): 시트 이름
        """
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.df = pd.read_excel(excel_path, sheet_name=sheet_name)
        
        # 날짜 컬럼들을 datetime으로 변환
        date_columns = ['DSV Indoor', 'DSV Al Markaz', 'DSV Outdoor', 'Hauler Indoor', 'DSV MZP', 'MOSB', 'MIR', 'SHU', 'DAS', 'AGI']
        for col in date_columns:
            if col in self.df.columns:
                self.df[col] = pd.to_datetime(self.df[col], errors='coerce')
        
        # 창고 및 현장 컬럼 정의
        self.warehouse_cols = ['DSV Indoor', 'DSV Al Markaz', 'DSV Outdoor', 'Hauler Indoor', 'DSV MZP', 'MOSB']
        self.site_cols = ['MIR', 'SHU', 'DAS', 'AGI']
        
        print(f"입고 컬럼: {self.warehouse_cols}")
        print(f"출고 컬럼: {self.site_cols}")
    
    def get_case_timeline(self, case_row):
        """
        개별 Case의 이벤트 타임라인 생성
        
        Args:
            case_row: Case 데이터 행
            
        Returns:
            list: 시간순 정렬된 이벤트 리스트 [(date, location, event_type), ...]
        """
        events = []
        case_no = case_row['Case No.']
        
        # 입고 이벤트 수집
        for warehouse in self.warehouse_cols:
            if pd.notna(case_row[warehouse]):
                events.append((case_row[warehouse], warehouse, 'warehouse_in'))
        
        # 출고 이벤트 수집
        for site in self.site_cols:
            if pd.notna(case_row[site]):
                events.append((case_row[site], site, 'site_out'))
        
        # 시간순 정렬
        events = sorted(events, key=lambda x: x[0])
        return events
    
    def calculate_monthly_events(self, start_date='2023-01-01', end_date='2025-12-31'):
        """
        월별 이벤트 계산 (정확한 알고리즘)
        
        Args:
            start_date (str): 시작 날짜
            end_date (str): 종료 날짜
            
        Returns:
            dict: 월별 이벤트 데이터
        """
        start_date = pd.to_datetime(start_date)
        end_date = pd.to_datetime(end_date)
        
        # 월별 기간 생성 (월초 기준으로 수정)
        months = pd.date_range(start=start_date, end=end_date, freq='MS')
        
        # 창고별 월별 이벤트 초기화
        warehouse_monthly = {warehouse: {
            'inbound': {month: 0 for month in months},
            'outbound': {month: 0 for month in months},
            'stock': {month: 0 for month in months}
        } for warehouse in self.warehouse_cols}
        
        # 현장별 월별 이벤트 초기화
        site_monthly = {site: {
            'inbound': {month: 0 for month in months},
            'cumulative': {month: 0 for month in months}
        } for site in self.site_cols}
        
        # Case별 이벤트 처리
        for idx, row in self.df.iterrows():
            events = self.get_case_timeline(row)
            if not events:
                continue
            
            # Case별 상태 추적
            current_warehouse = None
            
            for date, location, event_type in events:
                # 월별 집계를 위한 월초 날짜 계산
                month_key = date.replace(day=1)
                
                if event_type == 'warehouse_in':
                    # 이전 창고에서 출고 처리
                    if current_warehouse is not None:
                        warehouse_monthly[current_warehouse]['outbound'][month_key] += 1
                    
                    # 새 창고 입고
                    warehouse_monthly[location]['inbound'][month_key] += 1
                    current_warehouse = location
                    
                elif event_type == 'site_out':
                    # 창고에서 출고
                    if current_warehouse is not None:
                        warehouse_monthly[current_warehouse]['outbound'][month_key] += 1
                    
                    # 현장 입고
                    site_monthly[location]['inbound'][month_key] += 1
                    current_warehouse = None
            
            # 마지막 창고에 잔재고 처리
            if current_warehouse is not None:
                # 마지막 이벤트의 월부터 모든 월에 재고 반영
                last_month = events[-1][0].replace(day=1)
                for month in months:
                    if month >= last_month:
                        warehouse_monthly[current_warehouse]['stock'][month] += 1
        
        # 현장별 누적 재고 계산
        for site in self.site_cols:
            cumulative = 0
            for month in months:
                cumulative += site_monthly[site]['inbound'][month]
                site_monthly[site]['cumulative'][month] = cumulative
        
        return warehouse_monthly, site_monthly
    
    def generate_corrected_report(self, start_date='2023-01-01', end_date='2025-12-31'):
        """
        정확한 종합 리포트 생성
        
        Args:
            start_date (str): 시작 날짜
            end_date (str): 종료 날짜
            
        Returns:
            dict: 정확한 분석 결과
        """
        print("=== 정확한 Case별 이벤트 기반 분석 시작 ===")
        print(f"분석 기간: {start_date} ~ {end_date}")
        print(f"총 케이스 수: {len(self.df)}")
        
        # 월별 이벤트 계산
        warehouse_monthly, site_monthly = self.calculate_monthly_events(start_date, end_date)
        
        # DataFrame으로 변환
        warehouse_stock = {}
        site_stock = {}
        
        # 창고별 DataFrame 생성
        for warehouse in self.warehouse_cols:
            months = list(warehouse_monthly[warehouse]['inbound'].keys())
            data = {
                '월': [month.strftime('%Y-%m') for month in months],
                '입고': [warehouse_monthly[warehouse]['inbound'][month] for month in months],
                '출고': [warehouse_monthly[warehouse]['outbound'][month] for month in months],
                '재고': [warehouse_monthly[warehouse]['stock'][month] for month in months]
            }
            warehouse_stock[warehouse] = pd.DataFrame(data).set_index('월')
        
        # 현장별 DataFrame 생성
        for site in self.site_cols:
            months = list(site_monthly[site]['inbound'].keys())
            data = {
                '월': [month.strftime('%Y-%m') for month in months],
                '입고': [site_monthly[site]['inbound'][month] for month in months],
                '누적재고': [site_monthly[site]['cumulative'][month] for month in months]
            }
            site_stock[site] = pd.DataFrame(data).set_index('월')
        
        # Dead Stock 분석
        dead_stock = self.analyze_dead_stock()
        
        # 요약 정보 출력
        self.print_summary(warehouse_stock, site_stock, dead_stock)
        
        return {
            'warehouse_stock': warehouse_stock,
            'site_stock': site_stock,
            'dead_stock': dead_stock
        }
    
    def analyze_dead_stock(self, days_threshold=90):
        """
        Dead Stock 분석 (90일 이상 미출고)
        
        Args:
            days_threshold (int): Dead Stock 기준일수
            
        Returns:
            DataFrame: Dead Stock 목록
        """
        dead_stock_cases = []
        current_date = datetime.now()
        
        for idx, row in self.df.iterrows():
            case_no = row['Case No.']
            events = self.get_case_timeline(row)
            
            if not events:
                continue
            
            # 마지막 이벤트 확인
            last_event = events[-1]
            last_date, last_location, last_event_type = last_event
            
            # 마지막 이벤트가 창고 입고이고, 이후 출고가 없으면 Dead Stock
            if last_event_type == 'warehouse_in':
                days_since_last = (current_date - last_date).days
                if days_since_last >= days_threshold:
                    dead_stock_cases.append({
                        'Case No.': case_no,
                        '마지막위치': last_location,
                        '마지막입고일': last_date,
                        '입고후경과일': days_since_last
                    })
        
        return pd.DataFrame(dead_stock_cases)
    
    def print_summary(self, warehouse_stock, site_stock, dead_stock):
        """
        분석 결과 요약 출력
        """
        print("\n=== 창고별 최근 12개월 요약 ===")
        for warehouse, df in warehouse_stock.items():
            recent_12 = df.tail(12)
            total_inbound = recent_12['입고'].sum()
            total_outbound = recent_12['출고'].sum()
            current_stock = recent_12['재고'].iloc[-1] if len(recent_12) > 0 else 0
            
            print(f"{warehouse}:")
            print(f"  - 최근 12개월 입고: {total_inbound}건")
            print(f"  - 최근 12개월 출고: {total_outbound}건")
            print(f"  - 현재 재고: {current_stock}건")
        
        print("\n=== 현장별 최근 12개월 요약 ===")
        for site, df in site_stock.items():
            recent_12 = df.tail(12)
            total_inbound = recent_12['입고'].sum()
            current_stock = recent_12['누적재고'].iloc[-1] if len(recent_12) > 0 else 0
            
            print(f"{site}:")
            print(f"  - 최근 12개월 입고: {total_inbound}건")
            print(f"  - 누적 재고: {current_stock}건")
        
        print(f"\n=== Dead Stock 분석 ({len(dead_stock)}건) ===")
        if len(dead_stock) > 0:
            print("상위 10건:")
            print(dead_stock.head(10)[['Case No.', '마지막위치', '입고후경과일']]) 