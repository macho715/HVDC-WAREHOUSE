#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC Warehouse 실무형 인벤토리 대시보드 생성기
- 오빠두엑셀 스타일의 전문적인 대시보드
- KPI 타일, 차트, 조건부서식 자동 생성
- 실무에서 바로 사용 가능한 레이아웃
"""

import pandas as pd
import numpy as np
from collections import Counter
from datetime import datetime
import xlsxwriter
import os
import sys

# scripts 디렉토리를 Python 경로에 추가
sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts'))

from corrected_warehouse_analyzer import CorrectedWarehouseAnalyzer

def create_dashboard():
    """실무형 인벤토리 대시보드 생성"""
    print("=== HVDC Warehouse 실무형 인벤토리 대시보드 생성기 ===")
    print(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 1. 데이터 로드 및 분석
    print("\n📁 데이터 로드 및 분석 중...")
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(current_dir, 'data', 'HVDC WAREHOUSE_HITACHI(HE).xlsx')
    
    if not os.path.exists(excel_path):
        print(f"❌ 오류: 데이터 파일을 찾을 수 없습니다: {excel_path}")
        return
    
    # 분석기 초기화 및 데이터 분석
    analyzer = CorrectedWarehouseAnalyzer(excel_path, sheet_name='CASE LIST')
    analysis_result = analyzer.generate_corrected_report(
        start_date='2023-01-01',
        end_date='2025-12-31'
    )
    
    warehouse_data = analysis_result.get('warehouse_stock', {})
    site_data = analysis_result.get('site_stock', {})
    
    # 2. 대시보드용 데이터 준비
    print("📊 대시보드 데이터 준비 중...")
    
    # 전체 월별 데이터 통합
    all_months = set()
    for data in warehouse_data.values():
        all_months.update(data.index)
    for data in site_data.values():
        all_months.update(data.index)
    
    months_sorted = sorted(list(all_months))
    
    # 통합 월별 데이터프레임 생성
    dashboard_data = []
    for month in months_sorted:
        row = {'월': month}
        
        # 창고별 데이터
        for wh, data in warehouse_data.items():
            if month in data.index:
                row[f'{wh}_입고'] = data.loc[month, '입고'] if '입고' in data.columns else 0
                row[f'{wh}_출고'] = data.loc[month, '출고'] if '출고' in data.columns else 0
                row[f'{wh}_재고'] = data.loc[month, '재고'] if '재고' in data.columns else 0
            else:
                row[f'{wh}_입고'] = 0
                row[f'{wh}_출고'] = 0
                row[f'{wh}_재고'] = 0
        
        # 현장별 데이터
        for site, data in site_data.items():
            if month in data.index:
                row[f'{site}_입고'] = data.loc[month, '입고'] if '입고' in data.columns else 0
                row[f'{site}_누적재고'] = data.loc[month, '누적재고'] if '누적재고' in data.columns else 0
            else:
                row[f'{site}_입고'] = 0
                row[f'{site}_누적재고'] = 0
        
        dashboard_data.append(row)
    
    df_dashboard = pd.DataFrame(dashboard_data)
    df_dashboard.set_index('월', inplace=True)
    
    # 3. KPI 계산
    print("📋 KPI 계산 중...")
    
    # 전체 합계 계산
    total_inbound = sum(df_dashboard[[col for col in df_dashboard.columns if '입고' in col]].sum())
    total_outbound = sum(df_dashboard[[col for col in df_dashboard.columns if '출고' in col]].sum())
    
    # 현재 재고 (마지막 월 기준)
    current_stock = 0
    for wh in warehouse_data.keys():
        if f'{wh}_재고' in df_dashboard.columns:
            current_stock += df_dashboard[f'{wh}_재고'].iloc[-1]
    
    # Dead Stock 계산 (90일 이상)
    today = pd.Timestamp(datetime.today().strftime('%Y-%m-%d'))
    dead_stock_count = 0
    dead_stock_list = []
    
    df = pd.read_excel(excel_path, sheet_name='CASE LIST')
    warehouse_cols = ['DSV Outdoor', 'DSV Indoor', 'DSV Al Markaz', 'Hauler Indoor', 'DSV MZP', 'MOSB']
    site_cols = ['DAS', 'MIR', 'SHU', 'AGI']
    
    for _, row in df.iterrows():
        if not any(pd.notna(row[site]) for site in site_cols):
            inbound_dates = {wh: row[wh] for wh in warehouse_cols if pd.notna(row[wh])}
            if inbound_dates:
                last_date = max(inbound_dates.values())
                days_since = (today - pd.to_datetime(last_date)).days
                if days_since > 90:
                    dead_stock_count += 1
                    dead_stock_list.append({
                        'Case No.': row['Case No.'],
                        '마지막입고일': last_date,
                        '입고후경과일': days_since,
                        '위험도': '높음' if days_since > 180 else '보통'
                    })
    
    # 4. 대시보드 엑셀 파일 생성
    print("📋 대시보드 엑셀 파일 생성 중...")
    output_dir = 'outputs'
    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(output_dir, f'HVDC_Inventory_Dashboard_{timestamp}.xlsx')
    
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        # 데이터 시트 생성
        df_dashboard.to_excel(writer, sheet_name='Data', index=True)
        
        # 워크북과 워크시트 객체 가져오기
        workbook = writer.book
        worksheet = writer.sheets['Data']
        
        # 5. 포맷 정의
        print("🎨 포맷 및 스타일 적용 중...")
        
        # KPI 타일 포맷
        kpi_header_fmt = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'font_color': 'white',
            'bg_color': '#4472C4',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        kpi_value_fmt = workbook.add_format({
            'bold': True,
            'font_size': 18,
            'font_color': 'white',
            'bg_color': '#70AD47',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#,##0'
        })
        
        kpi_warning_fmt = workbook.add_format({
            'bold': True,
            'font_size': 18,
            'font_color': 'white',
            'bg_color': '#C5504B',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#,##0'
        })
        
        # 헤더 포맷
        header_fmt = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'font_color': 'white',
            'bg_color': '#5B9BD5',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # 데이터 포맷
        data_fmt = workbook.add_format({
            'font_size': 10,
            'border': 1,
            'align': 'center',
            'num_format': '#,##0'
        })
        
        # 조건부서식 (Dead Stock)
        dead_stock_warning_fmt = workbook.add_format({
            'font_color': 'white',
            'bg_color': '#FFC000',
            'border': 1
        })
        
        dead_stock_critical_fmt = workbook.add_format({
            'font_color': 'white',
            'bg_color': '#C5504B',
            'border': 1
        })
        
        # 6. KPI 타일 배치 (상단)
        print("📊 KPI 타일 배치 중...")
        
        # KPI 섹션 제목
        worksheet.write('A1', 'HVDC Warehouse 인벤토리 대시보드', workbook.add_format({
            'bold': True, 'font_size': 20, 'font_color': '#4472C4'
        }))
        
        # KPI 타일들
        kpi_start_row = 3
        worksheet.write(f'A{kpi_start_row}', '총 입고', kpi_header_fmt)
        worksheet.write(f'B{kpi_start_row}', total_inbound, kpi_value_fmt)
        
        worksheet.write(f'D{kpi_start_row}', '총 출고', kpi_header_fmt)
        worksheet.write(f'E{kpi_start_row}', total_outbound, kpi_value_fmt)
        
        worksheet.write(f'G{kpi_start_row}', '현재 재고', kpi_header_fmt)
        worksheet.write(f'H{kpi_start_row}', current_stock, kpi_value_fmt)
        
        worksheet.write(f'J{kpi_start_row}', 'Dead Stock', kpi_header_fmt)
        worksheet.write(f'K{kpi_start_row}', dead_stock_count, kpi_warning_fmt)
        
        # 7. 차트 생성
        print("📈 차트 생성 중...")
        
        # 월별 입출고/재고 추이 차트
        chart1 = workbook.add_chart({'type': 'line'})
        
        # 창고별 입고 데이터 추가
        for i, wh in enumerate(warehouse_data.keys()):
            col_name = f'{wh}_입고'
            if col_name in df_dashboard.columns:
                chart1.add_series({
                    'name': f'{wh} 입고',
                    'categories': ['Data', 1, 0, len(df_dashboard), 0],
                    'values': ['Data', 1, df_dashboard.columns.get_loc(col_name) + 1, 
                              len(df_dashboard), df_dashboard.columns.get_loc(col_name) + 1],
                    'line': {'color': '#5B9BD5', 'width': 2.25},
                    'marker': {'type': 'circle', 'size': 4}
                })
        
        chart1.set_title({'name': '창고별 월별 입고 추이', 'font': {'size': 14, 'bold': True}})
        chart1.set_x_axis({'name': '월', 'font': {'size': 10}})
        chart1.set_y_axis({'name': '입고 수량', 'font': {'size': 10}})
        chart1.set_size({'width': 800, 'height': 400})
        worksheet.insert_chart('A8', chart1)
        
        # 월별 출고 추이 차트
        chart2 = workbook.add_chart({'type': 'line'})
        
        for i, wh in enumerate(warehouse_data.keys()):
            col_name = f'{wh}_출고'
            if col_name in df_dashboard.columns:
                chart2.add_series({
                    'name': f'{wh} 출고',
                    'categories': ['Data', 1, 0, len(df_dashboard), 0],
                    'values': ['Data', 1, df_dashboard.columns.get_loc(col_name) + 1, 
                              len(df_dashboard), df_dashboard.columns.get_loc(col_name) + 1],
                    'line': {'color': '#ED7D31', 'width': 2.25},
                    'marker': {'type': 'diamond', 'size': 4}
                })
        
        chart2.set_title({'name': '창고별 월별 출고 추이', 'font': {'size': 14, 'bold': True}})
        chart2.set_x_axis({'name': '월', 'font': {'size': 10}})
        chart2.set_y_axis({'name': '출고 수량', 'font': {'size': 10}})
        chart2.set_size({'width': 800, 'height': 400})
        worksheet.insert_chart('A25', chart2)
        
        # 8. Dead Stock 상세 리스트
        print("🚨 Dead Stock 상세 리스트 생성 중...")
        
        if dead_stock_list:
            dead_stock_df = pd.DataFrame(dead_stock_list)
            dead_stock_df.to_excel(writer, sheet_name='DeadStock', index=False)
            
            # Dead Stock 시트 포맷팅
            dead_stock_ws = writer.sheets['DeadStock']
            
            # 헤더 포맷 적용
            for col_num, value in enumerate(dead_stock_df.columns.values):
                dead_stock_ws.write(0, col_num, value, header_fmt)
            
            # 데이터 포맷 적용 및 조건부서식
            for row_num in range(len(dead_stock_df)):
                for col_num in range(len(dead_stock_df.columns)):
                    value = dead_stock_df.iloc[row_num, col_num]
                    if col_num == 2:  # 입고후경과일 컬럼
                        if value > 180:
                            dead_stock_ws.write(row_num + 1, col_num, value, dead_stock_critical_fmt)
                        elif value > 90:
                            dead_stock_ws.write(row_num + 1, col_num, value, dead_stock_warning_fmt)
                        else:
                            dead_stock_ws.write(row_num + 1, col_num, value, data_fmt)
                    else:
                        dead_stock_ws.write(row_num + 1, col_num, value, data_fmt)
        
        # 9. 현장별 누적 입고 차트
        print("🏗️  현장별 누적 입고 차트 생성 중...")
        
        chart3 = workbook.add_chart({'type': 'column'})
        
        for i, site in enumerate(site_data.keys()):
            col_name = f'{site}_누적재고'
            if col_name in df_dashboard.columns:
                chart3.add_series({
                    'name': f'{site} 누적입고',
                    'categories': ['Data', 1, 0, len(df_dashboard), 0],
                    'values': ['Data', 1, df_dashboard.columns.get_loc(col_name) + 1, 
                              len(df_dashboard), df_dashboard.columns.get_loc(col_name) + 1],
                    'fill': {'color': '#70AD47'},
                    'border': {'color': '#70AD47'}
                })
        
        chart3.set_title({'name': '현장별 누적 입고 현황', 'font': {'size': 14, 'bold': True}})
        chart3.set_x_axis({'name': '월', 'font': {'size': 10}})
        chart3.set_y_axis({'name': '누적 입고 수량', 'font': {'size': 10}})
        chart3.set_size({'width': 800, 'height': 400})
        worksheet.insert_chart('A42', chart3)
        
        # 10. 컬럼 너비 자동 조정
        for col_num, value in enumerate(df_dashboard.columns.values):
            max_length = max(len(str(value)), 
                           df_dashboard[value].astype(str).str.len().max())
            worksheet.set_column(col_num + 1, col_num + 1, max_length + 2)
        
        # 11. 사용법 안내 시트 생성
        print("📖 사용법 안내 시트 생성 중...")
        
        help_ws = workbook.add_worksheet('사용법')
        
        help_content = [
            ['HVDC Warehouse 인벤토리 대시보드 사용법', ''],
            ['', ''],
            ['📊 대시보드 구성', ''],
            ['• 상단 KPI 타일', '총 입고, 총 출고, 현재 재고, Dead Stock 현황'],
            ['• 월별 입고 추이 차트', '창고별 월별 입고량 변화'],
            ['• 월별 출고 추이 차트', '창고별 월별 출고량 변화'],
            ['• 현장별 누적 입고 차트', '현장별 누적 입고 현황'],
            ['• Dead Stock 상세 리스트', '90일 이상 미출고 Case 목록'],
            ['', ''],
            ['🚨 Dead Stock 위험도 구분', ''],
            ['• 노란색 (90-180일)', '주의 필요'],
            ['• 빨간색 (180일+)', '긴급 조치 필요'],
            ['', ''],
            ['🔄 데이터 새로고침', ''],
            ['1. 원본 데이터 파일 업데이트', ''],
            ['2. 대시보드 생성 스크립트 재실행', ''],
            ['3. 새로운 타임스탬프 파일 생성', ''],
            ['', ''],
            ['📞 문의사항', ''],
            ['• 기술 지원: IT팀', ''],
            ['• 데이터 문의: 물류팀', '']
        ]
        
        for row_num, row_data in enumerate(help_content):
            for col_num, value in enumerate(row_data):
                if row_num == 0:
                    help_ws.write(row_num, col_num, value, workbook.add_format({
                        'bold': True, 'font_size': 16, 'font_color': '#4472C4'
                    }))
                elif row_num in [2, 8, 13, 18, 22]:
                    help_ws.write(row_num, col_num, value, workbook.add_format({
                        'bold': True, 'font_size': 12, 'font_color': '#5B9BD5'
                    }))
                else:
                    help_ws.write(row_num, col_num, value)
        
        # 컬럼 너비 조정
        help_ws.set_column('A:A', 30)
        help_ws.set_column('B:B', 40)
    
    print(f"\n✅ 실무형 인벤토리 대시보드 생성 완료!")
    print(f"📄 파일명: {os.path.basename(output_path)}")
    print(f"📄 파일 크기: {os.path.getsize(output_path) / 1024:.1f} KB")
    
    # 12. 엑셀 파일 자동 열기
    try:
        os.startfile(output_path)
        print(f"\n🔓 대시보드가 자동으로 열렸습니다.")
    except:
        print(f"\n💡 대시보드를 수동으로 열어주세요: {output_path}")
    
    print(f"\n📋 대시보드 구성:")
    print("  - 📊 상단 KPI 타일 (총 입고, 총 출고, 현재 재고, Dead Stock)")
    print("  - 📈 월별 입고 추이 차트 (창고별)")
    print("  - 📉 월별 출고 추이 차트 (창고별)")
    print("  - 🏗️  현장별 누적 입고 차트")
    print("  - 🚨 Dead Stock 상세 리스트 (조건부서식)")
    print("  - 📖 사용법 안내 시트")
    
    return output_path

if __name__ == "__main__":
    create_dashboard() 