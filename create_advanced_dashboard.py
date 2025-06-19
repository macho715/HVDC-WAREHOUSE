#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC Warehouse 완전 자동화 실무형 대시보드 생성기
- Win32Com을 활용한 피벗테이블/피벗차트/슬라이서 자동 생성
- 오빠두엑셀 스타일의 뉴모피즘 도형 및 KPI 타일
- 실무에서 바로 사용 가능한 전문적인 레이아웃
"""

import pandas as pd
import numpy as np
from collections import Counter
from datetime import datetime
import xlsxwriter
import win32com.client as win32
import os
import sys
import time

# scripts 디렉토리를 Python 경로에 추가
sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts'))

from corrected_warehouse_analyzer import CorrectedWarehouseAnalyzer

def create_advanced_dashboard():
    """완전 자동화된 실무형 대시보드 생성"""
    print("=== HVDC Warehouse 완전 자동화 실무형 대시보드 생성기 ===")
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
    
    # 4. xlsxwriter로 기본 데이터 및 KPI 생성
    print("📋 기본 데이터 및 KPI 생성 중...")
    output_dir = 'outputs'
    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(output_dir, f'HVDC_Advanced_Dashboard_{timestamp}.xlsx')
    
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        # 데이터 시트 생성
        df_dashboard.to_excel(writer, sheet_name='Data', index=False)
        
        # 워크북과 워크시트 객체 가져오기
        workbook = writer.book
        worksheet = writer.sheets['Data']
        
        # KPI 타일 포맷
        kpi_header_fmt = workbook.add_format({
            'bold': True,
            'font_size': 16,
            'font_color': 'white',
            'bg_color': '#4472C4',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        kpi_value_fmt = workbook.add_format({
            'bold': True,
            'font_size': 20,
            'font_color': 'white',
            'bg_color': '#70AD47',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#,##0'
        })
        
        kpi_warning_fmt = workbook.add_format({
            'bold': True,
            'font_size': 20,
            'font_color': 'white',
            'bg_color': '#C5504B',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#,##0'
        })
        
        # 대시보드 제목
        worksheet.write('A1', 'HVDC Warehouse 완전 자동화 인벤토리 대시보드', workbook.add_format({
            'bold': True, 'font_size': 24, 'font_color': '#4472C4'
        }))
        
        # KPI 타일들 (상단)
        kpi_start_row = 3
        worksheet.merge_range(f'A{kpi_start_row}:B{kpi_start_row}', '총 입고', kpi_header_fmt)
        worksheet.merge_range(f'C{kpi_start_row}:D{kpi_start_row}', total_inbound, kpi_value_fmt)
        
        worksheet.merge_range(f'F{kpi_start_row}:G{kpi_start_row}', '총 출고', kpi_header_fmt)
        worksheet.merge_range(f'H{kpi_start_row}:I{kpi_start_row}', total_outbound, kpi_value_fmt)
        
        worksheet.merge_range(f'K{kpi_start_row}:L{kpi_start_row}', '현재 재고', kpi_header_fmt)
        worksheet.merge_range(f'M{kpi_start_row}:N{kpi_start_row}', current_stock, kpi_value_fmt)
        
        worksheet.merge_range(f'P{kpi_start_row}:Q{kpi_start_row}', 'Dead Stock', kpi_header_fmt)
        worksheet.merge_range(f'R{kpi_start_row}:S{kpi_start_row}', dead_stock_count, kpi_warning_fmt)
        
        # 기본 차트 생성
        print("📈 기본 차트 생성 중...")
        
        # 월별 입고 추이 차트
        chart1 = workbook.add_chart({'type': 'line'})
        
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
    
    print(f"✅ 기본 데이터 및 KPI 생성 완료: {output_path}")
    
    # 5. Win32Com으로 고급 기능 추가
    print("\n⚙️ Win32Com으로 고급 기능 추가 중...")
    
    try:
        # Excel 애플리케이션 시작
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False  # 백그라운드에서 실행
        excel.DisplayAlerts = False
        
        # 워크북 열기
        wb = excel.Workbooks.Open(os.path.abspath(output_path))
        ws = wb.Sheets('Data')
        
        # 6. 뉴모피즘 스타일 도형 생성
        print("🎨 뉴모피즘 스타일 도형 생성 중...")
        
        # 배경색 설정
        ws.Range('A1:Z50').Interior.Color = int('0xF8F9FA', 16)  # 연한 회색 배경
        
        # 대시보드 제목 도형
        sh = ws.Shapes
        title_shape = sh.AddShape(1, 0, 0, 1200, 50)  # 직사각형 도형
        title_shape.Fill.ForeColor.RGB = int('0x4472C4', 16)  # 파란색 배경
        title_shape.TextFrame.Characters().Text = "HVDC Warehouse 완전 자동화 인벤토리 대시보드"
        title_shape.TextFrame.Characters().Font.Size = 20
        title_shape.TextFrame.Characters().Font.Bold = True
        title_shape.TextFrame.Characters().Font.Color = 0xFFFFFF  # 흰색 텍스트
        title_shape.TextFrame.HorizontalAlignment = 1  # 가운데 정렬
        
        # KPI 타일 도형들
        kpi_positions = [
            (0, 60, 200, 40, '총 입고', total_inbound, 0x70AD47),
            (220, 60, 200, 40, '총 출고', total_outbound, 0x70AD47),
            (440, 60, 200, 40, '현재 재고', current_stock, 0x70AD47),
            (660, 60, 200, 40, 'Dead Stock', dead_stock_count, 0xC5504B)
        ]
        
        for left, top, width, height, label, value, color in kpi_positions:
            # KPI 박스
            kpi_box = sh.AddShape(5, left, top, width, height)
            kpi_box.Fill.ForeColor.RGB = color
            
            # KPI 라벨
            label_shape = sh.AddShape(1, left, top, width, height//2)
            label_shape.Fill.Visible = False
            label_shape.Line.Visible = False
            label_shape.TextFrame.Characters().Text = label
            label_shape.TextFrame.Characters().Font.Size = 12
            label_shape.TextFrame.Characters().Font.Bold = True
            label_shape.TextFrame.Characters().Font.Color = 0xFFFFFF
            label_shape.TextFrame.HorizontalAlignment = 1
            label_shape.TextFrame.VerticalAlignment = 2
            
            # KPI 값
            value_shape = sh.AddShape(1, left, top + height//2, width, height//2)
            value_shape.Fill.Visible = False
            value_shape.Line.Visible = False
            value_shape.TextFrame.Characters().Text = f"{value:,}"
            value_shape.TextFrame.Characters().Font.Size = 16
            value_shape.TextFrame.Characters().Font.Bold = True
            value_shape.TextFrame.Characters().Font.Color = 0xFFFFFF
            value_shape.TextFrame.HorizontalAlignment = 1
            value_shape.TextFrame.VerticalAlignment = 1
        
        # 7. 피벗테이블 생성
        print("📊 피벗테이블 생성 중...")
        
        # 피벗 시트 생성
        pivot_ws = wb.Sheets.Add()
        pivot_ws.Name = 'Pivot'
        
        # 피벗 캐시 생성
        data_range = ws.UsedRange
        pc = wb.PivotCaches().Create(SourceType=1, SourceData=data_range)
        
        # 피벗테이블 생성
        pt = pc.CreatePivotTable(TableDestination='Pivot!R1C1', TableName='PivotInventory')
        
        # 피벗 필드 설정
        pt.PivotFields('월').Orientation = 1  # xlRowField
        pt.PivotFields('월').Position = 1
        
        # 데이터 필드 추가
        for field in ['DSV Outdoor_입고', 'DSV Indoor_입고', 'DSV Al Markaz_입고']:
            if field in [f.name for f in pt.PivotFields()]:
                pt.PivotFields(field).Orientation = 4  # xlDataField
                pt.PivotFields(field).Function = -4157  # xlSum
        
        # 8. 피벗차트 생성
        print("📈 피벗차트 생성 중...")
        
        # 피벗차트 생성
        chart = pivot_ws.Shapes.AddChart2(201, 4, 0, 0, 800, 400)  # xlLine
        chart.Chart.SetSourceData(pt.TableRange1)
        chart.Chart.ChartType = 4  # xlLine
        chart.Chart.HasTitle = True
        chart.Chart.ChartTitle.Text = '창고별 월별 입고 추이 (피벗차트)'
        chart.Chart.HasLegend = True
        
        # 9. 슬라이서 생성 (Windows Excel만 가능)
        print("🔧 슬라이서 생성 중...")
        
        try:
            # 슬라이서 캐시 생성
            slicer_cache = wb.SlicerCaches.Add(pt, '월')
            
            # 슬라이서 추가
            slicer = slicer_cache.Slicers.Add(pivot_ws, '월 필터', 850, 50, 200, 200)
            slicer.Style = 'SlicerStyleLight1'
        except Exception as e:
            print(f"⚠️ 슬라이서 생성 실패 (일부 Excel 버전에서 지원하지 않음): {e}")
        
        # 10. 차트 박스 생성 (뉴모피즘 스타일)
        print("📦 차트 박스 생성 중...")
        
        # 메인 대시보드로 돌아가기
        ws = wb.Sheets('Data')
        
        # 차트 영역 박스들
        chart_boxes = [
            (0, 120, 400, 250, '입고 추이'),
            (420, 120, 400, 250, '출고 추이'),
            (840, 120, 400, 250, '재고 현황')
        ]
        
        for left, top, width, height, title in chart_boxes:
            # 차트 박스
            box = sh.AddShape(5, left, top, width, height)
            box.Fill.ForeColor.RGB = int('0xFFFFFF', 16)
            box.Fill.Transparency = 0.1
            box.Line.ForeColor.RGB = int('0xE0E0E0', 16)
            box.Line.Weight = 1
            
            # 박스 제목
            title_shape = sh.AddShape(1, left, top - 20, width, 20)
            title_shape.Fill.Visible = False
            title_shape.Line.Visible = False
            title_shape.TextFrame.Characters().Text = title
            title_shape.TextFrame.Characters().Font.Size = 12
            title_shape.TextFrame.Characters().Font.Bold = True
            title_shape.TextFrame.Characters().Font.Color = int('0x4472C4', 16)
            title_shape.TextFrame.HorizontalAlignment = 1
        
        # 11. 사용법 안내 시트 생성
        print("📖 사용법 안내 시트 생성 중...")
        
        help_ws = wb.Sheets.Add()
        help_ws.Name = '사용법'
        
        help_content = [
            ['HVDC Warehouse 완전 자동화 인벤토리 대시보드 사용법', ''],
            ['', ''],
            ['📊 대시보드 구성', ''],
            ['• 상단 KPI 타일', '총 입고, 총 출고, 현재 재고, Dead Stock 현황'],
            ['• 뉴모피즘 스타일 차트 박스', '입고/출고/재고 추이 차트'],
            ['• 피벗테이블/피벗차트', '동적 데이터 분석 및 시각화'],
            ['• 슬라이서', '월별 필터링 기능 (Windows Excel)'],
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
            ['⚙️ 고급 기능', ''],
            ['• 피벗테이블: 동적 데이터 분석', ''],
            ['• 피벗차트: 실시간 차트 업데이트', ''],
            ['• 슬라이서: 월별 필터링', ''],
            ['• 뉴모피즘: 현대적인 UI 디자인', ''],
            ['', ''],
            ['📞 문의사항', ''],
            ['• 기술 지원: IT팀', ''],
            ['• 데이터 문의: 물류팀', '']
        ]
        
        for row_num, row_data in enumerate(help_content):
            for col_num, value in enumerate(row_data):
                cell = help_ws.Cells(row_num + 1, col_num + 1)
                cell.Value = value
                
                if row_num == 0:
                    cell.Font.Size = 16
                    cell.Font.Bold = True
                    cell.Font.Color = int('0x4472C4', 16)
                elif row_num in [2, 8, 13, 18, 23, 28]:
                    cell.Font.Size = 12
                    cell.Font.Bold = True
                    cell.Font.Color = int('0x5B9BD5', 16)
        
        # 컬럼 너비 조정
        help_ws.Columns('A').ColumnWidth = 40
        help_ws.Columns('B').ColumnWidth = 50
        
        # 12. 파일 저장 및 종료
        print("💾 파일 저장 중...")
        wb.Save()
        wb.Close()
        excel.Quit()
        
        print(f"\n✅ 완전 자동화 실무형 대시보드 생성 완료!")
        print(f"📄 파일명: {os.path.basename(output_path)}")
        print(f"📄 파일 크기: {os.path.getsize(output_path) / 1024:.1f} KB")
        
        # 13. 엑셀 파일 자동 열기
        try:
            os.startfile(output_path)
            print(f"\n🔓 대시보드가 자동으로 열렸습니다.")
        except:
            print(f"\n💡 대시보드를 수동으로 열어주세요: {output_path}")
        
        print(f"\n📋 고급 대시보드 구성:")
        print("  - 📊 뉴모피즘 스타일 KPI 타일")
        print("  - 📈 피벗테이블/피벗차트 (동적 분석)")
        print("  - 🔧 슬라이서 (월별 필터링)")
        print("  - 🎨 뉴모피즘 스타일 차트 박스")
        print("  - 📖 상세 사용법 안내")
        
        return output_path
        
    except Exception as e:
        print(f"\n❌ Win32Com 오류 발생: {str(e)}")
        print("💡 기본 대시보드만 생성되었습니다.")
        return output_path

if __name__ == "__main__":
    create_advanced_dashboard() 