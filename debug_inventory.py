import pandas as pd
import os
import sys
from datetime import datetime

def load_latest_excel():
    """가장 최근 생성된 엑셀 파일 로드"""
    output_dir = 'outputs'
    excel_files = [f for f in os.listdir(output_dir) if f.endswith('.xlsx') and not f.startswith('~$') and '개선된_분석' in f]
    excel_files.sort(key=lambda x: os.path.getmtime(os.path.join(output_dir, x)), reverse=True)
    
    if not excel_files:
        print("❌ 엑셀 파일을 찾을 수 없습니다.")
        return None
    
    latest_file = os.path.join(output_dir, excel_files[0])
    print(f"📄 분석 대상 파일: {excel_files[0]}")
    return latest_file

def analyze_warehouse_inventory(excel_file):
    """창고별 재고 데이터 분석"""
    print("\n=== 창고별 재고 데이터 분석 ===")
    
    xl = pd.ExcelFile(excel_file)
    warehouse_sheets = [sheet for sheet in xl.sheet_names if sheet.startswith('창고_')]
    
    warehouse_results = {}
    
    for sheet in warehouse_sheets:
        warehouse_name = sheet.replace('창고_', '')
        print(f"\n📊 {warehouse_name} 분석:")
        
        df = pd.read_excel(excel_file, sheet_name=sheet)
        
        # 월별 데이터 분석
        monthly_data = df[df.index != '총합'] if '총합' in df.index else df
        
        print(f"  📅 분석 기간: {len(monthly_data)}개월")
        
        # 최근 12개월 데이터
        recent_12 = monthly_data.tail(12)
        
        # 입고/출고/재고 누적 계산
        total_inbound = recent_12['입고'].sum()
        total_outbound = recent_12['출고'].sum()
        current_stock = recent_12['재고'].iloc[-1] if len(recent_12) > 0 else 0
        
        # 재고 검증: 입고 - 출고 = 재고
        calculated_stock = total_inbound - total_outbound
        stock_diff = current_stock - calculated_stock
        
        warehouse_results[warehouse_name] = {
            'total_inbound': total_inbound,
            'total_outbound': total_outbound,
            'current_stock': current_stock,
            'calculated_stock': calculated_stock,
            'stock_diff': stock_diff,
            'monthly_data': monthly_data
        }
        
        print(f"    📥 총 입고: {total_inbound}")
        print(f"    📤 총 출고: {total_outbound}")
        print(f"    📦 현재 재고: {current_stock}")
        print(f"    🧮 계산된 재고: {calculated_stock}")
        print(f"    ⚠️  재고 차이: {stock_diff}")
        
        if abs(stock_diff) > 0:
            print(f"    ❌ 재고 불일치 발견!")
        else:
            print(f"    ✅ 재고 일치")
    
    return warehouse_results

def case_level_inventory_check():
    """Case별 이벤트 타임라인 재구성 및 검증"""
    print("\n=== Case별 이벤트 타임라인 분석 ===")
    
    # 원본 데이터 로드
    data_file = os.path.join('data', 'HVDC WAREHOUSE_HITACHI(HE).xlsx')
    df = pd.read_excel(data_file, sheet_name='CASE LIST')
    
    # 입고/출고 컬럼 식별
    warehouse_cols = ['DSV Indoor', 'DSV Al Markaz', 'DSV Outdoor', 'Hauler Indoor', 'DSV MZP', 'MOSB']
    site_cols = ['MIR', 'SHU', 'DAS', 'AGI']
    
    print(f"📊 총 Case 수: {len(df)}")
    print(f"🏭 창고 컬럼: {warehouse_cols}")
    print(f"🏗️  현장 컬럼: {site_cols}")
    
    # Case별 이벤트 타임라인 구성
    case_timelines = []
    warehouse_final_stock = {warehouse: 0 for warehouse in warehouse_cols}
    site_total_inbound = {site: 0 for site in site_cols}
    
    for idx, row in df.iterrows():
        case = row['Case No.']
        events = []
        
        # 입고 이벤트 수집
        for warehouse in warehouse_cols:
            if pd.notna(row[warehouse]):
                try:
                    date = pd.to_datetime(row[warehouse])
                    events.append((date, warehouse, 'warehouse_in'))
                except:
                    continue
        
        # 출고 이벤트 수집
        for site in site_cols:
            if pd.notna(row[site]):
                try:
                    date = pd.to_datetime(row[site])
                    events.append((date, site, 'site_out'))
                except:
                    continue
        
        if not events:
            continue
        
        # 시간순 정렬
        events = sorted(events, key=lambda x: x[0])
        
        # 이벤트 타임라인 구성
        timeline = []
        prev_warehouse = None
        
        for date, location, event_type in events:
            if event_type == 'warehouse_in':
                # 이전 창고에서 출고 처리
                if prev_warehouse is not None:
                    timeline.append((date, prev_warehouse, 'warehouse_out'))
                
                # 새 창고 입고
                timeline.append((date, location, 'warehouse_in'))
                prev_warehouse = location
                
            elif event_type == 'site_out':
                # 창고에서 출고
                if prev_warehouse is not None:
                    timeline.append((date, prev_warehouse, 'warehouse_out'))
                
                # 현장 입고
                timeline.append((date, location, 'site_in'))
                site_total_inbound[location] += 1
                prev_warehouse = None
        
        # 마지막 창고에 잔재고 처리
        if prev_warehouse is not None:
            timeline.append((date, prev_warehouse, 'remain_stock'))
            warehouse_final_stock[prev_warehouse] += 1
        
        case_timelines.append((case, timeline))
    
    print(f"\n📋 Case별 이벤트 분석 결과:")
    print(f"  📊 분석된 Case 수: {len(case_timelines)}")
    
    print(f"\n🏭 창고별 최종 재고 (Case별 계산):")
    for warehouse, stock in warehouse_final_stock.items():
        print(f"  {warehouse}: {stock}건")
    
    print(f"\n🏗️  현장별 총 입고 (Case별 계산):")
    for site, inbound in site_total_inbound.items():
        print(f"  {site}: {inbound}건")
    
    return case_timelines, warehouse_final_stock, site_total_inbound

def main():
    print("=== 재고 오류 진단 및 분석 시스템 ===")
    print(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 1. 최근 엑셀 파일 로드
    excel_file = load_latest_excel()
    if not excel_file:
        return
    
    # 2. 창고별 재고 데이터 분석
    warehouse_results = analyze_warehouse_inventory(excel_file)
    
    # 3. Case별 이벤트 타임라인 분석
    case_timelines, warehouse_final_stock, site_total_inbound = case_level_inventory_check()
    
    print(f"\n✅ 진단 완료!")

if __name__ == "__main__":
    main() 