#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
실전 Case별 이벤트 기반 월별 창고별/현장별 재고 집계 + 엑셀 저장
- Case별 실제 마지막 위치 기준으로 월별 잔재고 집계
- 이중 카운트 방지
- 정확한 재고 산출
"""

import pandas as pd
import os
from datetime import datetime
from pandas.tseries.offsets import MonthEnd

def main():
    print("=== 실전 Case별 이벤트 기반 재고 집계 시스템 ===")
    print(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 1. 데이터 불러오기
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(current_dir, 'data', 'HVDC WAREHOUSE_HITACHI(HE).xlsx')
    
    if not os.path.exists(excel_path):
        print(f"❌ 오류: 데이터 파일을 찾을 수 없습니다: {excel_path}")
        return
    
    print(f"📁 데이터 파일: {excel_path}")
    df = pd.read_excel(excel_path, sheet_name='CASE LIST')
    print(f"📊 총 Case 수: {len(df)}")
    
    # 창고 및 현장 컬럼 정의
    warehouse_cols = ['DSV Outdoor', 'DSV Indoor', 'DSV Al Markaz', 'Hauler Indoor', 'DSV MZP', 'MOSB']
    site_cols = ['DAS', 'MIR', 'SHU', 'AGI']
    
    print(f"🏭 창고 컬럼: {warehouse_cols}")
    print(f"🏗️  현장 컬럼: {site_cols}")
    
    # 날짜 컬럼 변환
    for col in warehouse_cols + site_cols:
        df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # 2. 월별 집계용 month list (입고/출고가 일어난 월 전체)
    print("\n📅 월별 집계 기간 계산 중...")
    all_months = set()
    for col in warehouse_cols + site_cols:
        all_months |= set(df[col].dropna().dt.to_period('M'))
    month_list = sorted(all_months)
    month_strs = [str(m) for m in month_list]
    
    print(f"📅 분석 기간: {month_strs[0]} ~ {month_strs[-1]} ({len(month_strs)}개월)")
    
    # 3. Case별로 이벤트 타임라인 추적 및 마지막 창고/현장 위치 기록
    print("\n🔍 Case별 이벤트 타임라인 분석 중...")
    case_status = []
    processed_cases = 0
    
    for idx, row in df.iterrows():
        case = row['Case No.']
        events = []
        
        # 입고 이벤트 수집
        for w in warehouse_cols:
            if pd.notna(row[w]):
                events.append((row[w], w, 'in'))
        
        # 출고 이벤트 수집
        for s in site_cols:
            if pd.notna(row[s]):
                events.append((row[s], s, 'out'))
        
        if not events:
            continue
        
        events.sort(key=lambda x: x[0])
        last_date, last_loc, last_type = events[-1]
        last_month = str(pd.to_datetime(last_date).to_period('M'))
        
        if last_type == 'in':  # 출고 안 된 재고(잔존)
            case_status.append({'case': case, 'loc': last_loc, 'type': 'warehouse', 'month': last_month})
        elif last_type == 'out':
            case_status.append({'case': case, 'loc': last_loc, 'type': 'site', 'month': last_month})
        
        processed_cases += 1
        if processed_cases % 1000 == 0:
            print(f"  📊 처리된 Case: {processed_cases}/{len(df)}")
    
    print(f"✅ 총 {len(case_status)}개 Case 이벤트 분석 완료")
    
    # 4. 월별/창고별 재고 테이블, 월별/현장별 누적입고 테이블 생성
    print("\n📊 월별 재고 집계 중...")
    
    # 창고: 해당 월까지 남은 Case수 (누적)
    warehouse_stock_table = {w: [] for w in warehouse_cols}
    for m in month_strs:
        # 월별 잔존 Case 카운트 (해당 월까지 출고되지 않은 케이스)
        for w in warehouse_cols:
            cnt = sum((s['loc'] == w and s['type'] == 'warehouse' and s['month'] <= m) for s in case_status)
            warehouse_stock_table[w].append(cnt)
    
    warehouse_df = pd.DataFrame({'월': month_strs})
    for w in warehouse_cols:
        warehouse_df[w] = warehouse_stock_table[w]
    
    # 현장: 해당 월까지 누적 도달 Case수 (누적)
    site_stock_table = {s: [] for s in site_cols}
    for m in month_strs:
        for s in site_cols:
            cnt = sum((st['loc'] == s and st['type'] == 'site' and st['month'] <= m) for st in case_status)
            site_stock_table[s].append(cnt)
    
    site_df = pd.DataFrame({'월': month_strs})
    for s in site_cols:
        site_df[s] = site_stock_table[s]
    
    # 5. 엑셀로 저장
    print("\n💾 엑셀 파일 저장 중...")
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_dir = os.path.join(current_dir, 'outputs')
    os.makedirs(output_dir, exist_ok=True)
    
    output_file = os.path.join(output_dir, f'정확재고_케이스별월별_{timestamp}_창고_현장.xlsx')
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        warehouse_df.to_excel(writer, sheet_name='창고별_월별재고', index=False)
        site_df.to_excel(writer, sheet_name='현장별_월별누적입고', index=False)
        
        # 요약 정보 추가
        summary_data = []
        
        # 창고별 최종 재고
        for w in warehouse_cols:
            final_stock = warehouse_df[w].iloc[-1] if len(warehouse_df) > 0 else 0
            summary_data.append({
                '구분': f'창고_{w}',
                '최종재고': final_stock,
                '유형': '창고'
            })
        
        # 현장별 최종 누적입고
        for s in site_cols:
            final_inbound = site_df[s].iloc[-1] if len(site_df) > 0 else 0
            summary_data.append({
                '구분': f'현장_{s}',
                '최종재고': final_inbound,
                '유형': '현장'
            })
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='요약', index=False)
    
    print(f"✅ 엑셀 파일 생성 완료: {os.path.basename(output_file)}")
    print(f"📄 파일 크기: {os.path.getsize(output_file) / 1024:.1f} KB")
    
    # 6. 결과 요약 출력
    print(f"\n📋 생성된 시트:")
    print("  - 창고별_월별재고: 각 창고별 해당 월 잔재고 (실제 Case별 기준)")
    print("  - 현장별_월별누적입고: 각 현장별 해당 월 누적 입고량 (Case별)")
    print("  - 요약: 창고/현장별 최종 재고/누적입고")
    
    print(f"\n🏭 창고별 최종 재고:")
    for w in warehouse_cols:
        final_stock = warehouse_df[w].iloc[-1] if len(warehouse_df) > 0 else 0
        print(f"  {w}: {final_stock}건")
    
    print(f"\n🏗️  현장별 최종 누적입고:")
    for s in site_cols:
        final_inbound = site_df[s].iloc[-1] if len(site_df) > 0 else 0
        print(f"  {s}: {final_inbound}건")
    
    # 7. 엑셀 파일 자동 열기
    try:
        os.startfile(output_file)
        print(f"\n🔓 엑셀 파일을 자동으로 열었습니다.")
    except:
        print(f"\n💡 엑셀 파일을 수동으로 열어주세요: {output_file}")
    
    return output_file

if __name__ == "__main__":
    main() 