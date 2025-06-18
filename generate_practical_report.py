#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
실무용 창고 분석 리포트 생성
- 2025-06까지 월별 데이터만 출력
- 마지막 행에 TOTAL(합계) 행 추가
- 실무에서 바로 사용 가능한 엑셀 구조
"""

import os
import sys
import pandas as pd
from datetime import datetime

# scripts 폴더의 모듈들을 import하기 위한 경로 설정
sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts'))

from corrected_warehouse_analyzer import CorrectedWarehouseAnalyzer

def add_total_row(df, target_month="2025-06"):
    """
    DataFrame에 TOTAL 행 추가
    
    Args:
        df: 원본 DataFrame
        target_month: 목표 월 (기본값: 2025-06)
        
    Returns:
        DataFrame: TOTAL 행이 추가된 DataFrame
    """
    # 현재 월까지만 필터
    df_filtered = df[df.index <= target_month].copy()
    
    if df_filtered.empty:
        return df
    
    # TOTAL 계산
    total_in = df_filtered['입고'].sum()
    total_out = df_filtered['출고'].sum()
    last_stock = df_filtered['재고'].iloc[-1] if len(df_filtered) > 0 else 0
    
    # TOTAL 행 생성
    total_row = pd.DataFrame([{
        '입고': total_in,
        '출고': total_out,
        '재고': last_stock
    }], index=['TOTAL'])
    
    # DataFrame 결합
    df_final = pd.concat([df_filtered, total_row])
    
    return df_final

def add_total_row_site(df, target_month="2025-06"):
    """
    현장용 DataFrame에 TOTAL 행 추가 (누적재고 포함)
    
    Args:
        df: 원본 DataFrame
        target_month: 목표 월 (기본값: 2025-06)
        
    Returns:
        DataFrame: TOTAL 행이 추가된 DataFrame
    """
    # 현재 월까지만 필터
    df_filtered = df[df.index <= target_month].copy()
    
    if df_filtered.empty:
        return df
    
    # TOTAL 계산
    total_in = df_filtered['입고'].sum()
    last_cumulative = df_filtered['누적재고'].iloc[-1] if len(df_filtered) > 0 else 0
    
    # TOTAL 행 생성
    total_row = pd.DataFrame([{
        '입고': total_in,
        '누적재고': last_cumulative
    }], index=['TOTAL'])
    
    # DataFrame 결합
    df_final = pd.concat([df_filtered, total_row])
    
    return df_final

def main():
    print("=== 실무용 창고 분석 리포트 생성 (2025-06까지 + TOTAL) ===")
    print(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 1. 데이터 파일 경로 설정
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(current_dir, 'data', 'HVDC WAREHOUSE_HITACHI(HE).xlsx')
    
    if not os.path.exists(excel_path):
        print(f"❌ 오류: 데이터 파일을 찾을 수 없습니다: {excel_path}")
        return
    
    print(f"📁 데이터 파일: {excel_path}")
    
    try:
        # 2. 정확한 분석기 초기화
        print("\n🔍 정확한 분석기 초기화 중...")
        analyzer = CorrectedWarehouseAnalyzer(excel_path, sheet_name='CASE LIST')
        
        # 3. Case별 이벤트 기반으로 월별 집계
        print("📊 Case별 이벤트 기반 월별 집계 중...")
        result = analyzer.generate_corrected_report(
            start_date='2023-01-01', 
            end_date='2025-12-31'
        )
        
        # 4. 결과 추출
        warehouse_stock = result['warehouse_stock']
        site_stock = result['site_stock']
        dead_stock = result['dead_stock']
        
        print(f"✅ 분석 완료!")
        
        # 5. 실무용 엑셀 파일로 저장
        print("\n💾 실무용 엑셀 파일 저장 중...")
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_dir = os.path.join(current_dir, 'outputs')
        os.makedirs(output_dir, exist_ok=True)
        
        output_file = os.path.join(output_dir, f'실무용_창고분석_{timestamp}_202506까지_TOTAL.xlsx')
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 창고별 시트 생성 (TOTAL 행 포함)
            print("  📋 창고별 시트 생성 중...")
            for warehouse, df in warehouse_stock.items():
                df_with_total = add_total_row(df, "2025-06")
                df_with_total.to_excel(writer, sheet_name=f'창고_{warehouse}')
                print(f"    ✅ 창고_{warehouse} (TOTAL 포함)")
            
            # 현장별 시트 생성 (TOTAL 행 포함)
            print("  📋 현장별 시트 생성 중...")
            for site, df in site_stock.items():
                df_with_total = add_total_row_site(df, "2025-06")
                df_with_total.to_excel(writer, sheet_name=f'Site_{site}')
                print(f"    ✅ Site_{site} (TOTAL 포함)")
            
            # Dead Stock 시트 생성
            if len(dead_stock) > 0:
                print("  📋 Dead Stock 시트 생성 중...")
                dead_stock.to_excel(writer, sheet_name='DeadStock_90일+', index=False)
                print(f"    ✅ DeadStock_90일+")
            
            # 실무용 요약 시트 생성
            print("  📋 실무용 요약 시트 생성 중...")
            summary_data = []
            
            # 창고별 최종 재고 (TOTAL 행에서 추출)
            for warehouse, df in warehouse_stock.items():
                df_with_total = add_total_row(df, "2025-06")
                if 'TOTAL' in df_with_total.index:
                    total_row = df_with_total.loc['TOTAL']
                    summary_data.append({
                        '구분': f'창고_{warehouse}',
                        '총입고': total_row['입고'],
                        '총출고': total_row['출고'],
                        '현재재고': total_row['재고'],
                        '유형': '창고'
                    })
            
            # 현장별 최종 누적입고 (TOTAL 행에서 추출)
            for site, df in site_stock.items():
                df_with_total = add_total_row_site(df, "2025-06")
                if 'TOTAL' in df_with_total.index:
                    total_row = df_with_total.loc['TOTAL']
                    summary_data.append({
                        '구분': f'현장_{site}',
                        '총입고': total_row['입고'],
                        '총출고': 0,  # 현장은 출고 없음
                        '현재재고': total_row['누적재고'],
                        '유형': '현장'
                    })
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='실무요약', index=False)
            print("    ✅ 실무요약")
        
        print(f"\n✅ 실무용 엑셀 파일 생성 완료!")
        print(f"📄 파일명: {os.path.basename(output_file)}")
        print(f"📄 파일 크기: {os.path.getsize(output_file) / 1024:.1f} KB")
        
        # 6. 실무용 검증 정보 출력
        print(f"\n🔍 실무용 검증 정보 (2025-06까지):")
        
        # 창고별 TOTAL 정보 출력
        print(f"\n🏭 창고별 TOTAL 정보:")
        for warehouse, df in warehouse_stock.items():
            df_with_total = add_total_row(df, "2025-06")
            if 'TOTAL' in df_with_total.index:
                total_row = df_with_total.loc['TOTAL']
                print(f"  {warehouse}:")
                print(f"    - 총입고: {total_row['입고']:,}건")
                print(f"    - 총출고: {total_row['출고']:,}건")
                print(f"    - 현재재고: {total_row['재고']:,}건")
        
        # 현장별 TOTAL 정보 출력
        print(f"\n🏗️  현장별 TOTAL 정보:")
        for site, df in site_stock.items():
            df_with_total = add_total_row_site(df, "2025-06")
            if 'TOTAL' in df_with_total.index:
                total_row = df_with_total.loc['TOTAL']
                print(f"  {site}:")
                print(f"    - 총입고: {total_row['입고']:,}건")
                print(f"    - 누적재고: {total_row['누적재고']:,}건")
        
        # 7. 엑셀 파일 자동 열기
        try:
            os.startfile(output_file)
            print(f"\n🔓 엑셀 파일을 자동으로 열었습니다.")
        except:
            print(f"\n💡 엑셀 파일을 수동으로 열어주세요: {output_file}")
        
        print(f"\n📋 실무용 시트 구조:")
        print("  - 창고별 시트: 2025-06까지 월별 데이터 + TOTAL 행")
        print("  - 현장별 시트: 2025-06까지 월별 데이터 + TOTAL 행")
        print("  - DeadStock_90일+: 90일 이상 미출고 Case 목록")
        print("  - 실무요약: 창고/현장별 TOTAL 정보 요약")
        
        print(f"\n✅ 실무용 표 구조:")
        print("  | 월 | 입고 | 출고 | 재고 |")
        print("  |----|------|------|------|")
        print("  | 2023-01 | 0 | 0 | 0 |")
        print("  | ... | ... | ... | ... |")
        print("  | 2025-06 | 0 | 69 | 414 |")
        print("  | **TOTAL** | **1132** | **940** | **414** |")
        
        return output_file
        
    except Exception as e:
        print(f"\n❌ 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    main() 