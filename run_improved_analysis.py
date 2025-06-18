#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC Warehouse 개선된 분석 실행 스크립트
- 정확한 이벤트 추적 기반 월별 입출고/재고 집계
- 자동 엑셀 파일 생성
"""

import os
import sys
import pandas as pd
from datetime import datetime

# scripts 폴더의 모듈들을 import하기 위한 경로 설정
sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts'))

from improved_warehouse_analyzer import ImprovedWarehouseAnalyzer

def main():
    print("=== HVDC Warehouse 개선된 분석 시스템 ===")
    print(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    # 현재 디렉토리 기준으로 상대 경로 설정
    current_dir = os.path.dirname(os.path.abspath(__file__))
    data_file = os.path.join(current_dir, 'data', 'HVDC WAREHOUSE_HITACHI(HE).xlsx')
    
    # 데이터 파일 존재 확인
    if not os.path.exists(data_file):
        print(f"❌ 오류: 데이터 파일을 찾을 수 없습니다: {data_file}")
        print("데이터 파일이 'data' 폴더에 있는지 확인해주세요.")
        return
    
    print(f"📁 데이터 파일: {data_file}")
    print(f"📁 현재 작업 디렉토리: {current_dir}")
    
    try:
        # 개선된 분석기 초기화
        print("\n🔍 개선된 분석기 초기화 중...")
        improved_analyzer = ImprovedWarehouseAnalyzer(data_file, sheet_name='CASE LIST')
        
        # 종합 리포트 생성
        print("📊 종합 리포트 생성 중...")
        results = improved_analyzer.generate_comprehensive_report(
            start_date='2023-01-01', 
            end_date='2025-12-31'
        )
        
        # outputs 폴더 생성
        output_dir = os.path.join(current_dir, 'outputs')
        os.makedirs(output_dir, exist_ok=True)
        print(f"📁 출력 폴더: {output_dir}")
        
        def add_total_row(df, label='총합'):
            """총합 행 추가"""
            if df.empty:
                return df
            sums = df.sum(numeric_only=True)
            total_row = pd.DataFrame([sums], index=[label])
            for col in df.columns:
                if col not in sums.index:
                    total_row[col] = ''
            return pd.concat([df, total_row], axis=0)
        
        def format_index_to_ymd(df):
            """날짜 포맷을 yyyy-mm-dd로 변환"""
            if df.empty:
                return df
            idx = df.index
            if isinstance(idx, pd.DatetimeIndex):
                idx = idx.strftime('%Y-%m-%d')
            idx = [str(i) for i in idx]
            df.index = idx
            return df
        
        # 타임스탬프가 포함된 파일명 생성
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_excel = os.path.join(output_dir, f'개선된_분석_{timestamp}_월별_창고_현장_입출고재고_집계.xlsx')
        
        print(f"💾 엑셀 파일 저장 중: {os.path.basename(output_excel)}")
        
        # 개선된 분석 결과를 엑셀로 저장
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            # 창고별 월별 입출고/재고
            print("  📋 창고별 시트 생성 중...")
            for warehouse, stock_df in results['warehouse_stock'].items():
                df_with_total = add_total_row(stock_df)
                df_with_total = format_index_to_ymd(df_with_total)
                df_with_total.to_excel(writer, sheet_name=f'창고_{warehouse}')
                print(f"    ✅ 창고_{warehouse}")
            
            # 현장별 월별 입고/누적재고
            print("  📋 현장별 시트 생성 중...")
            for site, stock_df in results['site_stock'].items():
                df_with_total = add_total_row(stock_df)
                df_with_total = format_index_to_ymd(df_with_total)
                df_with_total.to_excel(writer, sheet_name=f'Site_{site}')
                print(f"    ✅ Site_{site}")
            
            # Dead Stock 분석
            if len(results['dead_stock']) > 0:
                print("  📋 Dead Stock 시트 생성 중...")
                dead_stock_formatted = results['dead_stock'].copy()
                if '마지막입고일' in dead_stock_formatted.columns:
                    dead_stock_formatted['마지막입고일'] = dead_stock_formatted['마지막입고일'].dt.strftime('%Y-%m-%d')
                add_total_row(dead_stock_formatted).to_excel(writer, sheet_name='DeadStock_90일+', index=False)
                print("    ✅ DeadStock_90일+")
            
            # 요약 정보
            print("  📋 요약 시트 생성 중...")
            summary_data = []
            for warehouse, stock_df in results['warehouse_stock'].items():
                if not stock_df.empty:
                    recent_12 = stock_df.tail(12)
                    summary_data.append({
                        '구분': f'창고_{warehouse}',
                        '최근12개월_입고': recent_12['입고'].sum(),
                        '최근12개월_출고': recent_12['출고'].sum(),
                        '현재재고': recent_12['재고'].iloc[-1] if len(recent_12) > 0 else 0
                    })
            
            for site, stock_df in results['site_stock'].items():
                if not stock_df.empty:
                    recent_12 = stock_df.tail(12)
                    summary_data.append({
                        '구분': f'Site_{site}',
                        '최근12개월_입고': recent_12['입고'].sum(),
                        '최근12개월_출고': 0,  # 현장은 출고 없음
                        '현재재고': recent_12['누적재고'].iloc[-1] if len(recent_12) > 0 else 0
                    })
            
            summary_df = pd.DataFrame(summary_data)
            add_total_row(summary_df).to_excel(writer, sheet_name='요약', index=False)
            print("    ✅ 요약")
        
        print(f"\n✅ 분석 완료!")
        print(f"📄 결과 파일: {output_excel}")
        print(f"📄 파일 크기: {os.path.getsize(output_excel) / 1024:.1f} KB")
        
        # 엑셀 파일 자동 열기
        try:
            os.startfile(output_excel)
            print("🔓 엑셀 파일을 자동으로 열었습니다.")
        except:
            print(f"💡 엑셀 파일을 수동으로 열어주세요: {output_excel}")
        
        print("\n📋 생성된 시트 목록:")
        print("- 창고별 월별 입출고/재고 (각 창고별)")
        print("- 현장별 월별 입고/누적재고 (각 현장별)")
        print("- Dead Stock 분석 (90일 이상 미출고)")
        print("- 요약 정보 (창고/현장별 최근 12개월 요약)")
        
        return output_excel
        
    except Exception as e:
        print(f"\n❌ 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    main() 