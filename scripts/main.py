from warehouse_analyzer import WarehouseAnalyzer
from warehouse_monthly_analyzer import WarehouseMonthlyAnalyzer
from improved_warehouse_analyzer import ImprovedWarehouseAnalyzer
import os
import pandas as pd

if __name__ == "__main__":
    print("=== HVDC Warehouse 분석 시스템 ===")
    print("1. 기존 분석 (기본)")
    print("2. 개선된 분석 (정확한 이벤트 추적)")
    print()
    
    # 개선된 분석 실행
    print("=== 개선된 분석 실행 ===")
    improved_analyzer = ImprovedWarehouseAnalyzer('C:/WAREHOUSE/warehouse_analytics/data/HVDC WAREHOUSE_HITACHI(HE).xlsx', sheet_name='CASE LIST')
    
    # 종합 리포트 생성
    results = improved_analyzer.generate_comprehensive_report(start_date='2023-01-01', end_date='2025-12-31')
    
    # 엑셀 저장
    output_dir = os.path.join(os.path.dirname(__file__), '..', 'outputs')
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    
    def add_total_row(df, label='총합'):
        """총합 행 추가"""
        sums = df.sum(numeric_only=True)
        total_row = pd.DataFrame([sums], index=[label])
        for col in df.columns:
            if col not in sums.index:
                total_row[col] = ''
        return pd.concat([df, total_row], axis=0)
    
    def format_index_to_ymd(df):
        """날짜 포맷을 yyyy-mm-dd로 변환"""
        idx = df.index
        if isinstance(idx, pd.DatetimeIndex):
            idx = idx.strftime('%Y-%m-%d')
        idx = [str(i) for i in idx]
        df.index = idx
        return df
    
    # 개선된 분석 결과를 엑셀로 저장
    output_excel = os.path.join(output_dir, '개선된_월별_창고_현장_입출고재고_집계.xlsx')
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        # 창고별 월별 입출고/재고
        for warehouse, stock_df in results['warehouse_stock'].items():
            df_with_total = add_total_row(stock_df)
            df_with_total = format_index_to_ymd(df_with_total)
            df_with_total.to_excel(writer, sheet_name=f'창고_{warehouse}')
        
        # 현장별 월별 입고/누적재고
        for site, stock_df in results['site_stock'].items():
            df_with_total = add_total_row(stock_df)
            df_with_total = format_index_to_ymd(df_with_total)
            df_with_total.to_excel(writer, sheet_name=f'Site_{site}')
        
        # Dead Stock 분석
        if len(results['dead_stock']) > 0:
            dead_stock_formatted = results['dead_stock'].copy()
            dead_stock_formatted['마지막입고일'] = dead_stock_formatted['마지막입고일'].dt.strftime('%Y-%m-%d')
            add_total_row(dead_stock_formatted).to_excel(writer, sheet_name='DeadStock_90일+', index=False)
        
        # 요약 정보
        summary_data = []
        for warehouse, stock_df in results['warehouse_stock'].items():
            recent_12 = stock_df.tail(12)
            summary_data.append({
                '구분': f'창고_{warehouse}',
                '최근12개월_입고': recent_12['입고'].sum(),
                '최근12개월_출고': recent_12['출고'].sum(),
                '현재재고': recent_12['재고'].iloc[-1]
            })
        
        for site, stock_df in results['site_stock'].items():
            recent_12 = stock_df.tail(12)
            summary_data.append({
                '구분': f'Site_{site}',
                '최근12개월_입고': recent_12['입고'].sum(),
                '최근12개월_출고': 0,  # 현장은 출고 없음
                '현재재고': recent_12['누적재고'].iloc[-1]
            })
        
        summary_df = pd.DataFrame(summary_data)
        add_total_row(summary_df).to_excel(writer, sheet_name='요약', index=False)
    
    print(f"\n=== 분석 완료 ===")
    print(f"결과 파일: {output_excel}")
    print("엑셀 파일에 다음 시트들이 생성되었습니다:")
    print("- 창고별 월별 입출고/재고 (각 창고별)")
    print("- 현장별 월별 입고/누적재고 (각 현장별)")
    print("- Dead Stock 분석 (90일 이상 미출고)")
    print("- 요약 정보 (창고/현장별 최근 12개월 요약)") 