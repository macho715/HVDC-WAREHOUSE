from warehouse_analyzer import WarehouseAnalyzer
from warehouse_monthly_analyzer import WarehouseMonthlyAnalyzer
import os
import pandas as pd

if __name__ == "__main__":
    analyzer = WarehouseAnalyzer('C:/WAREHOUSE/warehouse_analytics/data/HVDC WAREHOUSE_HITACHI(HE).xlsx', sheet_name='CASE LIST')
    
    # 1. 월별 입출고/재고 집계표
    monthly_summary = analyzer.get_monthly_summary()
    print("월별 입출고/재고 요약:")
    print(monthly_summary.tail())

    # 2. Dead Stock(90일 이상 미출고)
    dead_stock_df = analyzer.get_dead_stock(days=90)
    print("\nDead Stock (90일 이상 미출고):")
    print(dead_stock_df[['Case No.', '입고일', '입고후경과', 'Storage_Type']])

    # 3. KPI(도달률, 평균 리드타임)
    site_kpi_df = analyzer.get_kpi()
    print("\nSite별 KPI:")
    print(site_kpi_df)

    # 4. 창고별 회전율 예시 (DSV Indoor)
    indoor_turnover = analyzer.get_warehouse_turnover('DSV Indoor')
    print("\nDSV Indoor 창고별 회전율:")
    print(indoor_turnover)

    # === 창고/현장별 월별 입출고/재고 자동 집계 및 하나의 엑셀 파일로 저장 ===
    output_dir = os.path.join(os.path.dirname(__file__), '..', 'outputs')
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    monthly_analyzer = WarehouseMonthlyAnalyzer('C:/WAREHOUSE/warehouse_analytics/data/HVDC WAREHOUSE_HITACHI(HE).xlsx', sheet_name='CASE LIST')

    wh_result = monthly_analyzer.warehouse_monthly_inout_stock()
    site_result = monthly_analyzer.site_monthly_in_stock()

    def add_total_row(df, label='총합'):
        # 합계 row를 DataFrame 마지막에 추가
        sums = df.sum(numeric_only=True)
        total_row = pd.DataFrame([sums], index=[label])
        for col in df.columns:
            if col not in sums.index:
                total_row[col] = ''
        return pd.concat([df, total_row], axis=0)

    def format_index_to_ymd(df):
        # DatetimeIndex를 'yyyy-mm-dd' 문자열로 변환 (총합 행은 그대로)
        idx = df.index
        if isinstance(idx, pd.DatetimeIndex):
            # 월말 날짜를 해당 월의 마지막 날로 변환
            idx = idx.strftime('%Y-%m-%d')
        idx = [str(i) for i in idx]
        df.index = idx
        return df

    output_excel = os.path.join(output_dir, '월별_창고_현장_입출고재고_집계.xlsx')
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        # 창고별 집계 시트
        for wh, df in wh_result.items():
            df_with_total = add_total_row(df)
            df_with_total = format_index_to_ymd(df_with_total)
            df_with_total.to_excel(writer, sheet_name=f'창고_{wh}')
        # 현장별 집계 시트
        for site, df in site_result.items():
            df_with_total = add_total_row(df)
            df_with_total = format_index_to_ymd(df_with_total)
            df_with_total.to_excel(writer, sheet_name=f'Site_{site}')
        # 요약, Dead Stock, KPI, 회전율도 추가 저장 (합계 적용)
        format_index_to_ymd(add_total_row(monthly_summary)).to_excel(writer, sheet_name='전체_월별요약')
        format_index_to_ymd(add_total_row(dead_stock_df)).to_excel(writer, sheet_name='DeadStock_90일+')
        format_index_to_ymd(add_total_row(site_kpi_df)).to_excel(writer, sheet_name='Site별KPI')
        format_index_to_ymd(add_total_row(indoor_turnover)).to_excel(writer, sheet_name='DSVIndoor_회전율') 