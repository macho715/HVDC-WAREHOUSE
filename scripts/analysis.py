import pandas as pd
from pandas.tseries.offsets import MonthEnd

def monthly_summary(df, warehouse_cols, site_cols, start='2023-01-01', end='2025-06-30'):
    df['입고일'] = df[warehouse_cols].min(axis=1)
    df['출고일'] = df[site_cols].max(axis=1)
    df['입고월'] = df['입고일'].dt.to_period('M').dt.to_timestamp() + MonthEnd(0)
    df['출고월'] = df['출고일'].dt.to_period('M').dt.to_timestamp() + MonthEnd(0)
    full_months = pd.date_range(start=start, end=end, freq='M')
    inbound_all = df.groupby('입고월').size().reindex(full_months, fill_value=0)
    outbound_all = df.groupby('출고월').size().reindex(full_months, fill_value=0)
    stock_all = inbound_all.cumsum() - outbound_all.cumsum()
    return pd.DataFrame({'입고총합': inbound_all, '출고총합': outbound_all, '재고': stock_all})

def site_kpi(df, warehouse_cols, site_cols):
    total_cases = df['Case No.'].nunique()
    site_kpi_list = []
    for site in site_cols:
        temp = df[df[site].notna()].copy()
        reached = temp['Case No.'].nunique()
        reach_rate = round((reached / total_cases) * 100, 2)
        temp['입고일'] = df[warehouse_cols].min(axis=1)
        temp['리드타임(일)'] = (temp[site] - temp['입고일']).dt.days
        avg_leadtime = round(temp['리드타임(일)'].mean(), 1)
        site_kpi_list.append({'Site': site, '도달건수': reached, '도달률(%)': reach_rate, '평균 리드타임(일)': avg_leadtime})
    return pd.DataFrame(site_kpi_list) 