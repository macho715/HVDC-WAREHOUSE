import pandas as pd
from pandas.tseries.offsets import MonthEnd
from datetime import datetime

class WarehouseAnalyzer:
    def __init__(self, excel_path, sheet_name='Sheet1'):
        self.df = pd.read_excel(excel_path, sheet_name=sheet_name)
        self.warehouse_cols = ['DSV Outdoor', 'DSV Indoor', 'DSV Al Markaz', 'Hauler Indoor', 'DSV MZP', 'MOSB']
        self.site_cols = ['DAS', 'MIR', 'SHU', 'AGI']
        self._preprocess()
        
    def _preprocess(self):
        for col in self.warehouse_cols + self.site_cols:
            self.df[col] = pd.to_datetime(self.df[col], errors='coerce')
        self.df['Site'] = self.df['Site'].replace({'Das': 'DAS'})
        self.df['Case No.'] = self.df['Case No.'].astype(str)
        self.df['Storage_Type'] = self.df['Storage'].apply(self.classify_storage)

    @staticmethod
    def classify_storage(val):
        val = str(val).lower()
        if 'indoor' in val:
            return 'Indoor'
        elif 'covered' in val:
            return 'Outdoor Covered'
        elif 'open' in val:
            return 'Outdoor Open'
        else:
            return '기타'
    
    def calculate_inout_dates(self):
        self.df['입고일'] = self.df[self.warehouse_cols].min(axis=1)
        self.df['출고일'] = self.df[self.site_cols].max(axis=1)
        self.df['입고월'] = self.df['입고일'].dt.to_period('M').dt.to_timestamp() + MonthEnd(0)
        self.df['출고월'] = self.df['출고일'].dt.to_period('M').dt.to_timestamp() + MonthEnd(0)
        
    def get_monthly_summary(self, start='2023-01-01', end='2025-06-30'):
        self.calculate_inout_dates()
        full_months = pd.date_range(start=start, end=end, freq='M')
        inbound = self.df.groupby('입고월').size().reindex(full_months, fill_value=0)
        outbound = self.df.groupby('출고월').size().reindex(full_months, fill_value=0)
        stock = inbound.cumsum() - outbound.cumsum()
        return pd.DataFrame({'입고총합': inbound, '출고총합': outbound, '재고': stock})
    
    def get_dead_stock(self, days=90):
        self.calculate_inout_dates()
        today = pd.Timestamp(datetime.today())
        self.df['출고여부'] = self.df[self.site_cols].notna().any(axis=1)
        self.df['입고후경과'] = (today - self.df['입고일']).dt.days
        return self.df[(self.df['출고여부'] == False) & (self.df['입고후경과'] > days)]
    
    def get_kpi(self):
        total_cases = self.df['Case No.'].nunique()
        kpi_list = []
        for site in self.site_cols:
            temp = self.df[self.df[site].notna()].copy()
            reached = temp['Case No.'].nunique()
            reach_rate = round((reached / total_cases) * 100, 2)
            temp['입고일'] = self.df[self.warehouse_cols].min(axis=1)
            temp['리드타임(일)'] = (temp[site] - temp['입고일']).dt.days
            avg_leadtime = round(temp['리드타임(일)'].mean(), 1)
            kpi_list.append({'Site': site, '도달건수': reached, '도달률(%)': reach_rate, '평균 리드타임(일)': avg_leadtime})
        return pd.DataFrame(kpi_list)
    
    def get_warehouse_turnover(self, wh_col):
        self.calculate_inout_dates()
        df_temp = self.df[self.df[wh_col].notna()].copy()
        df_temp['입고월'] = df_temp[wh_col].dt.to_period('M').dt.to_timestamp() + MonthEnd(0)
        df_temp['출고'] = df_temp[self.site_cols].notna().any(axis=1)
        monthly_in = df_temp.groupby('입고월').size()
        monthly_out = df_temp[df_temp['출고']].groupby('입고월').size()
        turnover = (monthly_out / monthly_in).fillna(0).clip(upper=1)
        return pd.DataFrame({'입고': monthly_in, '출고': monthly_out.reindex(monthly_in.index, fill_value=0), '회전율': turnover}) 