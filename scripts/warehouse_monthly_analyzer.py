import pandas as pd
from pandas.tseries.offsets import MonthEnd

class WarehouseMonthlyAnalyzer:
    def __init__(self, excel_path, sheet_name='Sheet1'):
        self.df = pd.read_excel(excel_path, sheet_name=sheet_name)
        self.warehouse_cols = ['DSV Outdoor', 'DSV Indoor', 'DSV Al Markaz', 'Hauler Indoor', 'DSV MZP', 'MOSB']
        self.site_cols = ['DAS', 'MIR', 'SHU', 'AGI']
        self._preprocess()
        
    def _preprocess(self):
        for col in self.warehouse_cols + self.site_cols:
            self.df[col] = pd.to_datetime(self.df[col], errors='coerce')
        self.df['Case No.'] = self.df['Case No.'].astype(str)
    
    def get_month_list(self, start='2023-01-01', end='2025-06-30'):
        return pd.date_range(start=start, end=end, freq='M')
    
    def warehouse_monthly_inout_stock(self):
        """창고별 월별 입고/출고/재고 DataFrame 리턴"""
        months = self.get_month_list()
        result = {}
        for wh in self.warehouse_cols:
            temp = self.df[self.df[wh].notna()].copy()
            temp['입고월'] = temp[wh].dt.to_period('M').dt.to_timestamp() + MonthEnd(0)
            in_monthly = temp.groupby('입고월').size().reindex(months, fill_value=0)
            temp['출고여부'] = temp[self.site_cols].notna().any(axis=1)
            out_monthly = temp[temp['출고여부']].groupby('입고월').size().reindex(months, fill_value=0)
            stock = in_monthly.cumsum() - out_monthly.cumsum()
            result[wh] = pd.DataFrame({'입고': in_monthly, '출고': out_monthly, '재고': stock})
        return result
    
    def site_monthly_in_stock(self):
        """현장별 월별 입고, 누적 재고(=누적 입고) DataFrame 리턴"""
        months = self.get_month_list()
        result = {}
        for site in self.site_cols:
            temp = self.df[self.df[site].notna()].copy()
            temp['입고월'] = temp[site].dt.to_period('M').dt.to_timestamp() + MonthEnd(0)
            in_monthly = temp.groupby('입고월').size().reindex(months, fill_value=0)
            stock = in_monthly.cumsum()
            result[site] = pd.DataFrame({'입고': in_monthly, '누적재고': stock})
        return result 