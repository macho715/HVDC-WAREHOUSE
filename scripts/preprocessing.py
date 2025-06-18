import pandas as pd
from pandas.tseries.offsets import MonthEnd
import os

def load_and_clean_data(excel_path):
    warehouse_cols = ['DSV Outdoor', 'DSV Indoor', 'DSV Al Markaz', 'Hauler Indoor', 'DSV MZP', 'MOSB']
    site_cols = ['DAS', 'MIR', 'SHU', 'AGI']
    df = pd.read_excel(excel_path, sheet_name='Sheet1')
    for col in warehouse_cols + site_cols:
        df[col] = pd.to_datetime(df[col], errors='coerce')
    df['Site'] = df['Site'].replace({'Das': 'DAS'})
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
    df['Storage_Type'] = df['Storage'].apply(classify_storage)
    df = df[df['Case No.'].notna()]
    df['Case No.'] = df['Case No.'].astype(str)
    return df

if __name__ == "__main__":
    df = load_and_clean_data("../data/HVDC WAREHOUSE_HITACHI(HE).xlsx")
    print(df.head())
    output_dir = os.path.join(os.path.dirname(__file__), '..', 'outputs')
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True) 