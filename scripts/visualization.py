import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd

def plot_leadtime_distribution(df, warehouse_cols, site_cols):
    leadtime_all = []
    for site in site_cols:
        temp = df[df[site].notna()].copy()
        temp['입고일'] = df[warehouse_cols].min(axis=1)
        temp['리드타임(일)'] = (temp[site] - temp['입고일']).dt.days
        temp['Site'] = site
        leadtime_all.append(temp[['리드타임(일)', 'Site']])
    leadtime_df = pd.concat(leadtime_all).dropna()
    plt.figure(figsize=(10, 6))
    sns.histplot(data=leadtime_df, x='리드타임(일)', hue='Site', bins=30, kde=True, multiple="stack")
    plt.title('Site별 리드타임(일) 분포')
    plt.xlabel('리드타임 (일)')
    plt.ylabel('자재 수량')
    plt.grid(True)
    plt.tight_layout()
    plt.show()

def plot_site_kpi(site_kpi_df):
    fig, ax = plt.subplots(figsize=(10, 6))
    bars = ax.bar(site_kpi_df['Site'], site_kpi_df['도달률(%)'], color='skyblue', label='도달률 (%)')
    ax2 = ax.twinx()
    ax2.plot(site_kpi_df['Site'], site_kpi_df['평균 리드타임(일)'], color='orange', marker='o', label='평균 리드타임(일)')
    plt.title('Site별 도달률 & 리드타임')
    fig.tight_layout()
    plt.show() 