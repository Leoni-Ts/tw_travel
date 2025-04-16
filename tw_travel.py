import numpy as np
import pandas as pd
import re

#county為單位：df
year =  '104'
file_path = f'data/月表2-2({year}年1月至12月中華民國國民出國_按目的地分析).xlsx'
df = pd.read_excel(file_path, skiprows=[0])
df = df.iloc[:39, [0, 1, 5]]
df = df.rename(columns={'首站抵達地\nFirst Destination': '州', 'Unnamed: 1': 'county', '104年1-12月\nJan.-Dec., 2015':'2015'})
df['州'] = df['州'].ffill()
df["county"] = df["county"].apply(lambda x: re.sub(r'[A-Za-z, .]+', '', x))

for year_int in range(105, 114):
    year_str = str(year_int)
    file_path = f'data/月表2-2({year_str}年1月至12月中華民國國民出國_按目的地分析).xlsx'

    df_year = pd.read_excel(file_path, skiprows=[0])
    df_year = df_year.iloc[:39, [5]]

    col_name = df_year.columns[0]
    new_col = re.search(r'(\d{4})', col_name).group(1)

    df_year = df_year.rename(columns={col_name: new_col})

    df = pd.concat([df, df_year], axis=1)

rows_to_drop = [17, 21, 30, 35, 38]
df.drop(index=rows_to_drop, inplace=True)
df.reset_index(drop=True, inplace=True)

#用非0資料製作另一df，把原df生成兩個NaN欄位，再將非0資料放回去
df['總成長率 (%)'] = pd.Series([np.nan] * len(df), dtype='float64')
df['CAGR (%)'] = pd.Series([np.nan] * len(df), dtype='float64')

valid_growth = df[df['2015'] != 0].copy()
valid_cagr = df[(df['2015'] > 0) & (df['2024'] > 0)].copy()
valid_growth['總成長率 (%)'] = (valid_growth['2024'] - valid_growth['2015']) / valid_growth['2015']
valid_cagr['CAGR (%)'] = (valid_cagr['2024'] / valid_cagr['2015']) ** (1/9) - 1

df.loc[valid_growth.index, '總成長率 (%)'] = valid_growth['總成長率 (%)']
df.loc[valid_cagr.index, 'CAGR (%)'] = valid_cagr['CAGR (%)']
df['總成長率 (%)'] *= 100
df['CAGR (%)'] *= 100

df.to_excel('data/tw_travel_2015_2024.xlsx', index=False)

#州為單位：df_region
year_cols = [str(year) for year in range(2015, 2025)]
df[year_cols] = df[year_cols].apply(pd.to_numeric, errors='coerce')
df_region = df.groupby('州').sum(numeric_only=True).reset_index()
df_region['總成長率 (%)'] = ((df_region['2024'] - df_region['2015']) / df_region['2015'])* 100
df_region['CAGR (%)'] = (((df_region['2024'] / df_region['2015']) ** (1/9)) - 1)* 100

df_region.to_excel('data/tw_travel_by_region.xlsx', index=False)

#df寬轉長
df_long = pd.melt(
    df,
    id_vars=['州', 'county'],
    value_vars=[str(y) for y in range(2015, 2025)],
    var_name='年度',
    value_name='人次'
)

df_long.to_excel('data/tw_travel_long_format.xlsx', index=False)
