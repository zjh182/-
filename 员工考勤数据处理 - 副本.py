import pandas as pd
from dateutil.parser import parse
import warnings

warnings.filterwarnings("ignore")

path = r"C:考勤报表.xlsx"# 原始数据路径
result_path = r"C:\1\din\员工考勤_res.xlsx" # 结果路径
data = pd.read_excel(path, sheet_name='每日统计', skiprows=2, header=[0,1])

name = data['姓名']
name.columns = ['姓名']

punch_in = data['上班1'][['打卡时间']]
punch_in.columns = ['早班打卡']

punch_out = data['下班1'][['打卡时间']]
punch_out.columns = ['中班打卡']

data_prepared = pd.concat([name, punch_in, punch_out], axis=1)

data_prepared.insert(2, '早班打卡要求时间', '6:00')
data_prepared.insert(4, '中班打卡要求时间', '14:00')

data_prepared.insert(3, '早班打卡考核', data_prepared.apply(lambda x: 50 if pd.isna(x['早班打卡']) or parse(x['早班打卡'])  > parse(x['早班打卡要求时间']) else 0, axis=1))

data_prepared.insert(6, '中班打卡考核', data_prepared.apply(lambda x: 50 if pd.isna(x['中班打卡']) or parse(x['中班打卡']) > parse(x['中班打卡要求时间'])else 0, axis=1))

data_prepared.to_excel(result_path, index=False, freeze_panes=[1,0])
