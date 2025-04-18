#!/usr/bin/env python
# coding: utf-8

# In[32]:


import pandas as pd
import numpy as np
from datetime import datetime

# 读取 employment 数据
df = pd.read_excel('employment.xlsx')

def calc_work_length(row):
    if pd.isnull(row['Start Date']) or pd.isnull(row['End Date']):
        return ''
    start = pd.to_datetime(row['Start Date'], dayfirst=True, errors='coerce')
    end = pd.to_datetime(row['End Date'], dayfirst=True, errors='coerce')
    if pd.isnull(start) or pd.isnull(end):
        return ''
    years = end.year - start.year - ((end.month, end.day) < (start.month, start.day))
    months = (end.year - start.year) * 12 + end.month - start.month
    months = months % 12
    result = ''
    if years > 0:
        result += f'{years} year(s) '
    if months > 0:
        result += f'{months} month(s)'
    return result.strip()

df['Work Length'] = df.apply(calc_work_length, axis=1)
df.dropna(subset=['Start Date', 'End Date'], how='any', inplace=True)

# 拼接工作信息
df['Previous working experience'] = (
    df['Employer'].fillna('') + ', ' +
    df['Country'].fillna('') + ', ' +
    df['Job Title'].fillna('') + ', ' +
    df['Work Length'].replace('', 'N/A')
).str.replace(r'(, ){2,}', ', ', regex=True).str.strip(', ')

# 汇总工作信息
summary_df = df.groupby(['First Name', 'Last Name']).agg({
    'Previous working experience': lambda x: '\n'.join(x),
    'Grade (for UN staff)': 'first'  # 保留 grade，等下改 current grade
}).reset_index()

summary_df.rename(columns={'Grade (for UN staff)': 'Current Grade'}, inplace=True)
summary_df['Previous working experience'] = summary_df['Previous working experience'].str.replace(r',\s*,', ',', regex=True)
summary_df['Previous working experience'] = summary_df['Previous working experience'].str.replace(r'\s+,', ',', regex=True)

# 读取 education 数据
edu_df = pd.read_excel('education.xlsx')

# 计算年龄
edu_df['Age'] = edu_df['Date of Birth'].apply(
    lambda x: (datetime.now() - pd.to_datetime(x, dayfirst=True)).days // 365 if not pd.isnull(x) else np.nan
)

# 处理 Institution 和 Internal/External
edu_df['Institution'] = edu_df.apply(
    lambda row: row['Other Institution'] if row['Institution'] == '* Other – Cannot find my school in the list' else row['Institution'],
    axis=1
)

edu_df['Internal/External'] = edu_df['Is Internal'].map({'Yes': 'Internal', 'No': 'External'})

# 映射教育层级
edu_level_map = {
    "5 Master Degree": "MA",
    "6 PhD Doctorate Degree": "PHD",
    "4 Bachelor's Degree": "BA",
    "3 Technical Diploma": "TD",
    "1 Non-Degree Programme": "Non-Degree",
    "2 High School diploma": "High School"
}
edu_df['Education Level'] = edu_df['Education Level'].map(edu_level_map).fillna(edu_df['Education Level'])

# 计算最高学历
level_score = {
    "PHD": 5,
    "MA": 4,
    "BA": 3,
    "TD": 2,
    "Non-Degree": 0,
    "High School": 1
}
edu_df['Level Score'] = edu_df['Education Level'].map(level_score)
edu_df['Max Level Score'] = edu_df.groupby(['First Name', 'Last Name'])['Level Score'].transform('max')
edu_df['Keep'] = np.where(edu_df['Level Score'] == edu_df['Max Level Score'], 'Keep', '')

# 处理主修科目
edu_df['Main subject'] = edu_df['Main subject'].astype(str).str.split(',').str[0]

# 筛选出最高学历
edu_keep_df = edu_df[edu_df['Keep'] == 'Keep'].copy()

# 拼接教育信息
edu_keep_df['Highest Education'] = (
    edu_keep_df['Education Level'].fillna('') + ' - ' +
    edu_keep_df['Institution'].fillna('') + ' (' +
    edu_keep_df['Country'].fillna('') + ') - ' +
    edu_keep_df['Main subject'].fillna('')
)

# 提取相关信息
edu_summary_df = edu_keep_df[[
    'First Name', 'Last Name', 'Age', 'Internal/External',
    'Highest Education', 'Geo Dist. Representation', 'Primary Nationality'
]]

# 合并到最终表格
final_df = summary_df.merge(edu_summary_df, on=['First Name', 'Last Name'], how='left')

# 如果 Internal，则将 'Geo Dist. Representation' 改为 'Non impact'
final_df['Geo Dist. Representation'] = final_df.apply(
    lambda row: 'Non impact' if row['Internal/External'] == 'Internal' else row['Geo Dist. Representation'], axis=1
)

# 填充 'Current Grade' 的空值为 'N/A'
final_df['Current Grade'] = final_df['Current Grade'].fillna('N/A')

# 按照指定的列顺序进行排序
final_df = final_df[[
    'First Name', 'Last Name', 'Age', 'Current Grade', 'Internal/External', 'Primary Nationality', 
    'Geo Dist. Representation', 'Previous working experience', 'Highest Education'
]]

# 导出最终数据
final_df.to_excel('final_cleaned_data.xlsx', index=False)


# In[ ]:




