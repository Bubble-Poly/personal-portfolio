import os
import pandas as pd

# 获取工作目录
current_dir = os.getcwd()

# 路径
file_path = os.path.join(current_dir, '腾飞营雅思阅读第一课-单词复习计划表.xlsx')

# 读取Excel
df = pd.read_excel('d:\00 桌面\雅思\超级加倍.xlsx', sheet_name='Sheet1')

# 修改背诵内容
def double_pages(page_range, step=18):
    if '-' not in page_range:
        return page_range
    start, end = map(int, page_range[1:].split('-'))
    new_start = start
    new_end = start + step - 1
    return f'P{new_start}-{new_end}'

# 调整天数
def adjust_days(df, base_period_days, advanced_period_days, sprint_period_days):
    total_days = len(df)    # 总天数

    # 各阶段的天数
    base_days = int((base_period_days / (base_period_days + advanced_period_days + sprint_period_days)) * total_days)
    advanced_days = int((advanced_period_days / (base_period_days + advanced_period_days + sprint_period_days)) * total_days)
    sprint_days = total_days - base_days - advanced_days
    
    new_df = df.iloc[:base_days + advanced_days + sprint_days]    # 创建新的DataFrame
    new_df['背诵内容'] = new_df['背诵内容'].apply(double_pages)    # 修改 
    return new_df

# 设定各个阶段的原始天数比例
base_period_days = 10    # 原始的基础期为10天
advanced_period_days = 10    # 原始的进阶期为10天
sprint_period_days = 8    # 原始的冲刺期为8天

# 调整天数并加倍背诵内容
adjusted_df = adjust_days(df, base_period_days, advanced_period_days, sprint_period_days)

# 输出结果
adjusted_df.to_excel('output.xlsx', index=False)