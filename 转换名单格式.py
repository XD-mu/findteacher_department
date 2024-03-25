
import pandas as pd
import re

# 确保在读取Excel文件时正确设置了header参数
pre_conversion_df = pd.read_excel("./转换前名单表.xlsx", header=None)

# 移除空行
pre_conversion_df.dropna(inplace=True)

# 定义提取姓名的函数
def extract_name(text):
    match = re.search(r'\[(.*?)\]', text)
    if match:
        return match.group(1)
    else:
        return None

# 应用函数提取姓名
pre_conversion_df['姓名'] = pre_conversion_df.iloc[:, 0].apply(extract_name)  # 使用iloc[:, 0]来通用地引用第一列

# 移除任何可能由于匹配失败而产生的NaN值
converted_df = pre_conversion_df[['姓名']].dropna()

# 将结果保存到新的Excel文件中
converted_df.to_excel("./转换后名单表.xlsx", index=False)
