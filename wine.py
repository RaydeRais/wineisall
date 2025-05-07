import pandas as pd
from fuzzywuzzy import fuzz
import os

# 文件路径
price_table_path = r'D:\winee\价格表.xlsx'  # 价格表文件路径
training_data_path = r'D:\winee\训练集.xlsx'  # 训练集文件路径
output_path = r'D:\winee\编号.xlsx'          # 输出文件路径
extracted_fields_path = r'D:\winee\提取字段.xlsx'  # 提取字段的输出路径

# 读取 Excel 文件
price_table = pd.read_excel(price_table_path)
training_df = pd.read_excel(training_data_path)

# 定义查找最近匹配的函数
def find_closest_match(brand, price_table, threshold=80):
    best_match = None
    best_ratio = 0
    for _, row in price_table.iterrows():
        ratio = fuzz.token_sort_ratio(brand, row['品牌'])
        if ratio > best_ratio:
            best_ratio = ratio
            best_match = row
    if best_ratio >= threshold:
        return best_match, best_ratio
    return None, best_ratio

# 匹配并比较价格，收集符合条件的编号
results = []
matches = []

for _, row in training_df.iterrows():
    match, ratio = find_closest_match(row['品牌'], price_table)
    if match is not None:
        comparison_price = match['对比价格(欧元)']
        threshold = comparison_price * 0.95  # 计算 95% 阈值
        if row['单价（欧元）'] < threshold:
            results.append(row['编号'])
            matches.append({
                '编号': row['编号'],
                '经营单位名称': row['经营单位名称'],
                '商品规格、型号': row['商品规格、型号'],
                '单价（欧元）': row['单价（欧元）'],
                '酒名称': row['酒名称'],
                '品牌': row['品牌'],
                '年份': row['年份'],
                '匹配品牌': match['品牌'],
                '相似度': ratio,
                '对比价格(欧元)': comparison_price
            })

# 将提取的字段保存到新的 Excel 文件
if matches:
    df_matches = pd.DataFrame(matches)
    df_matches.to_excel(extracted_fields_path, index=False)
    print(f"提取的字段已保存到 {extracted_fields_path}")

# 将结果保存到新的 Excel 文件
if results:
    df_results = pd.DataFrame(results, columns=['编号'])
    df_results.to_excel(output_path, index=False)
    print(f"结果已保存到 {output_path}")
else:
    print("没有符合条件的编号。")