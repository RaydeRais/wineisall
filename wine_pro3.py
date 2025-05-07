import pandas as pd
from fuzzywuzzy import fuzz
import os

# 文件路径
price_table_path = r'D:\winee\价格表.xlsx'  # 价格表文件路径
training_data_path = r'D:\winee\训练集.xlsx'  # 训练集文件路径
output_path = r'D:\winee\提取字段.xlsx'  # 提取字段的输出路径

# 读取 Excel 文件
price_table = pd.read_excel(price_table_path)
training_df = pd.read_excel(training_data_path)

# 定义查找最近匹配的函数
def find_closest_match(brand, year, price_table, threshold=80):
    best_matches = []
    best_ratio = 0
    for _, row in price_table.iterrows():
        ratio = fuzz.token_sort_ratio(brand, row['品牌'])
        if ratio >= threshold:
            if ratio > best_ratio:
                best_ratio = ratio
            best_matches.append((row, ratio))
    
    if not best_matches:
        return None, None, 0
    
    # 如果有多个匹配项，进一步匹配年份
    if len(best_matches) > 1:
        best_match = None
        min_year_diff = float('inf')
        
        for match, _ in best_matches:
            if pd.notna(match['年份']):
                try:
                    if year == "未标注":
                        # 如果年份是“未标注”，选择年份最小的匹配项
                        if best_match is None or int(match['年份']) < int(best_match['年份']):
                            best_match = match
                    else:
                        year_diff = abs(int(year) - int(match['年份']))
                        if year_diff == 0:
                            best_match = match
                            break
                        elif year_diff < min_year_diff and year_diff < 6:
                            min_year_diff = year_diff
                            best_match = match
                except ValueError:
                    # 如果年份转换失败（例如包含非数字字符），则忽略年份匹配
                    pass
        
        if best_match is not None:
            return best_match, best_ratio, 1  # 年份匹配成功或年份差值小于6
        else:
            return None, None, 0  # 没有合适的年份匹配
    else:
        # 只有一个匹配项或没有年份信息
        match, _ = best_matches[0]
        if pd.notna(year) and pd.notna(match['年份']):
            try:
                year_diff = abs(int(year) - int(match['年份']))
                if year_diff > 5:
                    return None, None, 0  # 年份差值大于5，不作为匹配成功
            except ValueError:
                # 如果年份转换失败（例如包含非数字字符），则忽略年份匹配
                pass
        return match, best_ratio, 1  # 匹配成功

# 匹配并比较价格，收集符合条件的记录
matches = []

for _, row in training_df.iterrows():
    match, ratio, year_matched = find_closest_match(row['品牌'], row['年份'], price_table)
    if match is not None:
        comparison_price = match['对比价格(欧元)']
        threshold = comparison_price * 0.7  # 计算 60% 阈值
        if row['单价（欧元）'] < threshold:
            matches.append({
                '编号': row['编号'],
                '经营单位名称': row['经营单位名称'],
                '商品规格、型号': row['商品规格、型号'],
                '单价（欧元）': row['单价（欧元）'],
                '酒名称': row['酒名称'],
                '品牌': row['品牌'],
                '年份': row['年份'],
                '匹配品牌': match['品牌'],
                '匹配年份': match['年份'],
                '相似度': ratio,
                '年份匹配': '是' if year_matched else '否',
                '对比价格(欧元)': comparison_price
            })

# 将提取的字段保存到新的 Excel 文件
if matches:
    df_matches = pd.DataFrame(matches)
    df_matches.to_excel(output_path, index=False)
    print(f"提取的字段已保存到 {output_path}")
else:
    print("没有符合条件的记录。")