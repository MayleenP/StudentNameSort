import pandas as pd
from pathlib import Path

# ================= 配置区域 =================
base_dir = Path(r"C:\Users\yongz\Documents\ClassPy")
file_last = base_dir / "ClassA.xlsx"  # 去年名单
file_current = base_dir / "ClassB.xlsx"  # 今年名单
output_file = base_dir / "查找结果.xlsx"  # 结果文件
'''search_count = 5  # 查找前N名学生'''


# ===========================================

def clean_data(df):
    """数据清洗函数"""
    df['Name'] = df['Name'].astype(str).str.strip().str.lower()
    return df.dropna(subset=['Name'])


# 读取文件（注意：姓名在B列，使用usecols=[1]）
try:
    df_last = pd.read_excel(
        file_last,
        usecols=[1],  # 读取B列（索引1）
        names=['Name'],
        header=None,  # 无标题行
        converters={'Name': lambda x: str(x).strip().lower()}
    )

    df_current = pd.read_excel(
        file_current,
        usecols=[1],  # 读取B列（索引1）
        names=['Name'],
        header=None,  # 无标题行
        converters={'Name': lambda x: str(x).strip().lower()}
    )
except Exception as e:
    print(f"文件读取失败：{str(e)}")
    exit()

# 数据清洗
df_last = clean_data(df_last)
df_current = clean_data(df_current)

# 提取目标学生（前5个）
'''targets = df_last['Name'].head(search_count).tolist()'''
targets = ["james", "michael", "robert", "john", "david"]   #指定名字目标

# 执行查找
results = []
for idx, row in df_current.iterrows():
    if row['Name'] in targets:
        # 生成正确的单元格地址（B列，从B1开始）
        cell_address = f"B{idx + 1}"  # 从1开始，无标题行
        results.append({
            '目标姓名': row['Name'].title(),  # 首字母大写
            '所在文件': 'ClassB.xlsx',
            '单元格地址': cell_address,
            '行号': idx + 1
        })

# 保存结果
if results:
    result_df = pd.DataFrame(results)

    # 添加匹配统计
    match_count = result_df.groupby('目标姓名').size().reset_index(name='出现次数')
    final_df = pd.merge(result_df, match_count, on='目标姓名')

    final_df.to_excel(output_file, index=False)
    print(f"成功找到 {len(results)} 处匹配，结果已保存至：{output_file}")

    # 控制台输出统计信息
    print("\n匹配统计：")
    print(final_df[['目标姓名', '出现次数']].drop_duplicates())
else:
    print("未找到任何匹配项，请检查："
          "\n1. 两文件姓名是否一致"
          "\n2. 清洗后数据是否有效")

# 验证输出样例（前3条）
if len(results) >= 3:
    print("\n示例结果：")
    print(final_df[['目标姓名', '单元格地址']].head(3))