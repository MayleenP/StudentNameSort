import pandas as pd    # 引入pandas库 Import pandas library 
from pathlib import Path

# ================= 配置区域 =================
base_dir = Path(r"C:\Users\yongz\Documents\ClassPy")
file_last = base_dir / "Filename1.xlsx"  # 更改自己Excel文件1 Input Excel file 1
file_current = base_dir / "Filename2.xlsx"  # 更改自己Excel文件2 Input Excel file 2
output_file = base_dir / "Results.xlsx"  # 结果文件 Build file for displaying results
'''search_count = 5'''  # 查找前N名姓名 Search for first N names


# ===========================================

def clean_data(df):
    """数据清洗函数 clean data function"""
    df['Name'] = df['Name'].astype(str).str.strip().str.lower()
    return df.dropna(subset=['Name'])


# 读取文件（注意：如果姓名在B列，使用usecols=[1]；反之，如果姓名在A列，使用usecols=[0]）Read files (Note: if names are located at B row, then usecols=[1]; else if names are located at A row, then usecols=[0]
try:
    df_last = pd.read_excel(
        file_last,
        usecols=[1],  # 读取B列（索引1）Read B row
        names=['Name'],
        header=None,  # 无标题行 None header to read column 1 
        converters={'Name': lambda x: str(x).strip().lower()}
    )

    df_current = pd.read_excel(
        file_current,
        usecols=[1],  # 读取B列（索引1
        names=['Name'],
        header=None,  # 无标题行 
        converters={'Name': lambda x: str(x).strip().lower()}
    )
except Exception as e:
    print(f"文件读取失败：{str(e)}")
    exit()

# 数据清洗 Clean data
df_last = clean_data(df_last)
df_current = clean_data(df_current)


'''targets = df_last['Name'].head(search_count).tolist()'''    # 提取目标学生（前5个）Choose targeted names (first 5)
targets = ["james", "michael", "robert", "john", "david"]   # 指定名字目标 Choose targeted names

# 执行查找 Run search
results = []
for idx, row in df_current.iterrows():
    if row['Name'] in targets:
        # 生成正确的单元格地址（B列，从B1开始）Create cell address (If B rows, then start from B1)
        cell_address = f"B{idx + 1}"  # 从1开始，无标题行 None header, start from column 1 
        results.append({
            '目标姓名': row['Name'].title(),  # 首字母大写 First letter capital letter  
            '所在文件': 'ClassB.xlsx',
            '单元格地址': cell_address,
            '行号': idx + 1
        })

# 保存结果 Save results
if results:
    result_df = pd.DataFrame(results)

    # 添加匹配统计 Add count
    match_count = result_df.groupby('目标姓名').size().reset_index(name='出现次数')
    final_df = pd.merge(result_df, match_count, on='目标姓名')

    final_df.to_excel(output_file, index=False)
    print(f"成功找到 {len(results)} 处匹配，结果已保存至：{output_file}")

    # 控制台输出统计信息 Output details at the terminal
    print("\n匹配统计：")
    print(final_df[['目标姓名', '出现次数']].drop_duplicates())
else:
    print("未找到任何匹配项，请检查："
          "\n1. 两文件姓名是否一致"
          "\n2. 清洗后数据是否有效")

# 验证输出样例（前3条）For testing output (can be ignored)
if len(results) >= 3:
    print("\n示例结果：")
    print(final_df[['目标姓名', '单元格地址']].head(3))
