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
    print(f"FAILED TO READ FILES：{str(e)}")
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
            'NAME': row['Name'].title(),  # 首字母大写 First letter capital letter  
            'FILE': 'ClassB.xlsx',
            'CELL ADDRESS': cell_address,
            'COLUMN NUM': idx + 1
        })

# 保存结果 Save results
if results:
    result_df = pd.DataFrame(results)

    # 添加匹配统计 Add count
    match_count = result_df.groupby('NAME').size().reset_index(name='APPEARANCES')
    final_df = pd.merge(result_df, match_count, on='NAME')

    final_df.to_excel(output_file, index=False)
    print(f"{len(results)} SUBJECTS SUCCESSFULLY MATCH，THE RESULTS SAVED TO：{output_file}")

    # 控制台输出统计信息 Output details at the terminal
    print("\nMATCH STATISTICS：")
    print(final_df[['NAME', 'APPEARANCES']].drop_duplicates())
else:
    print("NO MATCH FOUND，PLEASE CHECK："
          "\n1. TWO FILENAMES ARE CORRECT"
          "\n2. IS THE CLEAN DATA WORKING?")

# 验证输出样例（前3条）For testing output (can be ignored)
if len(results) >= 3:
    print("\nEXAMPLE RESULTS：")
    print(final_df[['NAME', 'CELL ADDRESS']].head(3))
