import pandas as pd # 引入pandas库 Import pandas library 
from pathlib import Path

# ================= 配置区域 =================
base_dir = Path(r"C:\Users\yongz\Documents\ClassPy")
file_last = base_dir / "Filename1.xlsx"  # 更改自己Excel文件1 Input Excel file 1
file_current = base_dir / "Filename2.xlsx"  # 更改自己Excel文件2 Input Excel file 2
output_file = base_dir / "Results.xlsx"  # 结果文件 Build file for displaying results
'''search_count = 5'''  # 查找前N名学生 Search for first N names


# ===========================================

def clean_data(df):
    """数据清洗函数 clean data function"""
    df['Name'] = df['Name'].astype(str).str.strip().str.lower()
    return df.dropna(subset=['Name'])

# Read files (Note: if names are located at B row, then usecols=[1]; else if names are located at A row, then usecols=[0]
# 读取Filename1.xlsx（Excel文件1）Read Excel file 1
try:
    df_last = pd.read_excel(
        file_last,
        usecols=[1],  # 读取B列（索引1）Read B row
        names=['Name'],
        header=None,  # 无标题行 None header to read column 1 
        converters={'Name': lambda x: str(x).strip().lower()}
    )
except Exception as e:
    print(f"Filename1.xlsx FILE READ FAILED：{str(e)}")
    exit()

# 数据清洗 Clean data
df_last = clean_data(df_last)

'''targets = df_last['Name'].head(search_count).tolist()'''  # 提取目标学生（前5个）Choose targeted names (first 5)
targets = ["james", "michael", "robert", "john", "david"]   # 指定名字目标 Choose targeted names

# 读取Filename2.xlsx（Excel文件2，支持多Sheet）Read Excel file 2 (Multiple Sheets)
try:
    # 使用ExcelFile读取所有Sheet Use ExcelFile to read all Sheets
    excel_file = pd.ExcelFile(file_current)
    sheet_names = excel_file.sheet_names  # 获取所有Sheet名称 Get all Sheet names
except Exception as e:
    print(f"Filename2.xlsx FILE READ FAILED：{str(e)}")
    exit()

# 执行查找 Run search
results = []
for sheet_name in sheet_names:
    # 读取当前Sheet Read the current sheet
    df_current = pd.read_excel(
        excel_file,
        sheet_name=sheet_name,
        usecols=[1],  # 读取B列（索引1）
        names=['Name'],
        header=None,  # 无标题行
        converters={'Name': lambda x: str(x).strip().lower()}
    )

    # 数据清洗
    df_current = clean_data(df_current)

    # 在当前Sheet中查找匹配项 Match at the current Sheet
    for idx, row in df_current.iterrows():
        if row['Name'] in targets:
            # 生成正确的单元格地址（B列，从B1开始）Create cell address (If B rows, then start from B1)
            cell_address = f"B{idx + 1}"  # 从1开始，无标题行 None header, start from column 1 
            results.append({
                'NAME': row['Name'].title(),  # 首字母大写 First letter capital letter  
                'FILE': 'ClassB.xlsx',
                'SHEET NAME': sheet_name,  # 记录Sheet名称 Input Sheet name
                'CELL ADDRESS'': cell_address,
                'COLUMN NUM': idx + 1
            })

# 保存结果 Save results
if results:
    result_df = pd.DataFrame(results)

    # 添加匹配统计 Add count
    match_count = result_df.groupby(['NAME', 'SHEET NAME']).size().reset_index(name='APPEARANCES')
    final_df = pd.merge(result_df, match_count, on=['NAME', 'SHEET NAME'])

    final_df.to_excel(output_file, index=False)
    print(f"{len(results)} SUBJECTS SUCCESSFULLY MATCH，THE RESULTS SAVED TO：{output_file}")

    # 控制台输出统计信息 Output details at the terminal
    print("\nMATCH STATISTICS：")
    print(final_df[['NAME', 'SHEET NAME', 'APPEARANCES']].drop_duplicates())
else:
    print("NO MATCH FOUND，PLEASE CHECK："
          "\n1. ARE THE TWO FILENAMES CORRECT?"
          "\n2. IS THE CLEANED DATA WORKING?")

# 验证输出样例（前3条）For testing output (can be ignored)
if len(results) >= 3:
    print("\nEXAMPLE RESULTS：")
    print(final_df[['NAME', 'SHEET NAME', 'CELL ADDRESS']].head(3))
