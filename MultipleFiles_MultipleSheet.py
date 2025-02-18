# To process multiple files simultaneously (such as Filename2.xlsx and Filename3.xlsx), you can extend the code as follows:

files_current = [base_dir / "filename2.xlsx", base_dir / "filename3.xlsx"]
for file in files_current:
    excel_file = pd.ExcelFile(file)
    for sheet_name in excel_file.sheet_names:
        # 处理每个文件的每个Sheet Process each sheet of every file
