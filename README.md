# StudentNameSort
It is not only for matching or sorting student names. I did this script to help my dad.

## Objective
Using Python to match names and others from two different Excel files

## Key Features
1. Path Safety Handling: Use the pathlib module to construct file paths, avoiding errors from manually written string paths.
2. Save Results to a New File: Save the output to a new file "Results.xlsx" to prevent overwriting the original data.
3. Clean data: remove leading and trailing spaces. Ensure names are stripped of any leading or trailing spaces. Standardize names to lowercase to avoid mismatches due to case differences. Auto-Filter Empty Values (dropna).
4. Column Name Correction: Set the names parameter uniformly to a single column name ['Name']. Remove redundant column name definitions from the original code.
5. Error Handling: Catch file reading exceptions. Provide a user-friendly message for empty results.
6. Result Optimization: Display names with the first letter capitalized (.title()). Add match count statistics. Output key statistical information to the console.
7. Add the header=None parameter: Ensure that the first row is not mistakenly interpreted as a header row.
8. Support Multiple Sheets: Use pd.ExcelFile to read the file and obtain all sheet names through sheet_names. Iterate through each sheet and perform matching separately.
9. Record Sheet Name: Add a "Sheet Name" field in the results to record the sheet to which each match belongs.
10. Dynamic Matching: Process the data for each sheet independently to ensure the matching results are accurate.



## Future Progress
1. Support for Multiple Files and Multiple Sheets: please refer to MultipleFiles_MultipleSheets.py
2. Generate a visual report: please refer to Matplotlib.py
3. Support fuzzy matching: please refer to FuzzyMatching.py
4. Return Class Information: Add a "Class" column in this year's file and include the class field in the Python results. For example, df_current = pd.read_excel(..., usecols=[0, 1], names=['Name', 'Class'])
5. Highlighting: Use conditional formatting in Excel → Highlight Cell Rules.



## Note
1. To run this script or program, you must install pandas, openpyxl and xlrd libraries to your python lib.
pip install pandas openpyxl xlrd
2. In the main code, I did force type conversion: Use dtype={'Name': str} to ensure the Name column is read as a string. Also, use str(x) in converters to ensure all values are treated as strings.

Thank you for reading here because, without this script, you can also realize such action just by using commands in Excel. 
So, it may seem not useful yet useless at all. 
Here are the ways to realize this without using Python script:
1. =MATCH(A1, [Filename.xlsx]Sheet1!$B$1:$B$100, 0) (it will show the column number if matching successfully, else #N/A which means unmatch.)
2. =IFERROR(ADDRESS(MATCH(B1,[Filename.xlsx]Sheet1!$B$1:$B$100,0)+1,1), "NoFound") (this will show the cell address if matching successfully.)
