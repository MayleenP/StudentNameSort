# 使用 thefuzz 库实现模糊匹配 Use the thefuzz library to implement fuzzy matching. You can add the below code to the main code.

from thefuzz import process
matches = process.extractBests(target, df_current['Name'], limit=3)
