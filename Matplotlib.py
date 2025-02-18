# 使用 matplotlib 生成柱状图，按 Sheet 名称分组显示匹配次数 Use matplotlib to create a bar chart, grouping the match counts by sheet name.

import matplotlib.pyplot as plt
final_df.groupby('SHEET NAME')['NAME'].count().plot(kind='bar')
plt.savefig('MATCH STATISTICS.png')
