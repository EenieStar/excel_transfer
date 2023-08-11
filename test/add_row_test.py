# newline test
import pandas as pd

# 假设你的原始dataframe叫做df
df = pd.DataFrame({'客户名称': ['A', 'B', 'C'],
                   '变动原因': ['原因1', '原因2', '原因3'],
                   '回收笔数': [1, 2, 3],
                   '回收': [10, 20, 30],
                   '发放': [100, 200, None]}).set_index('客户名称')
print(df)
# 创建一个空的新dataframe，包含客户名称和变动原因两列
new_df = pd.DataFrame(columns=['客户名称', '变动原因', '贷款金额'])

# 遍历原始dataframe的每一行
for index, row in df.iterrows():
    # 如果发放笔数不为空，则提取其发放值
    if pd.notnull(row['发放']):
        new_row = {'客户名称': index, '变动原因': row['变动原因'], '值': row['发放']}
        new_df = new_df._append(new_row, ignore_index=True)
    # 如果回收笔数不为空，则提取其回收值
    if pd.notnull(row['回收']):
        new_row = {'客户名称': index, '变动原因': row['变动原因'], '值': row['回收']}
        new_df = new_df._append(new_row, ignore_index=True)

# 将新dataframe的index设为客户名称
new_df = new_df.set_index('客户名称')

print(new_df)
