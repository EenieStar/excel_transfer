# 测试df_iterrows()
import pandas as pd
df = pd.DataFrame([['liver', 'E', 89, 21, 24, 64],
                   ['Arry', 'C', 36, 37, 37, 57],
                   ['Ack', 'A', 57, 60, 18, 84],
                   ['Eorge', 'C', 93, 96, 71, 78],
                   ['Oah', 'D', 65, 49, 61, 86]
                  ],
                  columns=['name', 'team', 'Q1', 'Q2', 'Q3', 'Q4'])
# 使用df.iterrows()进行迭代操作
for index, row in df.iterrows():
    # print(index)
    # print(index, row['name'])
    a = df[index, row['变动原因']]
    print(a)


