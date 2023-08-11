import os

import win32com.client as win32
import pandas as pd
import numpy as np
from PySide6.QtWidgets import QApplication, QFileDialog

# 需要读取的文件信息，后续可以让用户键盘输入信息

excel_com = 'Excel.Application'



def df_from_pswxlsx(filename, gl_excel_com):
    """读取加密的EXCEL
    参数：
        filename: str -> 文件路径
        gl_excel_com: -> 不同配置的Excel-COM
                      *WPS OFFICE使用参数'KET.APPLICATION'
                      *MS Office Excel使用参数 'EXCEL.APPLICATION'
        sheetname: str -> 打开的工作表名"""
    psw_xlsx = win32.DispatchEx(gl_excel_com)  # 创建对象以打开Excel程序
    # 如果没有密码则直接打开
    # 如果存在密码则捕获异常再次尝试打开
    try:
        psw_xlsx.DisplayAlerts = 0  # 不显示Excel图形界面
        wb = psw_xlsx.Workbooks.Open(filename)
    except:
        password = input('文件密码：')
        wb = psw_xlsx.Workbooks.Open(filename,
                                     UpdateLinks=False,
                                     ReadOnly=False,
                                     Format=None,
                                     Password=password,
                                     WriteResPassword=password)
    psw_xlsx.DisplayAlerts = 0  # 不显示Excel图形界面
    # 获取sheet名
    sheetnames = [sheet.Name for sheet in wb.Sheets]
    if len(sheetnames) > 1:
        print('该文件存在以下Sheet:')
        for i, name in enumerate(sheetnames):
            print(f"第 {i+1} 个表是 {name}")
        sheet_index = int(input("请输入要读取的工作表的序号："))
        sheetname = wb.Sheets(sheet_index)
    else:
        print('默认读取第一个sheet\n')
        sheetname = wb.Sheets(1)
    psw_xlsx.DisplayAlerts = 0
    data_lst = list(sheetname.UsedRange())  # 将读取的sheet表形成一个data列表
    df = pd.DataFrame(data_lst[1:], columns=data_lst[0])  # 将数据转化为Dataframe
    wb.Close()  # 关闭文档

    psw_xlsx.Application.Quit()  # 退出程序
    return df


def check_none(df, col_name):
    """
    检查各个属性是否存在空值
    参数：
        df: ->dataframe  数据
        col_name: ->str 需要进行检查的列名
    """
    row_indexs = df[df[col_name].isnull()].index
    print(f"'{col_name}'为空的所在行的行索引：\n", row_indexs, "\n----------------")
    for row_index in row_indexs:
        print("所在行的值为：\n", df[df.index == row_index][:], "\n----------------")
    condition = True
    while condition:
        # check value 输出这些行的内容检查是否需要删除
        a = input("是否删除以上行的信息(t/f)：")
        if a == 'T' or a == 't':
            df = df.drop(index=row_indexs)
            row_indexs = df[df[col_name].isnull()].index
            print("删除后结果\n", f"'{col_name}'为空的所在行的行索引：", row_indexs)
            break
        elif a == 'F' or a == 'f':
            print("未做删除操作")
            break
        else:
            print("输入错误")
            continue
    # 返回经过删除操作的DataFrame
    return df


# Step1 调用窗口, 获取文件路径
app = QApplication()
file_dialog = QFileDialog
file_name = file_dialog.getOpenFileName()[0]

# 01_读取表1.2
table = df_from_pswxlsx(filename=file_name,
                        gl_excel_com=excel_com)
# 重命名属性名
table.columns = ['借/贷', '客户名称', '存款增减金额', '变动原因']

# Step2 空值检查
# 循环检查各个columns
flag = True
while flag:
    print("\n此文档存在多列: ")
    for col in table.columns:
        print(col)
    col_input = input("-------------------\n选择你需要的列进行空值检查(输入c退出): \n")
    if col_input in table.columns:
        table = check_none(df=table, col_name=col_input)
        continue
    elif col_input == 'c' or col_input == 'C':
        flag = False
    else:
        print("输入错误，重新输入")
        continue

table = table.set_index('客户名称')
# 02_计算所需数据
# 根据【'客户名称'】分组，计算不同组别的【非零的存款增减金额'笔数'】
df_count = pd.DataFrame(table.groupby(by=['客户名称'])['存款增减金额'].agg(np.count_nonzero)).reset_index()
df_count = df_count.rename(columns={'存款增减金额': '笔数'})        # 更改列名为【笔数】

# 根据【'客户名称'】分组，计算不同组别的【存款增减金额】之和
df_sum = pd.DataFrame(table.groupby(by=['客户名称'])['存款增减金额'].sum()).reset_index()

# 根据【客户名称】分组，合并变动原因
table['变动原因'] = table['变动原因'].astype(str)           # 首先将【客户名称转换为str类型】
df_course_of_change = pd.DataFrame(table.groupby(by=['客户名称']).agg({'变动原因': lambda x: '  //  '.join(x)}))

# 将sum和count表merge在一起，作为中间表。
df_merge = pd.merge(df_sum, df_count, on=['客户名称'], how='outer')
# 再 merge 'df_course_of_change' 以加入聚合后的 '变动原因'
df_merge = pd.merge(df_merge, df_course_of_change, on=['客户名称'], how='outer')
# # 保存merge中间表，作为检查的依据
# df_merge.to_excel('merge.xlsx')

# 设置'客户名称'为index
df_merge = df_merge.set_index('客户名称')


final_df = pd.DataFrame(columns=['客户名称', '存款增减金额', '变动原因'])
for index, row in df_merge.iterrows():
    # """
    # 在生成最后的数据时不需要判断是否为【空】
    # 在之前的数据整理阶段我们已经将'客户名称'、'存款增减金额'为空的数据作为无效数据，进行了自主性的删除'
    # """
    new_row = {'客户名称': index,
               '存款增减金额': row['存款增减金额'],
               '变动原因': "（" + str(int(row['笔数'])) + "笔）" + row['变动原因']
               }
    final_df = final_df._append(new_row, ignore_index=True)
# 设置'客户名称'为index
final_df = final_df.set_index('客户名称')
# 存储并保存至新的EXCEL
# 输出
save_name = input("命名保存文件:  ")     # 确定文件名
# 输出一个默认的保存文件的绝对路径
current_dir = os.getcwd()       # 获取当前工作目录的绝对路径
save_path = os.path.join(current_dir, save_name)        # 拼接绝对路径和文件名
final_df.to_excel(save_path+".xlsx")       # 保存

print(f"------------\n文件已保存至  {save_path}\n")
