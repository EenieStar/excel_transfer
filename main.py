import os
# 选择执行哪一个功能
# 1.转换表1.1格式的文件
# 2.转换表1.2格式的文件

choice = input("选择功能\n1.表1.1\n2.表1.2")
f = True
while f:
    if choice == '1':
        os.system('D:/re/data1 python tablea.py')
        f = False
    elif choice == '2':
        os.system('python/D:/re/data2/tableb.py')
        f = False
    else:
        print('无效选择')
        continue

