from cx_Freeze import setup, Executable
import sys

sys.setrecursionlimit(5000)
setup(
    name='tablea转换',
    version='1.0',
    description='汇总转换原始表数据，形成新的目标表格',
    executables=[Executable("tableb.py")]
)