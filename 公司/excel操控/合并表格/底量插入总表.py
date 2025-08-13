import pandas as pd

# 定义总表读取和写入函数
def merge_total_table(total, old):
    total_df = pd.read_excel(total)
    old_df = pd.read_excel(old)

    # 处理总表数据的合并
    old_df.iloc[:, 0:27] = total_df.iloc[:, 0:27]
    old_df.iloc[:, 28:29] = pd.NA
    old_df.iloc[:, 30:44] = total_df.iloc[:, 28:42]
    old_df.iloc[:, 45] = total_df.iloc[:, 44]
    old_df.iloc[:, 46:50] = total_df.iloc[:, 46:50]
    old_df.to_excel(total, index=False)

total = "C:\\Users\\fucha\\Desktop\\公司\\4.代码表格\\合表表格\\总表.xlsx"
old = "C:\\Users\\fucha\\Desktop\\公司\\4.代码表格\\合表表格\\底量.xlsx"
# 调用函数执行合并操作，并传递初始行数
merge_total_table(total, old)