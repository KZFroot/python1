import pandas as pd
import openpyxl
'''
distance_list是一个列表，我们的目标是将该列表作为一列插入表格
'''
# 先打开我们的目标表格，再打开我们的目标表单
wb=openpyxl.load_workbook('result/result_0.xlsx')
ws = wb['Sheet1']
# 取出distance_list列表中的每一个元素，openpyxl的行列号是从1开始取得，所以我这里i从1开始取
for i in range(1,len(distance_list)+1):
    distance=distance_list[i-1]
    # 写入位置的行列号可以任意改变，这里我是从第2行开始按行依次插入第11列
    ws.cell(row = i+1, column = 11).value =distance
# 保存操作
wb.save(r'D:\working\FirstPaper\user_info_distance.xlsx')