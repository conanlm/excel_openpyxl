# pip install openpyxl

from openpyxl import Workbook
from openpyxl import load_workbook
# 实例化
wb = Workbook()
# 激活 worksheet
# ws = wb.active

wb2 = load_workbook('分平台营业数据 (34).xlsx')

ws2=wb2.worksheets[0]
print(ws2["C2"].value)
wb3 = load_workbook('评论率 (33).xlsx')
ws3=wb3.worksheets[0]
ws3["G2"].value=ws2["C2"].value
wb3.save('评论率 (33).xlsx')