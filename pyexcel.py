# pip install openpyxl
import time
from openpyxl import Workbook
from openpyxl import load_workbook
import logging
import time
import traceback
import sys
import os.path
import os

start = time.perf_counter()
# 实例化
# wb = Workbook()
# # 激活 worksheet
# # ws = wb.active

# wb2 = load_workbook('分平台营业数据 (34).xlsx')

# ws2=wb2.worksheets[0]
# print(ws2["C2"].value)
# wb3 = load_workbook('评论率 (33).xlsx')
# ws3=wb3.worksheets[0]
# ws3["G2"].value=ws2["C2"].value
# wb3.save('评论率 (33).xlsx')


a = []
for line in open("name.txt"):
    a.append(line.strip())

def update(a):
    try:
        wb1 = load_workbook('分平台营业数据 (34).xlsx')
        wb2 = load_workbook('分平台营业数据 (34).xlsx')
        wb3 = load_workbook('分平台营业数据 (34).xlsx')
        wb4 = load_workbook('分平台营业数据 (34).xlsx')
    except Exception:
        
        # 第一步，创建一个logger
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)  # Log等级总开关
        # 第二步，创建一个handler，用于写入日志文件
        rq = time.strftime('%Y%m%d%H%M', time.localtime(time.time()))
        log_path = './Logs/'
        log_name = log_path + rq + '.log'
        logfile = log_name
        fh = logging.FileHandler(logfile, mode='a')
        fh.setLevel(logging.DEBUG)  # 输出到file的log等级的开关
        # 第三步，定义handler的输出格式
        formatter = logging.Formatter("%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s")
        fh.setFormatter(formatter)
        # 第四步，将logger添加到handler里面
        logger.addHandler(fh)
        logger.error("Faild to open sklearn.txt from logger.error",exc_info = True)

    

if(os.path.exists('./Logs')==False):
    os.mkdir("Logs")


update(a)
end = time.perf_counter()
print(end - start)