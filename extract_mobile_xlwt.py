# step1 : 遍历指定目录下的excel文件
# step2 : 逐行读取excel文件中的数据
# step3 : 获取第K列的数据，数据格式为 "0715834119 - JOHN MICHUBU SILAS"，通过-分割，获取手机号码，前面的数据为手机号码
# step4 : 提取出手机号码，并将首尾空格去掉，判断手机号是否是0开头且为10位数字，若是，则把0替换为254；若不是，则不做处理
# step5 : 将处理后的手机号码写入result.csv文件中
# step6 : 重复step1~step5，直到遍历完所有的excel文件
# step7 : 将result.csv文件中的数据进行去重处理


from concurrent.futures import ALL_COMPLETED, ThreadPoolExecutor, wait
import os
import re
import random
import threading
import time
from xlrd import open_workbook
import xlwt
import xlutils.copy
from xlutils.filter import process,XLRDReader,XLWTWriter

max_thread = 32

# 读取指定目录下的所有excel文件
def read_excel_file(path, out_path):
    pool= ThreadPoolExecutor(max_workers=max_thread)
    all_task = []
    # 获取指定目录下的所有文件
    file_list = os.listdir(path)
    # 遍历所有文件
    for file in file_list:
        # 判断是否为excel文件
        if file.endswith('.xls') or file.endswith('.xlsx'):
            # 获取文件的绝对路径
            file_path = os.path.join(path, file)
            # 保存过滤后的手机号码
            resust_path = os.path.join(out_path, file).replace('.xls', '_result.csv')
            if os.path.exists(resust_path):
                os.remove(resust_path)
            # 开线程去处理单个excel文件，提高处理效率，避免单个文件处理时间过长
            all_task.append(pool.submit(read_excel, file_path, resust_path))
            # 确保所有线程都处理完成
            # read_excel(file_path, resust_path, out_path)
        else:
            print('不是excel文件')
    # 等待所有线程处理完成
    result=[i.result() for i in all_task]
    print("----all complete-----")
# 读取excel文件
            
def read_excel(file_path, out_path):
    print('开始处理文件：%s' % file_path)
    print('写入文件：%s' % out_path)
    # 打开excel文件
    workbook = open_workbook(file_path, formatting_info=True, on_demand=True)
    # 获取excel文件中所有的sheet
    sheet_names = workbook.sheet_names()
    # 遍历所有的sheet
    mobiles = set()
    for sheet_name in sheet_names:
        # 获取sheet
        sheet = workbook.sheet_by_name(sheet_name)
        # 获取sheet中的行数
        rows = sheet.nrows
        # 获取sheet中的列数
        cols = sheet.ncols
        print('rows: %s, cols: %s' % (rows, cols))
        # 遍历sheet中的每一行
        for row in range(rows):
            # 定义匹配到的手机号码
            phone = ''
            # 获取第K列的数据
            data = sheet.cell(row, 10).value
            # 获取第K列的数据，数据格式为 "0715834119 - JOHN MICHUBU SILAS"，通过-分割，获取手机号码，前面的数据为手机号码
            if data and '-' in data:
                phone = data.split('-')[0].strip()
                # 提取出手机号码，并将首尾空格去掉，判断手机号是否是0开头且为10位数字，若是，则把0替换为254；若不是，则不做处理
                if phone.startswith('0') and len(phone) == 10:
                    phone = '254' + phone[1:]
                mobiles.add(phone)
    # 将处理后的手机号码写入result.csv文件中
    print('写入文件：%s' % out_path)
    with open(out_path, 'a') as f:
        for mobile in mobiles:
            f.write(mobile + '\n')
    # 文件处理完毕
    print('文件处理完毕：%s' % file_path)
    return 1

# step1 : 遍历指定目录下的excel文件
# step2 : 逐行读取excel文件中的数据
# step3 : 获取第K列的数据，数据格式为 "0715834119 - JOHN MICHUBU SILAS"，通过-分割，获取手机号码，前面的数据为手机号码
# step4 : 提取出手机号码，并将首尾空格去掉，判断手机号是否是0开头且为10位数字，若是，则把0替换为254；若不是，则不做处理
# step5 : 将处理后的手机号码写入result.csv文件中
# step6 : 重复step1~step5，直到遍历完所有的excel文件
# step7 : 将result.csv文件中的数据进行去重处理
    
if __name__ == '__main__':
    # 读取extracts_files目录下的所有excel文件
    path =  r'/Users/karl/Work/eagleMobi/python/replaceNewMobile/extracts_files'
    out_path = r'/Users/karl/Work/eagleMobi/python/replaceNewMobile/extracts_files_out'
    # 读取指定目录下的所有excel文件
    read_excel_file(path, out_path)
    # 等待所有线程处理完成
    print('等待所有线程处理完成')
    # 遍历out_path目录下的所有_result.csv文件
    result_csv_files = os.listdir(out_path)
    total_result_set = set()
    for result_csv in result_csv_files:
        # 将内容写入total_result.csv文件中
        with open(os.path.join(out_path, result_csv), 'r') as f:
            data = f.readlines()
        for line in data:
            total_result_set.add(line)
    # total_result_set 写入result.csv文件中
    with open('total_result.csv', 'w') as f:
        for line in total_result_set:
            f.write(line)
    print('处理完成')