# step1 : 遍历指定目录下的excel文件
# step2 : 逐行读取excel文件中的数据
# step3 : 匹配每一行中是否出现格式为: Business Payment to 254712366882的数据，其中254712366882为手机号码
# step4 : 提取出手机号码，并判断是否在指定的手机号码列表中，若在列表中，则从另一个匹配列表中随机获取一个手机号码进行替换
# step5 : 将替换后的数据写入新的excel文件中
# step6 : 重复step1~step5，直到遍历完所有的excel文件

import os
import re
import random
import xlwings as xw

# 读取指定目录下的所有excel文件
def read_excel_file(path, out_path):
    # 获取指定目录下的所有文件
    file_list = os.listdir(path)
    # 遍历所有文件
    for file in file_list:
        # 判断是否为excel文件
        if file.endswith('.xls') or file.endswith('.xlsx'):
            # 获取文件的绝对路径
            file_path = os.path.join(path, file)
            out_path = os.path.join(out_path, file)
            # 读取excel文件
            read_excel(file_path, out_path)
        else:
            print('不是excel文件')
    
# 读取excel文件
def read_excel(file_path, out_path):
    # 打开excel文件
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    workbook = app.books.open(file_path)
    # 新的excel文件名,末尾加_new
    new_file_name = out_path.replace('.xls', '_new.xls')
    # 若已经存在_new的文件，则删除
    if os.path.exists(new_file_name):
        os.remove(new_file_name)

    # 获取excel文件中所有的sheet
    sheet_names = workbook.sheet_names
    # 遍历所有的sheet
    for sheet_name in sheet_names:
        # 获取sheet
        sheet = workbook.sheets[sheet_name]
        # 获取sheet中的行数
        rows = sheet.used_range.rows.count
        # 获取sheet中的列数
        cols = sheet.used_range.columns.count
        print('rows: %s, cols: %s' % (rows, cols))
        # 遍历sheet中的每一行
        for row in range(rows):
            # 打印处理到第几行
            print('处理到第%s行' % row)
            if row == 20:
                break
            # 定义匹配到的手机号码
            phone = ''
            # 定义替换手机号码
            replace_phone = ''
            # 遍历sheet中的每一列
            for col in range(cols):
                # 获取单元格的值
                cell_value = sheet.cells(row+1, col+1).value
                # 判断单元格的值是否为字符串类型
                if isinstance(cell_value, str):
                    if phone == '':
                        # 匹配单元格中是否出现格式为: Business Payment to 任意数字 - 的数据
                        match = re.search(r'Business Payment to \d+', cell_value)
                        # 判断是否匹配成功
                        if match:
                            # 获取手机号码
                            tempPhone = match.group(0).split(' ')[-1]
                            # 判断手机号码是否在指定的手机号码列表中
                            if tempPhone in phone_list:
                                phone = tempPhone
                                # 从另一个匹配列表中随机获取一个手机号码进行替换
                                # 如果replace_phone不为空，则表示已经替换过手机号码，不需要再次替换
                                if replace_phone == '':
                                    # 随机获取一个手机号码
                                    replace_phone = random.choice(replace_phone_list)
                    if phone != '' and  replace_phone != '':
                        # 若包含手机号码，则替换手机号码
                        if cell_value.find(phone) != -1:
                            # 替换手机号码
                            old_value = cell_value
                            cell_value = cell_value.replace(phone, replace_phone)
                            # 将替换后的数据写入新的excel文件中
                            sheet.cells(row+1, col+1).value = cell_value
                            print('替换成功: %s -> %s' % (old_value, cell_value))
                            continue
                continue
            # 将整行数据写入新的excel文件中

    # 保存excel文件
    workbook.save(new_file_name)
    app.quit() 

if __name__ == '__main__':
    # 指定目录
    path = r'/Users/karl/Work/eagleMobi/python/replaceNewMobile/xlsx_files'
    out_path = r'/Users/karl/Work/eagleMobi/python/replaceNewMobile/xlsx_files_out'
    # 指定手机号码列表
    phone_list = ['254703722167']
    # 指定替换手机号码列表
    replace_phone_list = ['254712366887', '254712366888', '254712366889', '254712366890', '254712366891']
    # 读取指定目录下的所有excel文件
    read_excel_file(path, out_path)
