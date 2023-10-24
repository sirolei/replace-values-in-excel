# step1 : 遍历指定目录下的excel文件
# step2 : 逐行读取excel文件中的数据
# step3 : 匹配每一行中是否出现格式为: Business Payment to 254712366882的数据，其中254712366882为手机号码
# step4 : 提取出手机号码，并判断是否在指定的手机号码列表中，若在列表中，则从另一个匹配列表中随机获取一个手机号码进行替换
# step5 : 将替换后的数据写入新的excel文件中
# step6 : 重复step1~step5，直到遍历完所有的excel文件

import os
import re
import random
from xlrd import open_workbook
import xlwt
import xlutils.copy
from xlutils.filter import process,XLRDReader,XLWTWriter
from xlutils.save import save

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
    workbook = open_workbook(file_path, formatting_info=True, on_demand=True)
    # 新的excel文件名,末尾加_new
    new_file_name = out_path.replace('.xls', '_new.xls')
    # 若已经存在_new的文件，则删除
    if os.path.exists(new_file_name):
        os.remove(new_file_name)
    # excel副本
    new_workbook = xlutils.copy.copy(workbook)

    # 获取excel文件中所有的sheet
    sheets = workbook.sheet_names()
    # 遍历所有的sheet
    for sheet in sheets:
        # 获取sheet
        sheet = workbook.sheet_by_name(sheet)
        # 获取sheet中的行数
        rows = sheet.nrows
        # 获取sheet中的列数
        cols = sheet.ncols
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
                cell_value = sheet.cell_value(row, col)
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
                            write_new_excel(new_workbook, sheet.name, row, col, cell_value)
                            print('替换成功: %s -> %s' % (old_value, cell_value))
                            continue
                # 保持原样写到新的excel文件中
                write_new_excel( new_workbook,sheet.name, row, col, cell_value)
                continue
            # 将整行数据写入新的excel文件中

    # 保存excel文件
    new_workbook.save(new_file_name) 


def _getOutCell(outSheet, colIndex, rowIndex):
    """ HACK: Extract the internal xlwt cell representation. """
    row = outSheet._Worksheet__rows.get(rowIndex)
    if not row: return None

    cell = row._Row__cells.get(colIndex)
    return cell
def setOutCell(outSheet, col, row, value):
    """ Change cell value without changing formatting. """
    # HACK to retain cell style.
    previousCell = _getOutCell(outSheet, col, row)
    # END HACK, PART I
    outSheet.write(row, col, value)    
    # HACK, PART II
    if previousCell:
        newCell = _getOutCell(outSheet, col, row)
        if newCell:
            newCell.xf_idx = previousCell.xf_idx
    # END HACK

def getStyleList(wb):
    w = XLWTWriter()
    process(XLRDReader(wb,'unknown.xls'),w)
    return w.style_list

# 将替换后的数据写入新的excel文件中
def write_new_excel(new_workbook, sheet_name, row, col, cell_value):
    
    
    # 获取sheet
    sheet = new_workbook.get_sheet(sheet_name)
    # 写入数据
    setOutCell(sheet, col, row, cell_value)

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
