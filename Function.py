import shutil
import os
import re
import time
import win32com.client as win32
import openpyxl
from datetime import date, datetime

today = datetime.now().strftime("%Y-%m-%d")
month = datetime.now().strftime("%m")
day = datetime.now().strftime("%d")


# 刷新表格数据包括powerquery里面生成的表
def excel_refresh(file_path):
    # file_path = r"D:\移动终端广分互联网\小时达\管理单元及门店管理\管理单元&门店明细.xlsx"

    # 打开 Excel 应用程序
    excel_app = win32.gencache.EnsureDispatch('Excel.Application')
    try:
        excel_app.Visible = False  # 隐藏 Excel 窗口

        # 打开工作簿
        workbook = excel_app.Workbooks.Open(file_path)

        # 刷新所有数据连接和 Power Query 查询
        workbook.RefreshAll()

        # 等待刷新完成
        excel_app.CalculateUntilAsyncQueriesDone()

        # 保存并关闭工作簿
        workbook.Save()
        workbook.Close()

    finally:
        # 退出 Excel 应用程序
        excel_app.Quit()


# 分别更新管理单元&门店明细表
def excel_update_data():
    # 指定 Excel 文件的路径
    source_file_path = fr"D:\移动终端广分互联网\小时达\管理单元及门店管理\管理单元&门店明细.xlsx"  # 使用原始字符串（r）来避免转义问题
    gldy_file_path = fr"D:\移动终端广分互联网\小时达\管理单元及门店管理\{today}\最新-管理单元&超管号未通过邀请明细{month}{day}.xlsx"
    md_file_path = fr"D:\移动终端广分互联网\小时达\管理单元及门店管理\{today}\最新-门店&管理号未通过邀请明细{month}{day}.xlsx"

    source_wb = openpyxl.load_workbook(source_file_path)
    gldy_wb = openpyxl.load_workbook(gldy_file_path)
    md_wb = openpyxl.load_workbook(md_file_path)

    # 获取指定工作表
    source_gldy_ws = source_wb["管理单元信息"]
    source_md_ws = source_wb["门店信息"]
    target_gldy_ws = gldy_wb.worksheets[0]
    target_md_ws = md_wb.worksheets[0]

    # 复制源工作表内容到目标工作表
    for row in source_gldy_ws.iter_rows(min_row=2):     # 管理单元表格更新
        for cell in row:
            target_gldy_ws[cell.coordinate].value = cell.value

    gldy_wb.save(gldy_file_path)

    for row in source_md_ws.iter_rows(min_row=2):  # 门店表格更新
        for cell in row:
            target_md_ws[cell.coordinate].value = cell.value

    md_wb.save(md_file_path)


# 将小时达入驻情况表内的公式转值并附上对应数量至表格标题
def excel_conversion():
    # 指定 Excel 文件的路径
    file_path = fr"D:\移动终端广分互联网\小时达\管理单元及门店管理\{today}\小时达入驻情况.xlsx"  # 使用原始字符串（r）来避免转义问题
    file_path_backup = fr"D:\移动终端广分互联网\小时达\管理单元及门店管理\{today}\小时达入驻情况{month}{day}.xlsx"  # 使用原始字符串（r）来避免转义问题

    # 打开 Excel 文件，设置 data_only=True 以便读取公式计算后的值
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        print(file_path)
    except FileNotFoundError:
        wb = openpyxl.load_workbook(file_path_backup, data_only=True)
        print(file_path_backup)

    # 遍历每个工作表
    for sheet in wb.sheetnames:
        ws = wb[sheet]

        # 遍历每个单元格
        for row in ws.iter_rows():
            for cell in row:    # 将单元格的值替换为公式计算后的值（如果有公式）
                if cell.data_type == 'f':  # 检查单元格是否包含公式
                    cell.value = cell.value  # 将公式替换为计算后的数值

    # 修改每个工作表名称，将门店数附在名称后面
    sheet_name = wb.sheetnames
    for sn in sheet_name:
        ws = wb[sn]
        sn = re.sub(r"\d*$", "", sn)    # 清除后面带有的数字
        # 判断汇总表定位合计，找到合计数
        if sn == "省区管理单元" or "省区门店":
            search_value = "合计"
            # 仅遍历A列（即第一列）
            for cell in ws.iter_rows(min_col=1, max_col=1):
                if cell[0].value == search_value:  # cell[0] 是 A 列的单元格
                    sum_cell = ws.cell(cell[0].row, cell[0].column + 2)    # 合计单元格向右偏移2列即为合计数
                    ws.title = sn + str(sum_cell.value)
                    print(ws.title)
                    break
        # 判断是否地市表，若是在表面加上行数即门店数
        if len(sn) == 2:    # 一般地市都是2个字
            ws.title = sn + str(ws.max_row - 1)
            print(ws.title)

    # 保存修改后的文件
    wb.save(file_path_backup)  # 保存为新文件或覆盖原文件


# 将未能匹配的管理单元&门店复制到匹配表待补充
def city_need_add():
    # 加载明细表和需要补充的地市表
    dt_path = r"D:\移动终端广分互联网\小时达\管理单元及门店管理\管理单元&门店明细.xlsx"
    sup_path = r"D:\移动终端广分互联网\小时达\管理单元及门店管理\数据源\管理单元&门店地市匹配.xlsx"
    dt_wb = openpyxl.load_workbook(dt_path)
    sup_wb = openpyxl.load_workbook(sup_path)

    # 分管理单元和门店分别进行补充
    dt_sheet_name = ["管理单元信息", "门店信息"]
    sup_sheet_name = ["管理单元地市匹配", "门店地市匹配"]

    for dt, sup in zip(dt_sheet_name, sup_sheet_name):
        # 获取明细工作表和匹配工作表
        source_ws = dt_wb[dt]
        target_ws = sup_wb[sup]

        # 获取匹配表的最大行，准备在末尾插入数据
        target_max_row = target_ws.max_row

        # 遍历明细表的所有行，筛选第三列为空白的行
        for row in source_ws.iter_rows(min_row=2):  # 从第二行开始，假设第一行是表头
            third_col_value = row[2].value  # 第三列数据
            if third_col_value is None or third_col_value == "":  # 判断是否为空
                # 获取前两列的数据
                first_col_value = row[0].value
                second_col_value = row[1].value

                # 在目标表最后一行后面插入前两列数据
                target_ws.cell(row=target_max_row + 1, column=1, value=first_col_value)
                target_ws.cell(row=target_max_row + 1, column=2, value=second_col_value)

                # 更新目标表的最大行
                target_max_row += 1

    # 保存文件
    sup_wb.save(r"D:\移动终端广分互联网\小时达\管理单元及门店管理\数据源\管理单元&门店地市匹配.xlsx")  # 保存为新文件，或覆盖原文件


# 复制一个当天的文件夹，并将改日期名称
def copy_folder():
    # 定义源文件夹路径和目标文件夹路径
    source_folder = r"D:\移动终端广分互联网\小时达\管理单元及门店管理\模板"  # 替换为你的源文件夹路径
    source_file = r"D:\移动终端广分互联网\小时达\管理单元及门店管理\小时达入驻情况.xlsx"
    target_folder = os.path.join(os.path.dirname(source_folder), today)  # 在同目录下创建目标文件夹

    # 复制模板文件夹并重命名为当天日期
    shutil.copytree(source_folder, target_folder)

    # 复制小时达入驻情况.xlsx到新文件夹
    shutil.copy(source_file, target_folder)

    # 更新表格名称
    file_keywords = ["最新-管理单元&超管号未通过邀请明细", "最新-门店&管理号未通过邀请明细"]    # 要匹配的关键词
    # 正则表达式，匹配以数字结尾的文件名
    date_pattern = re.compile(r"(\d{4})")  # 匹配文件名4位数字月日

    # 遍历文件夹中的文件
    for file_name in os.listdir(target_folder):
        # 构建完整文件路径
        old_file_path = os.path.join(target_folder, file_name)

        # 仅处理文件，排除文件夹
        if os.path.isfile(old_file_path):
            # 检查文件名是否包含指定关键词
            if any(keyword in file_name for keyword in file_keywords):
                # 使用正则表达式匹配文件名中的日期部分
                today_md = datetime.now().strftime("%m%d")
                new_file_name = date_pattern.sub(today_md, file_name)
                new_file_path = os.path.join(target_folder, new_file_name)

                # 重命名文件
                os.rename(old_file_path, new_file_path)
                print(f"文件已重命名为: {new_file_name}")


