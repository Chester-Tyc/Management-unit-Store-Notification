import xlwings as xw
from PIL import ImageGrab
import os
from datetime import datetime


def filter_and_save_visible_as_image(source_path, sheet_name, column_letter, color_rgb):
    # 日期格式化
    month = datetime.now().strftime("%m")
    day = datetime.now().strftime("%d")
    # 文件路径格式化
    file_path = os.path.join(source_path, "小时达入驻情况.xlsx")
    app = xw.App(visible=False)

    try:
        # 打开 Excel 文件
        wb = app.books.open(file_path)
        os.makedirs(source_path, exist_ok=True)  # 确保源文件夹存在

        # 遍历所有工作表，找到目标表
        for sheet in wb.sheets:
            for keyword, image_name in sheet_name.items():
                if sheet.name == keyword:  # 匹配省区管理单元和省区门店表
                    print(f"正在处理表格: {sheet.name}")

                    # 遍历 B 列，从第2行开始，隐藏无颜色填充的行
                    for cell in sheet.range(f"{column_letter}2:{column_letter}{sheet.cells.last_cell.row}"):
                        if cell.offset(0, -1).value == "合计":
                            last_cell = sheet.range(sheet.used_range.last_cell.row, sheet.used_range.last_cell.column)
                            print(f"右下角单元格位置: {last_cell.address}, 值: {last_cell.value}")
                            break
                        if cell.color != color_rgb:
                            cell.api.EntireRow.Hidden = True  # 使用 api 来隐藏整行
                    print(f" {sheet.name}处理完成")

                    # 将筛选结果复制为图片
                    sheet.range(f"A1:{last_cell.address}").api.CopyPicture(Format=2)
                    img = ImageGrab.grabclipboard()

                    # 构建图片保存路径并保存图片
                    output_path = os.path.join(source_path, f"{image_name}{month}{day}.jpg")
                    if img:  # 确保截图成功
                        img.save(output_path, "JPEG")
                        print(f"筛选结果已保存为 {output_path}")
                    else:
                        print("未能获取截图，请确认表格内容或截图区域。")

                    # 恢复所有隐藏行
                    sheet.api.Rows.Hidden = False

    except Exception as e:
        print("Error processing file:", e)
    finally:
        # 清理并退出 Excel 应用
        wb.close()
        app.quit()

