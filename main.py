import Function
import DownloadFile
import PopUp_Confirmation as pc
from ExcelToImg import filter_and_save_visible_as_image as save_img
from datetime import date
import sys


# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    # 下载4个数据源文件,并覆盖数据源文件夹对应文件
    DownloadFile.export_data()

    # 刷新两个经过PowerQuery处理过的明细表
    file_path1 = r"D:\移动终端广分互联网\小时达\管理单元及门店管理\管理单元&门店明细.xlsx"
    file_path2 = r"D:\移动终端广分互联网\小时达\管理单元及门店管理\小时达入驻情况.xlsx"

    # 刷新表格并打印状态
    Function.excel_refresh(file_path1)
    print("完成《管理单元&门店明细》表格刷新")
    Function.excel_refresh(file_path2)
    print("完成《小时达入驻情况》表格刷新")

    # 将匹配不上地市的管理单元&门店及其id复制到匹配表
    Function.city_need_add()
    print("完成《管理单元&门店地市匹配》表格刷新")

    # 确保在所有操作完成后再弹出确认框
    confirmation_result = pc.city_confirmation()  # 确认结果存储在变量中

    # 弹窗等待用户确认补充完成
    if confirmation_result:
        print("已完成地市补充，继续执行后续步骤……")

        # 补充完成后重新刷新两个明细表
        Function.excel_refresh(file_path1)
        print("完成《管理单元&门店明细》表格二次刷新")
        Function.excel_refresh(file_path2)
        print("完成《小时达入驻情况》表格二次刷新")

        # 配置当天数据的文件夹
        Function.copy_folder()
        print("完成当天文件夹配置")

        # 更新当天文件夹内管理单元&门店明细的数据表数据
        Function.excel_update_data()
        print("完成文件夹内管理单元&门店明细的数据表刷新")

        # 操作小时达入驻情况表
        Function.excel_conversion()
        print("完成当天小时达入驻表格操作")

        # 将处理好的小时达入驻情况筛选并保存为图片
        source_path = fr"D:\移动终端广分互联网\小时达\管理单元及门店管理\{date.today()}"
        sheet_name = {"省区管理单元": "上级入驻情况", "省区门店": "门店入驻情况"}
        column_letter = "B"  # 要筛选颜色的列
        color_rgb = (254, 219, 97)  # 橙色的 RGB 值

        save_img(source_path, sheet_name, column_letter, color_rgb)

        # 后续继续完成将处理好的文件发送到微信群的进程

    else:
        print("用户取消，流程中止。")
        sys.exit("流程已被用户终止。")


