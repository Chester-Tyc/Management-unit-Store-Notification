import tkinter as tk
from tkinter import messagebox
import sys


def city_confirmation():
    # 创建一个Tk窗口，但不显示
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 弹出确认窗口
    response = messagebox.askokcancel("地市补充", "请完成管理单元&门店地市补充，完成后请点击已完成")

    # 销毁根窗口，关闭窗口
    root.destroy()
    return response


# # 弹出确认窗口，等待用户确认
# if show_confirmation():
#     print("已完成地市补充，继续执行后续步骤……")
#
# else:
#     print("用户取消，流程中止。")
#     sys.exit("流程已被用户终止。")
