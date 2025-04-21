# -*- coding: utf-8 -*-
"""
main.py

MediorNet TDM 连接计算器的主入口点。
负责初始化应用程序、加载样式、创建并显示主窗口。
"""

import sys
import os

# 导入 PySide6 组件
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt

# 导入重构后的主窗口类 (假设它在 ui 包中)
# 注意：需要确保 ui 目录在 Python 的搜索路径中，或者使用相对导入
# from ui.main_window import MainWindow # 如果 ui 是顶级目录
# 如果 core, ui, utils 与 main.py 在同一级，可以使用下面的方式
try:
    # 尝试从 ui 包导入
    from ui.main_window import MainWindow, APP_STYLE # 假设 APP_STYLE 也移到了 main_window
    print("成功从 ui.main_window 导入 MainWindow 和 APP_STYLE")
except ImportError as e1:
    print(f"尝试从 ui.main_window 导入失败: {e1}")
    try:
        # 如果失败，尝试假设 main_window.py 在当前目录 (用于简化初始运行)
        # 这不是最终结构，只是为了让第一步能运行
        from main_window import MainWindow, APP_STYLE # 临时方案
        print("警告: 从当前目录导入 MainWindow 和 APP_STYLE (临时方案)")
    except ImportError as e2:
         print(f"无法导入 MainWindow: {e2}. 请确保 main_window.py 文件存在且路径正确。")
         sys.exit(1) # 无法继续，退出

# --- 辅助函数 resource_path (暂时保留在这里，最终移到 utils) ---
def resource_path(relative_path):
    """ 获取资源的绝对路径，适用于开发环境和 PyInstaller 打包后 """
    try:
        # PyInstaller 创建一个临时文件夹并将路径存储在 _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # 未打包状态下，获取当前脚本所在目录
        try:
             # __file__ 在某些情况下可能未定义 (例如，在某些 IDE 的交互式控制台中)
             # 使用更健壮的方式获取基础路径
             if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
                 # PyInstaller 打包后的情况
                 base_path = sys._MEIPASS
             else:
                 # 普通 Python 环境
                 base_path = os.path.dirname(os.path.abspath(__file__))
        except NameError: # Fallback if __file__ is not defined
             base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
# --- 结束辅助函数 ---


# --- 程序入口 ---
if __name__ == "__main__":
    # 设置高 DPI 支持 (可选，根据需要取消注释)
    # try:
    #     # Windows specific AppUserModelID for taskbar grouping
    #     from PySide6 import QtWinExtras
    #     myappid = 'mycompany.myproduct.subproduct.version' # 替换为你自己的 ID
    #     QtWinExtras.QtWin.setCurrentProcessExplicitAppUserModelID(myappid)
    # except ImportError:
    #     pass # Not on Windows or QtWinExtras not available
    # QApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    # QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)

    # 创建 QApplication 实例
    app = QApplication(sys.argv)

    # 加载并应用 QSS 样式 (假设 APP_STYLE 在 MainWindow 中定义或导入)
    try:
        app.setStyleSheet(APP_STYLE)
        print("成功应用 QSS 样式。")
    except NameError:
        print("警告: 未找到 APP_STYLE 变量，无法应用样式。")
    except Exception as e:
        print(f"应用 QSS 样式时出错: {e}")


    # 创建主窗口实例
    try:
        window = MainWindow()
        print("MainWindow 实例已创建。")

        # 显示主窗口
        window.show()
        print("主窗口已显示。")

        # 启动应用程序事件循环
        sys.exit(app.exec())

    except Exception as e:
        print(f"创建或显示主窗口时发生严重错误: {e}")
        # 可以考虑显示一个简单的错误消息框
        # from PySide6.QtWidgets import QMessageBox
        # QMessageBox.critical(None, "启动错误", f"无法启动应用程序:\n{e}")
        sys.exit(1)

