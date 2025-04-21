# -*- coding: utf-8 -*-
"""
utils/misc_utils.py

包含通用的辅助函数。
"""

import sys
import os

def resource_path(relative_path: str) -> str:
    """
    获取资源的绝对路径，适用于开发环境和 PyInstaller 打包后的环境。

    Args:
        relative_path (str): 相对于基础路径的资源文件路径。

    Returns:
        str: 资源的绝对路径。
    """
    try:
        # PyInstaller 创建一个临时文件夹并将路径存储在 sys._MEIPASS
        # getattr 用于安全地访问属性，防止在非 PyInstaller 环境下出错
        base_path = getattr(sys, '_MEIPASS', None)
        if base_path is None:
            # 如果不是通过 PyInstaller 运行，获取当前脚本的目录
            # 使用 os.path.abspath 和 os.path.dirname(__file__) 获取脚本所在目录
            # __file__ 在某些执行方式下可能未定义 (例如，交互式解释器)
            try:
                 script_dir = os.path.dirname(os.path.abspath(__file__))
            except NameError:
                 script_dir = os.path.abspath(".") # Fallback 到当前工作目录
            base_path = script_dir
    except Exception as e:
        print(f"获取基础路径时出错: {e}，回退到当前工作目录。")
        # 最终回退到当前工作目录
        base_path = os.path.abspath(".")

    # 拼接基础路径和相对路径
    return os.path.join(base_path, relative_path)

# 未来可以添加其他通用辅助函数到这里，例如：
# def format_timestamp(ts):
#     # ... 实现时间戳格式化 ...
#     pass
