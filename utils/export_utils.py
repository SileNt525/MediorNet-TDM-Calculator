# -*- coding: utf-8 -*-
"""
utils/export_utils.py

提供导出连接列表、拓扑图和 HTML 报告的功能函数。
"""

import os
import base64
import io
import datetime
from typing import List, Tuple, TYPE_CHECKING, Optional, Any # 导入 Any

# 导入必要的 Qt 组件
from PySide6.QtWidgets import QFileDialog, QMessageBox, QWidget

# 导入 Matplotlib Figure 类型 (仅用于类型提示)
try:
    from matplotlib.figure import Figure
except ImportError:
    Figure = object # Fallback 类型

# 类型提示：避免循环导入，仅在类型检查时导入
if TYPE_CHECKING:
    # 仅在类型检查时导入，避免运行时循环导入
    # 使用字符串形式进行前向引用，Pylance 通常能更好地处理
    from core.device import Device
    ConnectionType = Tuple['Device', str, 'Device', str, str]
else:
    # 在运行时，定义一个足够使用的占位符类型或直接使用 Any
    # 注意：如果 ConnectionType 在运行时逻辑中需要被实例化或检查，
    # 则需要一个更具体的定义或从 network_manager 导入。
    # 但在这个导出工具模块中，主要用于类型提示，用 Any 或简单 Tuple 即可。
    ConnectionType = Tuple[Any, str, Any, str, str]


# 注意：将类型提示改为字符串形式 'QWidget', 'Optional[Figure]', 'List[ConnectionType]'
def export_connections_to_file(parent_window: 'QWidget', connections: List['ConnectionType']):
    """
    将连接列表导出为 TXT 或 CSV 文件。

    Args:
        parent_window ('QWidget'): 父窗口，用于 QFileDialog 和 QMessageBox。
        connections (List['ConnectionType']): 要导出的连接列表。
    """
    if not connections:
        QMessageBox.warning(parent_window, "提示", "没有连接结果可导出。")
        return

    filepath, selected_filter = QFileDialog.getSaveFileName(
        parent_window,
        "导出连接列表",
        "", # 默认目录
        "文本文件 (*.txt);;CSV 文件 (*.csv);;所有文件 (*)"
    )

    if not filepath:
        print("用户取消导出连接列表。")
        return

    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            is_csv = "csv" in selected_filter.lower()
            if is_csv:
                # 写入 CSV 头部
                f.write("序号,设备1,端口1,设备2,端口2,连接类型\n")
            else:
                 # 写入 TXT 头部 (可选)
                 f.write("MediorNet 连接列表\n")
                 f.write("=" * 30 + "\n")

            # 写入连接数据
            for i, conn in enumerate(connections):
                # 从元组中解包设备对象和端口信息
                # 注意：conn[0] 和 conn[2] 是 Device 对象 (或运行时为 Any)
                dev1, port1, dev2, port2, conn_type = conn
                # 安全地访问 name 属性，以防运行时类型为 Any
                dev1_name = getattr(dev1, 'name', 'Unknown Device')
                dev2_name = getattr(dev2, 'name', 'Unknown Device')
                if is_csv:
                    # CSV 格式，使用逗号分隔
                    f.write(f"{i+1},{dev1_name},{port1},{dev2_name},{port2},{conn_type}\n")
                else:
                    # TXT 格式
                    f.write(f"{i+1}. {dev1_name} [{port1}] <-> {dev2_name} [{port2}] ({conn_type})\n")

        QMessageBox.information(parent_window, "成功", f"连接列表已导出到:\n{filepath}")
        print(f"连接列表成功导出到: {filepath}")

    except Exception as e:
        QMessageBox.critical(parent_window, "导出失败", f"无法导出连接列表:\n{e}")
        print(f"导出连接列表错误: {e}")


def export_topology_to_file(parent_window: 'QWidget', figure: Optional['Figure']):
    """
    将 Matplotlib 绘制的拓扑图导出为图像文件 (PNG, PDF, SVG)。

    Args:
        parent_window ('QWidget'): 父窗口。
        figure (Optional['Figure']): Matplotlib 的 Figure 对象。
    """
    if not figure:
        QMessageBox.warning(parent_window, "提示", "没有拓扑图可导出。")
        return

    filepath, selected_filter = QFileDialog.getSaveFileName(
        parent_window,
        "导出拓扑图",
        "",
        "PNG 图像 (*.png);;PDF 文件 (*.pdf);;SVG 文件 (*.svg);;所有文件 (*)"
    )

    if not filepath:
        print("用户取消导出拓扑图。")
        return

    try:
        # 使用 Figure 对象保存图像
        figure.savefig(filepath, dpi=300, bbox_inches='tight')
        QMessageBox.information(parent_window, "成功", f"拓扑图已导出到:\n{filepath}")
        print(f"拓扑图成功导出到: {filepath}")

    except Exception as e:
        QMessageBox.critical(parent_window, "导出失败", f"无法导出拓扑图:\n{e}")
        print(f"导出拓扑图错误: {e}")


def export_report_to_html(parent_window: 'QWidget', figure: Optional['Figure'], connections: List['ConnectionType']):
    """
    导出包含拓扑图和连接列表的 HTML 报告。

    Args:
        parent_window ('QWidget'): 父窗口。
        figure (Optional['Figure']): Matplotlib 的 Figure 对象。
        connections (List['ConnectionType']): 连接列表。
    """
    if not figure:
         QMessageBox.warning(parent_window, "无法导出", "缺少拓扑图数据。")
         return
    # 注意：这里允许没有连接但有图的情况导出（只包含图）

    filepath, _ = QFileDialog.getSaveFileName(
        parent_window,
        "导出 HTML 报告",
        "",
        "HTML 文件 (*.html);;所有文件 (*)"
    )

    if not filepath:
        print("用户取消导出 HTML 报告。")
        return

    try:
        # 1. 将拓扑图保存到内存缓冲区并编码为 Base64
        buffer = io.BytesIO()
        figure.savefig(buffer, format='png', dpi=150, bbox_inches='tight') # 使用稍低 DPI 以减小文件大小
        buffer.seek(0)
        image_base64 = base64.b64encode(buffer.getvalue()).decode('utf-8')
        img_data_uri = f"data:image/png;base64,{image_base64}"
        buffer.close()

        # 2. 构建连接列表的 HTML 表格
        connections_table_html = """
        <div class="mt-8">
          <h2 class="text-lg font-semibold mb-3 text-gray-700">连接列表</h2>
          <div class="overflow-x-auto bg-white rounded-lg shadow">
            <table class="min-w-full leading-normal">
              <thead>
                <tr>
                  <th class="px-5 py-3 border-b-2 border-gray-200 bg-gray-100 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">序号</th>
                  <th class="px-5 py-3 border-b-2 border-gray-200 bg-gray-100 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">设备 1</th>
                  <th class="px-5 py-3 border-b-2 border-gray-200 bg-gray-100 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">端口 1</th>
                  <th class="px-5 py-3 border-b-2 border-gray-200 bg-gray-100 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">设备 2</th>
                  <th class="px-5 py-3 border-b-2 border-gray-200 bg-gray-100 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">端口 2</th>
                  <th class="px-5 py-3 border-b-2 border-gray-200 bg-gray-100 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">类型</th>
                </tr>
              </thead>
              <tbody>
        """
        if connections:
            for i, conn in enumerate(connections):
                dev1, port1, dev2, port2, conn_type = conn
                # 安全地访问 name 和 type 属性
                dev1_name = getattr(dev1, 'name', 'N/A')
                dev1_type = getattr(dev1, 'type', 'N/A')
                dev2_name = getattr(dev2, 'name', 'N/A')
                dev2_type = getattr(dev2, 'type', 'N/A')
                # 添加斑马纹背景
                bg_class = "bg-white" if i % 2 == 0 else "bg-gray-50"
                connections_table_html += f"""
                        <tr class="{bg_class}">
                          <td class="px-5 py-4 border-b border-gray-200 text-sm">{i+1}</td>
                          <td class="px-5 py-4 border-b border-gray-200 text-sm">{dev1_name} ({dev1_type})</td>
                          <td class="px-5 py-4 border-b border-gray-200 text-sm">{port1}</td>
                          <td class="px-5 py-4 border-b border-gray-200 text-sm">{dev2_name} ({dev2_type})</td>
                          <td class="px-5 py-4 border-b border-gray-200 text-sm">{port2}</td>
                          <td class="px-5 py-4 border-b border-gray-200 text-sm">{conn_type}</td>
                        </tr>
                """
        else:
            connections_table_html += """
                        <tr>
                          <td colspan="6" class="px-5 py-5 border-b border-gray-200 bg-white text-center text-sm text-gray-500">无连接</td>
                        </tr>
            """
        connections_table_html += """
              </tbody>
            </table>
          </div>
        </div>
        """

        # 3. 构建完整的 HTML 内容
        # !! Fix: Escape curly braces in the HTML comment !!
        html_content = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MediorNet 连接报告</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        /* 基础字体和打印样式优化 */
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, "Noto Sans", sans-serif, "Apple Color Emoji", "Segoe UI Emoji", "Segoe UI Symbol", "Noto Color Emoji";
        }}
        @media print {{
            body {{
                -webkit-print-color-adjust: exact; /* Chrome, Safari */
                print-color-adjust: exact; /* Firefox */
            }}
            /* 确保背景色和边框在打印时可见 */
            .bg-gray-100 {{ background-color: #f7fafc !important; }}
            .bg-gray-50 {{ background-color: #f9fafb !important; }}
            .border-b-2 {{ border-bottom-width: 2px !important; }}
            .border-gray-200 {{ border-color: #edf2f7 !important; }}
            .shadow, .shadow-xl, .shadow-inner {{ box-shadow: none !important; }}
            /* 可以考虑移除页面边距以更好地利用纸张 */
            .container {{ margin: 0 !important; padding: 10px !important; max-width: 100% !important; }}
            h1, h2 {{ margin-bottom: 1rem !important; }}
            table {{ width: 100% !important; }}
            img {{ max-width: 90% !important; display: block; margin-left: auto; margin-right: auto; }} /* 居中并缩小图像以防溢出 */
        }}
    </style>
</head>
<body class="bg-gray-100">
    <div class="container mx-auto p-6 md:p-10 bg-white rounded-lg shadow-xl my-10 max-w-6xl"> {{/* 稍微加宽容器 */}}
        <h1 class="text-2xl font-bold text-center mb-8 text-gray-700">MediorNet TDM 连接报告</h1>

        <div class="mb-8">
            <h2 class="text-lg font-semibold mb-3 text-gray-600">网络连接拓扑图</h2>
            <div class="flex justify-center p-4 border border-gray-200 rounded-lg bg-gray-50 shadow-inner">
                <img src="{img_data_uri}" alt="网络拓扑图" style="max-width: 100%; height: auto;" class="rounded">
            </div>
        </div>

        {connections_table_html}

        <div class="text-center text-xs text-gray-400 mt-10 pt-4 border-t border-gray-200">
            报告生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        </div>
    </div>
</body>
</html>
        """ # 确保 f-string 的结束 """ 在正确的位置

        # 4. 写入文件
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(html_content)

        QMessageBox.information(parent_window, "导出成功", f"HTML 报告已成功导出到:\n{filepath}")
        print(f"HTML 报告成功导出到: {filepath}")

    except Exception as e:
        QMessageBox.critical(parent_window, "导出失败", f"导出 HTML 报告时发生错误:\n{e}")
        print(f"导出 HTML 报告错误: {e}")

