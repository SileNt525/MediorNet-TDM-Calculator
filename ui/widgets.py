# -*- coding: utf-8 -*-
"""
ui/widgets.py

定义应用程序中使用的自定义 Qt 控件。
"""

from PySide6.QtWidgets import QTableWidgetItem
from PySide6.QtCore import Qt

class NumericTableWidgetItem(QTableWidgetItem):
    """
    自定义 QTableWidgetItem 以支持正确的数字排序。
    它重写了小于 (<) 操作符，尝试将单元格内容转换为数字进行比较。
    """
    def __lt__(self, other: QTableWidgetItem) -> bool:
        """
        重写小于比较操作符，用于表格排序。

        Args:
            other (QTableWidgetItem): 要比较的另一个单元格项。

        Returns:
            bool: 如果本项的数值小于另一项，则返回 True。
        """
        # 尝试从存储在 UserRole+1 中的数据获取数值（更可靠）
        data_self = self.data(Qt.ItemDataRole.UserRole + 1)
        data_other = other.data(Qt.ItemDataRole.UserRole + 1)

        try:
            # 优先使用存储的数值数据进行比较
            num_self = float(data_self)
            num_other = float(data_other)
            return num_self < num_other
        except (TypeError, ValueError):
            # 如果存储的数据无效或不存在，则回退到比较单元格显示的文本
            try:
                num_self = float(self.text())
                num_other = float(other.text())
                return num_self < num_other
            except ValueError:
                # 如果文本也无法转换为数字，则执行默认的字符串比较
                # 注意：这对于纯文本列是必要的
                return super().__lt__(other)

# 未来可以添加其他自定义控件到这个文件中，例如
# class DeviceTableWidget(QTableWidget):
#     # ... 实现特定于设备列表的功能 ...
#     pass
