# -*- coding: utf-8 -*-
"""
controllers/topology_controller.py

定义 TopologyController 类，负责处理拓扑图画布的用户交互逻辑，
例如节点选择、拖动、双击详情、Shift+拖拽连接等。
使用信号/槽机制与 MainWindow 解耦。
"""

import copy
from typing import Optional, Dict, Tuple, TYPE_CHECKING, Any

# PySide6 imports
from PySide6.QtCore import QObject, Slot, Qt, Signal # <--- **修复: 添加了 Qt, Signal 导入**
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import QMessageBox # 导入 QMessageBox

# Matplotlib imports
from matplotlib.lines import Line2D

# NetworkX import (可能需要，例如检查图属性)
import networkx as nx

# 项目模块导入 (使用字符串进行类型提示以避免循环导入)
if TYPE_CHECKING:
    # 仅在类型检查时导入
    from core.network_manager import NetworkManager, ConnectionType
    from core.device import Device
    from ui.main_window import MainWindow
    from ui.topology_canvas import MplCanvas

class TopologyController(QObject):
    """处理拓扑图画布交互的控制器。"""

    # --- 定义信号 ---
    view_needs_update = Signal()
    request_ui_update = Signal()
    request_device_details = Signal(object)
    connection_attempt_failed = Signal(str, str)
    enable_fill_buttons = Signal(bool)

    def __init__(self, main_window: 'MainWindow', network_manager: 'NetworkManager', mpl_canvas: 'MplCanvas', parent: Optional[QObject] = None):
        """
        初始化 TopologyController。

        Args:
            main_window ('MainWindow'): 主窗口实例的引用 (使用字符串类型提示)。
            network_manager ('NetworkManager'): 网络管理器实例的引用。
            mpl_canvas ('MplCanvas'): Matplotlib 画布实例的引用。
            parent (Optional[QObject]): 父对象 (通常为 None 或 main_window)。
        """
        super().__init__(parent)
        # !! 修改: 移除 TYPE_CHECKING 和断言，简化初始化 !!
        self.main_window = main_window
        self.network_manager = network_manager
        self.mpl_canvas = mpl_canvas

        # --- 从 MainWindow 移动过来的状态变量 ---
        self.node_positions: Optional[Dict[int, Tuple[float, float]]] = None
        self.selected_node_id: Optional[int] = None
        self.dragged_node_id: Optional[int] = None
        self.drag_offset: Tuple[float, float] = (0, 0)
        self.connecting_node_id: Optional[int] = None
        self.connection_line: Optional[Line2D] = None

    # --- 公共方法 (供 MainWindow 获取状态) ---
    # !! 新增 Getter 方法 !!
    def get_node_positions(self) -> Optional[Dict[int, Tuple[float, float]]]:
        """返回当前存储的节点位置。"""
        return self.node_positions

    def get_selected_node_id(self) -> Optional[int]:
        """返回当前选中的节点 ID。"""
        return self.selected_node_id

    def reset_layout_state(self):
        """重置布局相关的状态（例如，当布局算法改变或清除操作时）。"""
        print("Controller: 重置布局状态")
        self.node_positions = None
        self.selected_node_id = None
        self.dragged_node_id = None
        self.connecting_node_id = None
        self.view_needs_update.emit()


    # --- 画布事件处理方法 (从 MainWindow 移动过来) ---

    def _get_node_at_event(self, event) -> Optional[int]:
        """辅助函数：查找鼠标事件位置下的节点 ID。"""
        if event.inaxes != self.mpl_canvas.axes or not self.node_positions:
            return None
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return None
        clicked_node_id = None; min_dist_sq = float('inf')
        xlim = self.mpl_canvas.axes.get_xlim(); ylim = self.mpl_canvas.axes.get_ylim()
        threshold_dist_sq = ((xlim[1]-xlim[0])**2 + (ylim[1]-ylim[0])**2) * (0.03**2)
        for node_id, (nx, ny) in self.node_positions.items():
            dist_sq = (x - nx)**2 + (y - ny)**2
            if dist_sq < min_dist_sq and dist_sq < threshold_dist_sq:
                min_dist_sq = dist_sq; clicked_node_id = node_id
        return clicked_node_id

    def _start_node_drag(self, node_id: int, event_xdata: float, event_ydata: float):
        """辅助函数：开始节点拖动。"""
        if self.node_positions is None or node_id not in self.node_positions:
            print(f"警告: 尝试拖动未找到位置的节点 ID {node_id}"); return
        self.dragged_node_id = node_id; self.connecting_node_id = None
        nx, ny = self.node_positions[node_id]
        self.drag_offset = (event_xdata - nx, event_ydata - ny)
        if self.selected_node_id != node_id:
            self.selected_node_id = node_id
            print(f"选中节点 (准备拖动): ID={self.selected_node_id}")
            self.view_needs_update.emit() # 发射信号
        else: print(f"开始拖动节点: ID={self.selected_node_id}")

    def _start_connection_drag(self, node_id: int):
        """辅助函数：开始连接拖动 (Shift+Click)。"""
        if self.node_positions is None or node_id not in self.node_positions:
             print(f"警告: 尝试从未知位置的节点 ID {node_id} 开始连接拖动"); return
        print(f"开始连接拖动: 从 ID={node_id}")
        self.connecting_node_id = node_id; self.dragged_node_id = None
        if self.selected_node_id != node_id:
            self.selected_node_id = node_id
            self.view_needs_update.emit() # 发射信号

    def _handle_background_press(self):
        """辅助函数：处理画布背景点击，清除选择和拖动状态。"""
        needs_update = self.selected_node_id is not None or self.dragged_node_id is not None or self.connecting_node_id is not None
        self.selected_node_id = None; self.dragged_node_id = None; self.connecting_node_id = None
        if needs_update:
            print("清除选中/状态 (点击背景)")
            self.view_needs_update.emit() # 发射信号

    def _end_node_drag(self):
        """辅助函数：结束节点拖动。"""
        if self.dragged_node_id is not None:
            print(f"结束拖动节点: ID={self.dragged_node_id}")
            # 不需要同步回 MainWindow，MainWindow 会通过 getter 获取
            self.dragged_node_id = None
        else: print("调试: _end_node_drag 被调用但 self.dragged_node_id 为 None")

    def _end_connection_drag(self, event):
        """辅助函数：结束连接拖动并尝试通过 NetworkManager 创建连接。"""
        start_node_id = self.connecting_node_id; self.connecting_node_id = None
        if self.connection_line:
            try: self.connection_line.remove()
            except ValueError: pass
            finally: self.connection_line = None
            self.mpl_canvas.draw_idle()
        if start_node_id is None: print("调试: _end_connection_drag 启动节点 ID 为 None"); return

        target_node_id = self._get_node_at_event(event)
        if target_node_id is not None and target_node_id != start_node_id:
            print(f"尝试通过拖拽连接: ID {start_node_id} -> ID {target_node_id}")
            added_connection = self.network_manager.add_best_connection(start_node_id, target_node_id)
            if added_connection:
                 self.node_positions = None; self.selected_node_id = None
                 print(f"成功通过拖拽添加连接: {added_connection[0].name}[{added_connection[1]}] <-> {added_connection[2].name}[{added_connection[3]}]")
                 self.request_ui_update.emit() # 发射信号
                 has_connections = bool(self.network_manager.get_all_connections())
                 can_fill = has_connections or any(bool(dev.get_all_available_ports()) for dev in self.network_manager.get_all_devices())
                 self.enable_fill_buttons.emit(can_fill) # 发射信号
            else:
                 dev1 = self.network_manager.get_device_by_id(start_node_id); dev2 = self.network_manager.get_device_by_id(target_node_id)
                 dev1_name = dev1.name if dev1 else f"ID {start_node_id}"; dev2_name = dev2.name if dev2 else f"ID {target_node_id}"
                 self.connection_attempt_failed.emit(dev1_name, dev2_name) # 发射信号
                 self.request_ui_update.emit() # 发射信号 (更新端口)
        else:
            print("连接拖动取消或目标无效/相同。")
            self.view_needs_update.emit() # 确保清除选择状态


    # --- 作为槽函数连接到 MplCanvas 信号 ---
    @Slot(object)
    def on_canvas_press(self, event):
        """处理画布上的鼠标按下事件。"""
        if self.connection_line:
            try: self.connection_line.remove(); self.connection_line = None
            except ValueError: pass
        if event.inaxes != self.mpl_canvas.axes or event.xdata is None or event.ydata is None:
            self._handle_background_press(); return

        clicked_node_id = self._get_node_at_event(event)
        modifiers = QGuiApplication.keyboardModifiers()
        is_shift_pressed = modifiers == Qt.KeyboardModifier.ShiftModifier

        if event.dblclick:
            self.dragged_node_id = None; self.connecting_node_id = None
            if clicked_node_id is not None:
                device = self.network_manager.get_device_by_id(clicked_node_id)
                if device: self.request_device_details.emit(device) # 发射信号
        elif event.button == 1:
            if clicked_node_id is not None:
                if is_shift_pressed: self._start_connection_drag(clicked_node_id)
                else: self._start_node_drag(clicked_node_id, event.xdata, event.ydata)
            else: self._handle_background_press()

    @Slot(object)
    def on_canvas_motion(self, event):
        """处理画布上的鼠标移动事件。"""
        if event.inaxes != self.mpl_canvas.axes or event.xdata is None or event.ydata is None: return
        x, y = event.xdata, event.ydata

        if self.dragged_node_id is not None and event.button == 1 and self.node_positions:
            if self.dragged_node_id in self.node_positions:
                 new_x = x - self.drag_offset[0]; new_y = y - self.drag_offset[1]
                 self.node_positions[self.dragged_node_id] = (new_x, new_y)
                 self.view_needs_update.emit() # 发射信号
            else: print(f"警告: 尝试拖动节点 {self.dragged_node_id} 但其不在 node_positions 中"); self.dragged_node_id = None
        elif self.connecting_node_id is not None and event.button == 1 and self.node_positions:
            start_pos = self.node_positions.get(self.connecting_node_id)
            if start_pos is not None:
                if self.connection_line:
                    try: self.connection_line.remove(); self.connection_line = None
                    except ValueError: pass
                    except AttributeError: self.connection_line = None
                self.connection_line = Line2D([start_pos[0], x], [start_pos[1], y], ls='--', c='gray', lw=1.5, transform=self.mpl_canvas.axes.transData, zorder=10)
                self.mpl_canvas.axes.add_line(self.connection_line)
                self.mpl_canvas.draw_idle()
            else: print(f"警告: 尝试绘制连接线但起始节点 {self.connecting_node_id} 位置未知"); self.connecting_node_id = None

    @Slot(object)
    def on_canvas_release(self, event):
        """处理画布上的鼠标释放事件。"""
        if event.button == 1:
            if self.dragged_node_id is not None: self._end_node_drag()
            elif self.connecting_node_id is not None: self._end_connection_drag(event)

