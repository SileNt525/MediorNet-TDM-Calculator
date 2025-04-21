# -*- coding: utf-8 -*-
"""
controllers/topology_controller.py

定义 TopologyController 类，负责处理拓扑图画布的用户交互逻辑，
例如节点选择、拖动、双击详情、Shift+拖拽连接等。
"""

import copy
from typing import Optional, Dict, Tuple, TYPE_CHECKING, Any

# PySide6 imports
from PySide6.QtCore import QObject, Slot, Qt # <--- **修复: 添加了 Qt 导入**
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import QMessageBox # 导入 QMessageBox

# Matplotlib imports
from matplotlib.lines import Line2D

# NetworkX import (可能需要，例如检查图属性)
import networkx as nx

# 项目模块导入 (使用字符串进行类型提示以避免循环导入)
if TYPE_CHECKING:
    from core.network_manager import NetworkManager, ConnectionType
    from core.device import Device
    # !! 修改: 不再直接导入 MainWindow，使用字符串提示 !!
    # from ui.main_window import MainWindow
    from ui.topology_canvas import MplCanvas

class TopologyController(QObject):
    """处理拓扑图画布交互的控制器。"""

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
        # 类型断言仅用于静态分析，运行时不强制执行
        if TYPE_CHECKING:
             # 仅在类型检查时导入，避免运行时循环导入
             # 确保导入路径相对于项目根目录或已添加到 PYTHONPATH
             try:
                 from ui.main_window import MainWindow
                 assert isinstance(main_window, MainWindow)
             except ImportError: pass # 忽略类型检查时的导入错误
             try:
                 from core.network_manager import NetworkManager
                 assert isinstance(network_manager, NetworkManager)
             except ImportError: pass
             try:
                 from ui.topology_canvas import MplCanvas
                 assert isinstance(mpl_canvas, MplCanvas)
             except ImportError: pass


        self.main_window = main_window
        self.network_manager = network_manager
        self.mpl_canvas = mpl_canvas

        # --- 从 MainWindow 移动过来的状态变量 ---
        self.node_positions: Optional[Dict[int, Tuple[float, float]]] = None # 存储节点布局 {node_id: (x, y)}
        self.selected_node_id: Optional[int] = None # 当前选中的节点 ID
        self.dragged_node_id: Optional[int] = None # 当前正在拖动的节点 ID
        self.drag_offset: Tuple[float, float] = (0, 0)   # 拖动时的偏移量
        self.connecting_node_id: Optional[int] = None # Shift 拖动时起始节点的 ID
        self.connection_line: Optional[Line2D] = None # Shift 拖动时绘制的临时线对象

    # --- 画布事件处理方法 (从 MainWindow 移动过来) ---

    def _get_node_at_event(self, event) -> Optional[int]:
        """辅助函数：查找鼠标事件位置下的节点 ID。"""
        if event.inaxes != self.mpl_canvas.axes or not self.node_positions:
            return None
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return None

        clicked_node_id = None
        min_dist_sq = float('inf')

        xlim = self.mpl_canvas.axes.get_xlim()
        ylim = self.mpl_canvas.axes.get_ylim()
        threshold_dist_sq = ((xlim[1]-xlim[0])**2 + (ylim[1]-ylim[0])**2) * (0.03**2)

        for node_id, (nx, ny) in self.node_positions.items():
            dist_sq = (x - nx)**2 + (y - ny)**2
            if dist_sq < min_dist_sq and dist_sq < threshold_dist_sq:
                min_dist_sq = dist_sq
                clicked_node_id = node_id
        return clicked_node_id

    def _start_node_drag(self, node_id: int, event_xdata: float, event_ydata: float):
        """辅助函数：开始节点拖动。"""
        if self.node_positions is None or node_id not in self.node_positions:
            print(f"警告: 尝试拖动未找到位置的节点 ID {node_id}")
            return

        self.dragged_node_id = node_id
        self.connecting_node_id = None
        nx, ny = self.node_positions[node_id]
        self.drag_offset = (event_xdata - nx, event_ydata - ny)

        if self.selected_node_id != node_id:
            self.selected_node_id = node_id
            print(f"选中节点 (准备拖动): ID={self.selected_node_id}")
            self.main_window._update_connection_views()
        else:
            print(f"开始拖动节点: ID={self.selected_node_id}")

    def _start_connection_drag(self, node_id: int):
        """辅助函数：开始连接拖动 (Shift+Click)。"""
        if self.node_positions is None or node_id not in self.node_positions:
             print(f"警告: 尝试从未知位置的节点 ID {node_id} 开始连接拖动")
             return

        print(f"开始连接拖动: 从 ID={node_id}")
        self.connecting_node_id = node_id
        self.dragged_node_id = None

        if self.selected_node_id != node_id:
            self.selected_node_id = node_id
            self.main_window._update_connection_views()

    def _handle_background_press(self):
        """辅助函数：处理画布背景点击，清除选择和拖动状态。"""
        needs_redraw = self.selected_node_id is not None or \
                       self.dragged_node_id is not None or \
                       self.connecting_node_id is not None

        self.selected_node_id = None
        self.dragged_node_id = None
        self.connecting_node_id = None

        if needs_redraw:
            print("清除选中/状态 (点击背景)")
            self.main_window._update_connection_views()

    def _end_node_drag(self):
        """辅助函数：结束节点拖动。"""
        if self.dragged_node_id is not None:
            print(f"结束拖动节点: ID={self.dragged_node_id}")
            # 更新 MainWindow 中的 node_positions 副本 (如果 MainWindow 还需要它)
            # 或者让 MainWindow 在需要时从 Controller 获取
            # !! 注意: MainWindow 现在应该通过 self.topology_controller.node_positions 访问 !!
            # self.main_window.node_positions = self.node_positions # 这行可能不再需要，取决于 MainWindow 的实现
            self.dragged_node_id = None
        else:
             print("调试: _end_node_drag 被调用但 self.dragged_node_id 为 None")


    def _end_connection_drag(self, event):
        """辅助函数：结束连接拖动并尝试通过 NetworkManager 创建连接。"""
        start_node_id = self.connecting_node_id
        self.connecting_node_id = None # 重置连接拖动状态

        # 清理画布上的临时连接线
        if self.connection_line:
            try:
                self.connection_line.remove()
            except ValueError:
                pass
            finally:
                 self.connection_line = None
            self.mpl_canvas.draw_idle()

        if start_node_id is None:
            print("调试: _end_connection_drag 启动节点 ID 为 None")
            return

        target_node_id = self._get_node_at_event(event)

        if target_node_id is not None and target_node_id != start_node_id:
            # 调用 NetworkManager 的 add_best_connection 方法
            print(f"尝试通过拖拽连接: ID {start_node_id} -> ID {target_node_id}")
            added_connection = self.network_manager.add_best_connection(start_node_id, target_node_id)

            if added_connection:
                 # 连接成功，更新 UI (通过 MainWindow)
                 # !! 修改: 让 MainWindow 从 Controller 获取 node_positions !!
                 self.node_positions = None # Controller 重置自己的布局状态
                 self.selected_node_id = None # Controller 清除自己的选择状态
                 self.main_window._update_connection_views() # 通知 MainWindow 更新视图
                 self.main_window._update_device_table_connections()
                 self.main_window._update_manual_port_options()
                 self.main_window._update_port_totals_display()
                 # 更新填充按钮状态 (需要访问 MainWindow 的按钮)
                 has_connections = bool(self.network_manager.get_all_connections())
                 can_fill = has_connections or any(bool(dev.get_all_available_ports()) for dev in self.network_manager.get_all_devices())
                 self.main_window.fill_mesh_button.setEnabled(can_fill)
                 self.main_window.fill_ring_button.setEnabled(can_fill)
                 print(f"成功通过拖拽添加连接: {added_connection[0].name}[{added_connection[1]}] <-> {added_connection[2].name}[{added_connection[3]}]")
            else:
                 # 连接失败 (NetworkManager 内部已打印原因)
                 dev1 = self.network_manager.get_device_by_id(start_node_id)
                 dev2 = self.network_manager.get_device_by_id(target_node_id)
                 dev1_name = dev1.name if dev1 else f"ID {start_node_id}"
                 dev2_name = dev2.name if dev2 else f"ID {target_node_id}"
                 # 通过 MainWindow 显示消息框
                 QMessageBox.warning(self.main_window, "连接失败", f"无法在 {dev1_name} 和 {dev2_name} 之间自动添加连接（可能无可用兼容端口）。")
                 self.main_window._update_manual_port_options() # 刷新端口列表
        else:
            print("连接拖动取消或目标无效/相同。")


    # --- 作为槽函数连接到 MplCanvas 信号 ---
    @Slot(object)
    def on_canvas_press(self, event):
        """处理画布上的鼠标按下事件。"""
        if self.connection_line:
            try: self.connection_line.remove(); self.connection_line = None
            except ValueError: pass

        if event.inaxes != self.mpl_canvas.axes or event.xdata is None or event.ydata is None:
            self._handle_background_press()
            return

        clicked_node_id = self._get_node_at_event(event)
        modifiers = QGuiApplication.keyboardModifiers()
        # !! 使用导入的 Qt !!
        is_shift_pressed = modifiers == Qt.KeyboardModifier.ShiftModifier

        if event.dblclick:
            self.dragged_node_id = None
            self.connecting_node_id = None
            if clicked_node_id is not None:
                device = self.network_manager.get_device_by_id(clicked_node_id)
                if device:
                    # 通过 MainWindow 显示详情
                    self.main_window._display_device_details_popup(device)
        elif event.button == 1:
            if clicked_node_id is not None:
                if is_shift_pressed:
                    self._start_connection_drag(clicked_node_id)
                else:
                    self._start_node_drag(clicked_node_id, event.xdata, event.ydata)
            else:
                self._handle_background_press()

    @Slot(object)
    def on_canvas_motion(self, event):
        """处理画布上的鼠标移动事件。"""
        if event.inaxes != self.mpl_canvas.axes or event.xdata is None or event.ydata is None:
            return

        x, y = event.xdata, event.ydata

        # 处理节点拖动
        if self.dragged_node_id is not None and event.button == 1 and self.node_positions:
            if self.dragged_node_id in self.node_positions:
                 new_x = x - self.drag_offset[0]
                 new_y = y - self.drag_offset[1]
                 self.node_positions[self.dragged_node_id] = (new_x, new_y)
                 # 通过 MainWindow 更新视图
                 # !! MainWindow 现在应直接从 Controller 获取 node_positions !!
                 # self.main_window.topology_controller.node_positions = self.node_positions # 这行可能多余
                 self.main_window._update_connection_views()
            else:
                 print(f"警告: 尝试拖动节点 {self.dragged_node_id} 但其不在 node_positions 中")
                 self.dragged_node_id = None

        # 处理连接线拖动
        elif self.connecting_node_id is not None and event.button == 1 and self.node_positions:
            start_pos = self.node_positions.get(self.connecting_node_id)
            if start_pos is not None:
                if self.connection_line:
                    try: self.connection_line.remove(); self.connection_line = None
                    except ValueError: pass
                    except AttributeError: self.connection_line = None
                self.connection_line = Line2D([start_pos[0], x], [start_pos[1], y],
                                              ls='--', c='gray', lw=1.5,
                                              transform=self.mpl_canvas.axes.transData,
                                              zorder=10)
                self.mpl_canvas.axes.add_line(self.connection_line)
                self.mpl_canvas.draw_idle()
            else:
                 print(f"警告: 尝试绘制连接线但起始节点 {self.connecting_node_id} 位置未知")
                 self.connecting_node_id = None

    @Slot(object)
    def on_canvas_release(self, event):
        """处理画布上的鼠标释放事件。"""
        if event.button == 1:
            if self.dragged_node_id is not None:
                self._end_node_drag()
            elif self.connecting_node_id is not None:
                self._end_connection_drag(event)

