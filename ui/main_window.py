# -*- coding: utf-8 -*-
"""
ui/main_window.py

定义主应用程序窗口 MainWindow 类。
此类现在使用 Ui_MainWindow 来设置 UI，并包含事件处理逻辑。
"""

import sys
import os
import copy
import json
import base64
import io
import datetime
import random
from collections import defaultdict
from typing import Optional, List, Dict, Tuple, Set, Any

# PySide6 imports
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QTextEdit,
    QTabWidget, QFrame, QFileDialog, QMessageBox, QSpacerItem, QSizePolicy,
    QGridLayout, QListWidgetItem, QAbstractItemView, QTableWidget,
    QTableWidgetItem, QHeaderView, QListWidget, QSplitter, QCheckBox
)
from PySide6.QtCore import Slot, Qt
from PySide6.QtGui import QFont, QGuiApplication, QFontDatabase

# Matplotlib imports
from matplotlib import font_manager
import matplotlib.pyplot as plt
# from matplotlib.lines import Line2D # 由 Controller 管理

# NetworkX import
import networkx as nx

# --- 从项目模块导入 ---
try:
    from core.network_manager import NetworkManager, ConnectionType
    from core.device import (
        Device,
        DEV_UHD, DEV_HORIZON, DEV_MN, UHD_TYPES,
        PORT_MPO, PORT_LC, PORT_SFP, PORT_UNKNOWN,
        get_port_type_from_name
    )
    from .topology_canvas import MplCanvas
    from .widgets import NumericTableWidgetItem
    from controllers.topology_controller import TopologyController
    from .ui_main_window import Ui_MainWindow # <--- 导入 UI 定义类
    from utils.export_utils import export_connections_to_file, export_topology_to_file, export_report_to_html
    from utils.misc_utils import resource_path
except ImportError as e:
    print(f"导入错误 (main_window.py): {e} - 请确保所有模块已正确创建。")
    # Fallbacks
    NetworkManager = object; Device = object; ConnectionType = tuple
    DEV_UHD, DEV_HORIZON, DEV_MN, UHD_TYPES = '', '', '', []; PORT_MPO, PORT_LC, PORT_SFP, PORT_UNKNOWN = '', '', '', ''
    get_port_type_from_name = lambda x: ''; MplCanvas = QWidget; NumericTableWidgetItem = QTableWidgetItem
    TopologyController = object; Ui_MainWindow = object
    export_connections_to_file = lambda *args, **kwargs: None; export_topology_to_file = lambda *args, **kwargs: None; export_report_to_html = lambda *args, **kwargs: None
    resource_path = lambda x: x

# --- UI 常量 ---
COL_NAME = 0; COL_TYPE = 1; COL_MPO = 2; COL_LC = 3; COL_SFP = 4; COL_CONN = 5

# --- QSS 样式定义 ---
APP_STYLE = """
QMainWindow, QDialog, QMessageBox { }
QFrame { background-color: transparent; }
QFrame#addDeviceGroup, QFrame#listGroup, QFrame#fileGroup,
QFrame#calculateControlFrame, QFrame#addManualGroup, QFrame#removeManualGroup {
    border: 1px solid #cccccc; border-radius: 5px; margin-bottom: 5px;
}
QPushButton { border: 1px solid #bbbbbb; padding: 5px 10px; border-radius: 4px; min-height: 20px; min-width: 75px; }
QPushButton:hover { border: 1px solid #999999; }
QPushButton:pressed { border: 1px solid #777777; }
QPushButton:disabled { border: 1px solid #dddddd; }
QLineEdit, QComboBox, QTextEdit, QListWidget, QTableWidget { border: 1px solid #cccccc; border-radius: 3px; padding: 3px; }
QComboBox::drop-down { border: none; background: transparent; width: 15px; padding-right: 5px; }
QComboBox::down-arrow { width: 12px; height: 12px; }
QTableWidget { gridline-color: #e0e0e0; }
QHeaderView::section { padding: 4px; border: 1px solid #cccccc; border-left: none; font-weight: bold; }
QHeaderView::section:first { border-left: 1px solid #cccccc; }
QTabWidget::pane { border: 1px solid #cccccc; border-top: none; }
QTabBar::tab { border: 1px solid #cccccc; border-bottom: none; padding: 6px 12px; border-top-left-radius: 4px; border-top-right-radius: 4px; margin-right: 2px; }
QTabBar::tab:selected { margin-bottom: -1px; }
QSplitter::handle { background-color: #e0e0e0; border: none; }
QSplitter::handle:horizontal { width: 3px; }
QSplitter::handle:vertical { height: 3px; }
QSplitter::handle:hover { background-color: #d0d0d0; }
"""

# --- 主窗口类 ---
class MainWindow(QMainWindow):
    """应用程序的主窗口，包含所有 UI 元素和交互逻辑。"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MediorNet TDM 连接计算器 V1.4 (UI Separated)")
        self.setGeometry(100, 100, 1100, 800)

        # --- 核心数据管理器 ---
        self.network_manager = NetworkManager()

        # --- UI 状态变量 ---
        self.suppress_confirmations: bool = False

        # --- 字体加载 ---
        self.chinese_font = self._setup_fonts() # 需要先加载字体，setupUi 会用到

        # --- UI 定义和布局 ---
        self.ui = Ui_MainWindow()
        # 创建中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        # 创建 MplCanvas 实例 (setupUi 需要 MainWindow 有此属性)
        self.mpl_canvas = MplCanvas(central_widget)
        # 调用 setupUi 来构建界面 (它会将控件添加到 self 上)
        self.ui.setupUi(self)

        # --- 控制器 ---
        self.topology_controller = TopologyController(self, self.network_manager, self.mpl_canvas)

        # --- 连接信号 ---
        self._connect_ui_signals()
        self._connect_controller_signals()
        self._connect_canvas_signals()

        # --- 初始化 UI 状态 ---
        self.update_port_entries()
        self._update_port_totals_display()
        self._update_connection_views()
        self._update_device_combos()

    def _setup_fonts(self) -> QFont:
        """加载并设置字体，返回 QFont 对象。"""
        # (此方法内容不变)
        default_font = QFont(); default_font.setPointSize(10)
        try:
            font_relative_path = os.path.join('assets', 'NotoSansCJKsc-Regular.otf'); font_path = resource_path(font_relative_path)
            if os.path.exists(font_path):
                font_id = QFontDatabase.addApplicationFont(font_path)
                if font_id != -1:
                    families = QFontDatabase.applicationFontFamilies(font_id)
                    if families:
                        family = families[0]; print(f"成功加载并设置字体: {family} (路径: {font_path})")
                        font_manager.fontManager.addfont(font_path); plt.rcParams['font.sans-serif'] = [family] + plt.rcParams.get('font.sans-serif', []); plt.rcParams['axes.unicode_minus'] = False
                        return QFont(family, 10)
                    else: print(f"警告: 无法从字体文件获取 family name {font_path}")
                else: print(f"警告: 添加字体失败 {font_path}")
            else: print(f"警告: 未在路径 {font_path} 找到字体文件。将使用默认字体。")
        except ImportError: print("警告: 无法导入 QFontDatabase。将使用默认字体。")
        except Exception as e: print(f"加载或设置字体时出错: {e}")
        plt.rcParams['font.sans-serif'] = ['sans-serif']; plt.rcParams['axes.unicode_minus'] = False
        return default_font

    # !! 删除 _setup_ui 方法 !!

    def _connect_ui_signals(self):
        """连接 UI 控件的信号到 MainWindow 的槽函数。"""
        # !! 修改: 访问 self.xxx 而不是 self.ui.xxx !!
        self.device_type_combo.currentIndexChanged.connect(self.update_port_entries)
        self.add_button.clicked.connect(self.add_device)
        self.device_filter_entry.textChanged.connect(self.filter_device_table)
        self.device_tablewidget.itemDoubleClicked.connect(self.show_device_details_from_table)
        self.device_tablewidget.itemChanged.connect(self.on_device_item_changed)
        self.remove_button.clicked.connect(self.remove_device)
        self.clear_button.clicked.connect(self.clear_all_devices)
        self.save_button.clicked.connect(self.save_config)
        self.load_button.clicked.connect(self.load_config)
        self.export_list_button.clicked.connect(self.export_connections)
        self.export_topo_button.clicked.connect(self.export_topology)
        self.export_report_button.clicked.connect(self.export_html_report)
        self.suppress_confirm_checkbox.stateChanged.connect(self._toggle_suppress_confirmations)
        self.layout_combo.currentIndexChanged.connect(self.on_layout_change)
        self.calculate_button.clicked.connect(self.calculate_and_display)
        self.fill_mesh_button.clicked.connect(self.fill_remaining_mesh)
        self.fill_ring_button.clicked.connect(self.fill_remaining_ring)
        self.edit_dev1_combo.currentIndexChanged.connect(self._update_manual_port_options)
        self.edit_port1_combo.currentIndexChanged.connect(self._update_manual_port_options)
        self.edit_dev2_combo.currentIndexChanged.connect(self._update_manual_port_options)
        self.edit_port2_combo.currentIndexChanged.connect(self._update_manual_port_options)
        self.add_manual_button.clicked.connect(self.add_manual_connection)
        self.conn_filter_type_combo.currentIndexChanged.connect(self.filter_connection_list)
        self.conn_filter_device_entry.textChanged.connect(self.filter_connection_list)
        self.remove_manual_button.clicked.connect(self.remove_manual_connection)
        print("成功连接 UI 控件信号。")

    def _connect_controller_signals(self):
        """连接 TopologyController 的信号到 MainWindow 的槽函数。"""
        # (此方法内容不变)
        try:
            if not isinstance(self.topology_controller, TopologyController):
                 print("错误: TopologyController 实例无效，无法连接信号。")
                 try: from controllers.topology_controller import TopologyController as ActualController; assert isinstance(self.topology_controller, ActualController)
                 except (ImportError, AssertionError): raise TypeError("self.topology_controller 不是有效的 TopologyController 实例。")
            self.topology_controller.view_needs_update.connect(self._update_connection_views)
            self.topology_controller.request_device_details.connect(self._display_device_details_popup)
            self.topology_controller.request_ui_update.connect(self._full_ui_update_after_action)
            self.topology_controller.connection_attempt_failed.connect(self._show_connection_failure_message)
            self.topology_controller.enable_fill_buttons.connect(self._set_fill_buttons_enabled)
            print("成功连接 Controller 信号。")
        except AttributeError as e: print(f"严重错误: 连接 Controller 信号时发生属性错误: {e}"); QMessageBox.critical(self, "初始化错误", f"连接控制器信号失败: {e}\n请检查控制台输出。")
        except Exception as e: print(f"连接 Controller 信号时出错: {e}"); QMessageBox.critical(self, "初始化错误", f"连接控制器信号时发生未知错误: {e}")

    def _connect_canvas_signals(self):
        """连接 MplCanvas 的信号到 TopologyController 的槽函数。"""
        # (此方法内容不变)
        print("-" * 20); print(f"DEBUG: Connecting mpl signals..."); print(f"DEBUG: self.mpl_canvas type: {type(self.mpl_canvas)}"); print(f"DEBUG: self.topology_controller type: {type(self.topology_controller)}")
        try:
            press_slot = getattr(self.topology_controller, 'on_canvas_press', None); motion_slot = getattr(self.topology_controller, 'on_canvas_motion', None); release_slot = getattr(self.topology_controller, 'on_canvas_release', None)
            print(f"DEBUG: press_slot: {press_slot}"); print(f"DEBUG: motion_slot: {motion_slot}"); print(f"DEBUG: release_slot: {release_slot}")
            if not all([press_slot, motion_slot, release_slot]): raise AttributeError("一个或多个 TopologyController 槽函数未找到！")
            cid_press = self.mpl_canvas.mpl_connect('button_press_event', press_slot); cid_motion = self.mpl_canvas.mpl_connect('motion_notify_event', motion_slot); cid_release = self.mpl_canvas.mpl_connect('button_release_event', release_slot)
            print(f"DEBUG: mpl_connect calls executed. CIDs: {cid_press}, {cid_motion}, {cid_release}")
            def _debug_mpl_event(event): print(f"DEBUG (MainWindow): Matplotlib event received: {event.name}, button={event.button}, xdata={event.xdata}, ydata={event.ydata}")
            self._debug_event_cid = self.mpl_canvas.mpl_connect('button_press_event', _debug_mpl_event); print(f"DEBUG (MainWindow): Connected debug handler with ID: {self._debug_event_cid}")
        except Exception as e: print(f"!!! 严重错误: mpl_connect 失败: {e} !!!"); QMessageBox.critical(self, "错误", f"无法连接画布事件处理器: {e}")
        print("-" * 20)


    # --- 新增的槽函数，用于响应 Controller 信号 ---

    @Slot()
    def _full_ui_update_after_action(self):
        """响应 Controller 请求，更新多个相关的 UI 部件。"""
        print("槽函数: _full_ui_update_after_action 被调用")
        self._update_connection_views() # 更新图形和列表
        self._update_device_table_connections() # 更新表格连接数
        self._update_manual_port_options() # 更新手动编辑端口
        self._update_port_totals_display() # 更新总数标签

    @Slot(str, str)
    def _show_connection_failure_message(self, dev1_name: str, dev2_name: str):
        """响应 Controller 信号，显示连接失败的消息框。"""
        print(f"槽函数: _show_connection_failure_message 被调用 ({dev1_name}, {dev2_name})")
        QMessageBox.warning(self, "连接失败", f"无法在 {dev1_name} 和 {dev2_name} 之间自动添加连接（可能无可用兼容端口）。")

    @Slot(bool)
    def _set_fill_buttons_enabled(self, enabled: bool):
        """响应 Controller 信号，设置填充按钮的可用状态。"""
        print(f"槽函数: _set_fill_buttons_enabled 被调用 (enabled={enabled})")
        # !! 修改: 使用 self.xxx !!
        self.fill_mesh_button.setEnabled(enabled)
        self.fill_ring_button.setEnabled(enabled)


    # --- 现有槽函数 (访问控件改为 self.xxx) ---

    @Slot()
    def update_port_entries(self):
        """根据设备类型选择更新端口输入框的可见性。"""
        # !! 修改: 使用 self.xxx !!
        selected_type = self.device_type_combo.currentText()
        is_micron = selected_type == DEV_MN; is_uhd_horizon = selected_type in UHD_TYPES
        self.mpo_label.setVisible(is_uhd_horizon); self.mpo_entry.setVisible(is_uhd_horizon)
        self.lc_label.setVisible(is_uhd_horizon); self.lc_entry.setVisible(is_uhd_horizon)
        self.sfp_label.setVisible(is_micron); self.sfp_entry.setVisible(is_micron)

    @Slot()
    def add_device(self):
        """处理“添加设备”按钮点击事件。"""
        # !! 修改: 使用 self.xxx !!
        dtype = self.device_type_combo.currentText(); name = self.device_name_entry.text().strip()
        if not dtype: QMessageBox.critical(self, "错误", "请选择设备类型。"); return
        if not name: QMessageBox.critical(self, "错误", "请输入设备名称。"); return
        mpo_ports_str = self.mpo_entry.text() or "0"; lc_ports_str = self.lc_entry.text() or "0"; sfp_ports_str = self.sfp_entry.text() or "0"
        mpo_ports, lc_ports, sfp_ports = 0, 0, 0
        try:
            if dtype in UHD_TYPES: mpo_ports, lc_ports = int(mpo_ports_str), int(lc_ports_str); assert mpo_ports >= 0 and lc_ports >= 0
            elif dtype == DEV_MN: sfp_ports = int(sfp_ports_str); assert sfp_ports >= 0
            else: raise ValueError("无效类型")
            new_device = self.network_manager.add_device(name, dtype, mpo_ports, lc_ports, sfp_ports)
            if new_device:
                self._add_device_to_table(new_device); self.device_name_entry.clear(); self._update_device_combos()
                self.clear_results(); self._update_port_totals_display()
            else: QMessageBox.critical(self, "错误", f"无法添加设备 '{name}' (可能名称已存在)。")
        except (ValueError, AssertionError): QMessageBox.critical(self, "输入错误", "端口数量必须是非负整数。")
        except Exception as e: QMessageBox.critical(self, "添加失败", f"添加设备时发生未知错误: {e}")

    @Slot()
    def remove_device(self):
        """处理“移除选中”按钮点击事件。"""
        # !! 修改: 使用 self.xxx !!
        selected_rows = sorted(list(set(index.row() for index in self.device_tablewidget.selectedIndexes())), reverse=True)
        if not selected_rows: QMessageBox.warning(self, "提示", "请先在表格中选择要移除的设备行。"); return
        ids_to_remove = {self.device_tablewidget.item(r, COL_NAME).data(Qt.ItemDataRole.UserRole) for r in selected_rows if self.device_tablewidget.item(r, COL_NAME)}
        names_to_remove = [self.device_tablewidget.item(r, COL_NAME).text() for r in selected_rows if self.device_tablewidget.item(r, COL_NAME)]
        if not ids_to_remove: return
        user_confirmed = True
        if not self.suppress_confirmations:
            reply = QMessageBox.question(self, "确认移除", f"确定要移除选中的设备及其所有连接吗？\n({', '.join(names_to_remove)})", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            user_confirmed = (reply == QMessageBox.StandardButton.Yes)
        if user_confirmed:
            removed_count = sum(1 for dev_id in ids_to_remove if self.network_manager.remove_device(dev_id))
            if removed_count > 0:
                for row_index in selected_rows: self.device_tablewidget.removeRow(row_index)
                self._update_device_combos(); self.topology_controller.reset_layout_state(); self._update_port_totals_display()
                print(f"成功移除了 {removed_count} 个设备及其连接。")
            else: print("没有设备被移除。")
        else: print("用户取消移除设备。")

    @Slot()
    def clear_all_devices(self):
        """处理“清空所有”按钮点击事件。"""
        if not self.network_manager.get_all_devices(): return
        user_confirmed = True
        if not self.suppress_confirmations:
            reply = QMessageBox.question(self, "确认清空", "确定要清空所有设备和连接吗？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            user_confirmed = (reply == QMessageBox.StandardButton.Yes)
        if user_confirmed:
            self.network_manager.clear_all_devices_and_connections()
            self.device_tablewidget.setRowCount(0) # !! 修改: 使用 self.xxx !!
            self._update_device_combos(); self.topology_controller.reset_layout_state(); self._update_port_totals_display()
            self._set_fill_buttons_enabled(False); print("所有设备和连接已清空。")
        else: print("用户取消清空所有设备。")

    @Slot()
    def clear_results(self):
        """清除计算结果和连接，重置设备端口状态和画布状态。"""
        self.network_manager.clear_connections()
        self.topology_controller.reset_layout_state()
        # !! 修改: 使用 self.xxx !!
        self.connections_textedit.clear(); self.connections_textedit.append("无连接。")
        self.manual_connection_list.clear()
        # 清空画布由 Controller 的 reset_layout_state 触发的 view_needs_update 信号处理
        self.export_list_button.setEnabled(False); self.export_topo_button.setEnabled(False); self.export_report_button.setEnabled(False)
        self.remove_manual_button.setEnabled(False); self._set_fill_buttons_enabled(False)
        self._update_device_table_connections(); self._update_port_totals_display(); self._update_manual_port_options()
        print("计算结果和连接已清除。")

    @Slot()
    def calculate_and_display(self):
        """处理“计算连接”按钮点击事件。"""
        devices = self.network_manager.get_all_devices()
        if not devices: QMessageBox.information(self, "提示", "请先添加设备。"); return
        self.network_manager.clear_connections()
        mode = self.topology_mode_combo.currentText() # !! 修改: 使用 self.xxx !!
        calculated_connections_data: List[ConnectionType] = []; error_message = None
        if mode == "Mesh": calculated_connections_data = self.network_manager.calculate_mesh()
        elif mode == "环形": calculated_connections_data, error_message = self.network_manager.calculate_ring()
        else: QMessageBox.critical(self, "错误", f"未知的计算模式: {mode}"); return
        if error_message: QMessageBox.warning(self, f"{mode} 计算警告", error_message)
        added_count = 0
        if calculated_connections_data:
            print(f"计算得到 {len(calculated_connections_data)} 条连接，正在添加到管理器...")
            for conn_tuple_data in calculated_connections_data:
                dev1, port1, dev2, port2, _ = conn_tuple_data
                if self.network_manager.add_connection(dev1.id, port1, dev2.id, port2): added_count += 1
                else: print(f"警告: 将计算出的连接 {dev1.name}[{port1}]<->{dev2.name}[{port2}] 添加到管理器时失败。")
            print(f"成功添加了 {added_count} 条计算出的连接到管理器。")
        else: print("计算未产生任何连接。")
        self.topology_controller.reset_layout_state()
        self._update_device_table_connections(); self._update_device_combos(); self._update_manual_port_options()
        has_connections = bool(self.network_manager.get_all_connections())
        can_fill = has_connections or any(bool(dev.get_all_available_ports()) for dev in devices)
        self._set_fill_buttons_enabled(can_fill)
        # !! 修改: 使用 self.xxx !!
        self.export_list_button.setEnabled(has_connections); self.export_topo_button.setEnabled(bool(devices)); self.export_report_button.setEnabled(has_connections and bool(devices)); self.remove_manual_button.setEnabled(has_connections)

    @Slot()
    def fill_remaining_mesh(self):
        """处理“填充 (Mesh)”按钮点击事件。"""
        if not self.network_manager.get_all_devices(): QMessageBox.information(self, "提示", "请先添加设备。"); return
        print("开始填充剩余连接 (Mesh)...")
        new_connections = self.network_manager.fill_connections_mesh()
        if new_connections:
            self.topology_controller.reset_layout_state()
            self._update_device_table_connections(); self._update_manual_port_options(); self._update_port_totals_display()
            QMessageBox.information(self, "填充完成", f"成功添加了 {len(new_connections)} 条新 Mesh 连接。")
        else: QMessageBox.information(self, "填充完成", "没有找到更多可以建立的 Mesh 连接。")
        self._set_fill_buttons_enabled(False)

    @Slot()
    def fill_remaining_ring(self):
        """处理“填充 (环形)”按钮点击事件。"""
        if not self.network_manager.get_all_devices(): QMessageBox.information(self, "提示", "请先添加设备。"); return
        print("开始填充剩余连接 (环形)...")
        new_connections = self.network_manager.fill_connections_ring()
        if new_connections:
            self.topology_controller.reset_layout_state()
            self._update_device_table_connections(); self._update_manual_port_options(); self._update_port_totals_display()
            QMessageBox.information(self, "填充完成", f"成功添加了 {len(new_connections)} 条新环形连接段。")
        else: QMessageBox.information(self, "填充完成", "没有找到更多可以建立的环形连接段。")
        self._set_fill_buttons_enabled(False)

    @Slot()
    def save_config(self):
        """处理“保存配置”按钮点击事件。"""
        if not self.network_manager.get_all_devices(): QMessageBox.warning(self, "提示", "设备列表为空，无需保存。"); return
        filepath, _ = QFileDialog.getSaveFileName(self, "保存项目配置", "", "JSON 文件 (*.json);;所有文件 (*)")
        if not filepath: return
        if self.network_manager.save_project(filepath): QMessageBox.information(self, "成功", f"项目配置已保存到:\n{filepath}")
        else: QMessageBox.critical(self, "保存失败", f"无法保存项目配置文件。")

    @Slot()
    def load_config(self):
        """处理“加载配置”按钮点击事件。"""
        user_confirmed_load = True
        if self.network_manager.get_all_devices():
            if not self.suppress_confirmations:
                reply = QMessageBox.question(self, "确认加载", "加载配置将覆盖当前所有设备和连接，确定吗？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
                user_confirmed_load = (reply == QMessageBox.StandardButton.Yes)
        if not user_confirmed_load: print("用户取消加载配置。"); return
        filepath, _ = QFileDialog.getOpenFileName(self, "加载项目配置", "", "JSON 文件 (*.json);;所有文件 (*)")
        if not filepath: return
        if self.network_manager.load_project(filepath):
            self.device_tablewidget.setRowCount(0) # !! 修改: 使用 self.xxx !!
            for dev in self.network_manager.get_all_devices(): self._add_device_to_table(dev)
            self._update_device_combos(); self.topology_controller.reset_layout_state(); self._update_port_totals_display()
            has_connections = bool(self.network_manager.get_all_connections())
            can_fill = has_connections or any(bool(dev.get_all_available_ports()) for dev in self.network_manager.get_all_devices())
            self._set_fill_buttons_enabled(can_fill); QMessageBox.information(self, "成功", f"项目配置已从以下文件加载:\n{filepath}")
        else:
            self.device_tablewidget.setRowCount(0); self._update_device_combos() # !! 修改: 使用 self.xxx !!
            self.topology_controller.reset_layout_state(); self._update_port_totals_display(); self._set_fill_buttons_enabled(False)
            QMessageBox.critical(self, "加载失败", f"无法加载项目配置文件:\n{filepath}")

    @Slot()
    def export_connections(self):
        """处理“导出列表”按钮点击事件。"""
        connections = self.network_manager.get_all_connections()
        if not connections: QMessageBox.warning(self, "提示", "没有连接结果可导出。"); return
        export_connections_to_file(self, connections)

    @Slot()
    def export_topology(self):
        """处理“导出拓扑图”按钮点击事件。"""
        figure_to_export = self.mpl_canvas.fig # 从 MplCanvas 获取 fig 对象
        if not figure_to_export or not self.network_manager.get_all_devices(): QMessageBox.warning(self, "提示", "没有拓扑图可导出。"); return
        export_topology_to_file(self, figure_to_export)

    @Slot()
    def export_html_report(self):
        """处理“导出报告 (HTML)”按钮点击事件。"""
        devices = self.network_manager.get_all_devices(); connections = self.network_manager.get_all_connections()
        figure_to_export = self.mpl_canvas.fig # 从 MplCanvas 获取 fig 对象
        if not devices or not figure_to_export: QMessageBox.warning(self, "无法导出", "请先添加设备并生成拓扑图。"); return
        export_report_to_html(self, figure_to_export, connections)

    @Slot()
    def add_manual_connection(self):
        """处理手动编辑标签页中的“添加连接”按钮。"""
        # !! 修改: 使用 self.xxx !!
        dev1_id = self.edit_dev1_combo.currentData(); dev2_id = self.edit_dev2_combo.currentData()
        port1_text = self.edit_port1_combo.currentText(); port2_text = self.edit_port2_combo.currentText()
        if dev1_id is None or dev2_id is None or port1_text == "选择端口..." or port2_text == "选择端口...": QMessageBox.warning(self, "选择不完整", "请选择两个设备和它们各自的端口。"); return
        if dev1_id == dev2_id: QMessageBox.warning(self, "选择错误", "不能将设备连接到自身。"); return
        added_connection = self.network_manager.add_connection(dev1_id, port1_text, dev2_id, port2_text)
        if added_connection:
            self.topology_controller.reset_layout_state()
            self._update_device_table_connections(); self._update_manual_port_options(); self._update_port_totals_display()
            self._set_fill_buttons_enabled(True); print(f"成功添加手动连接: {added_connection[0].name}[{port1_text}] <-> {added_connection[2].name}[{port2_text}]")
        else:
            QMessageBox.warning(self, "添加失败", "无法添加手动连接，请检查端口兼容性、可用性或查看控制台输出。")
            self._update_manual_port_options()

    @Slot()
    def remove_manual_connection(self):
        """处理手动编辑标签页中的“移除选中连接”按钮。"""
        # !! 修改: 使用 self.xxx !!
        selected_items = self.manual_connection_list.selectedItems()
        if not selected_items: QMessageBox.warning(self, "提示", "请在下方列表中选择要移除的连接。"); return
        successfully_removed_items = []
        removed_count = 0
        for item in selected_items:
            conn_data = item.data(Qt.ItemDataRole.UserRole)
            if conn_data:
                dev1, port1, dev2, port2, _ = conn_data
                if self.network_manager.remove_connection(dev1.id, port1, dev2.id, port2):
                    removed_count += 1; successfully_removed_items.append(item)
                else: print(f"警告: 尝试从管理器移除连接时失败: {dev1.name}[{port1}] <-> {dev2.name}[{port2}]")
        if removed_count > 0:
            rows_to_remove = sorted([self.manual_connection_list.row(item) for item in successfully_removed_items], reverse=True)
            for row in rows_to_remove:
                if row != -1: taken_item = self.manual_connection_list.takeItem(row)
            self.topology_controller.reset_layout_state()
            self._update_device_table_connections(); self._update_manual_port_options(); self._update_port_totals_display()
            print(f"成功移除了 {removed_count} 条连接。")
            has_connections = bool(self.network_manager.get_all_connections())
            can_fill = has_connections or any(bool(dev.get_all_available_ports()) for dev in self.network_manager.get_all_devices())
            self._set_fill_buttons_enabled(can_fill); self.remove_manual_button.setEnabled(has_connections)
        else: print("没有连接被移除。")

    @Slot(int)
    def _toggle_suppress_confirmations(self, state):
        """更新是否跳过确认弹窗的状态。"""
        self.suppress_confirmations = (state == Qt.CheckState.Checked.value)
        print(f"跳过确认弹窗: {'已启用' if self.suppress_confirmations else '已禁用'}")

    @Slot(QTableWidgetItem)
    def on_device_item_changed(self, item: QTableWidgetItem):
        """处理设备表格中项目被编辑后的事件。"""
        # !! 修改: 使用 self.xxx !!
        if not item: return
        row = item.row(); col = item.column(); name_item = self.device_tablewidget.item(row, COL_NAME)
        if not name_item: return
        dev_id = name_item.data(Qt.ItemDataRole.UserRole)
        if dev_id is None: return
        self.device_tablewidget.blockSignals(True)
        success = False
        try:
            if col == COL_NAME:
                new_name = item.text().strip()
                if self.network_manager.update_device(dev_id, new_name=new_name):
                     success = True; self._update_device_combos(); self._update_connection_views()
                else:
                     device = self.network_manager.get_device_by_id(dev_id);
                     if device: item.setText(device.name)
                     QMessageBox.warning(self, "重命名失败", f"无法将设备重命名为 '{new_name}' (可能名称冲突或为空)。")
            elif col in [COL_MPO, COL_LC, COL_SFP]:
                port_attr_map = {COL_MPO: 'mpo', COL_LC: 'lc', COL_SFP: 'sfp'}; port_name_map = {COL_MPO: 'MPO', COL_LC: 'LC', COL_SFP: 'SFP+'}
                attr_suffix = port_attr_map.get(col); port_type_name = port_name_map.get(col)
                device = self.network_manager.get_device_by_id(dev_id);
                if not device: raise ValueError("找不到设备")
                is_uhd_horizon = device.type in UHD_TYPES; is_micron = device.type == DEV_MN
                can_edit_this_port = (attr_suffix in ['mpo', 'lc'] and is_uhd_horizon) or (attr_suffix == 'sfp' and is_micron)
                if not can_edit_this_port:
                     old_count = getattr(device, f"{attr_suffix}_total", 0); print(f"不允许修改设备类型 '{device.type}' 的 '{port_type_name}' 端口数量。"); item.setText(str(old_count))
                else:
                    old_count = getattr(device, f"{attr_suffix}_total", 0)
                    try:
                        new_count = int(item.text().strip()); assert new_count >= 0
                        if new_count != old_count:
                            user_confirmed = True
                            if not self.suppress_confirmations:
                                reply = QMessageBox.question(self, "确认修改端口数量", f"修改设备 '{device.name}' 的 {port_type_name} 端口数量将清除所有现有连接。\n是否继续？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
                                user_confirmed = (reply == QMessageBox.StandardButton.Yes)
                            if user_confirmed:
                                update_kwargs = {f"new_{attr_suffix}": new_count}
                                if self.network_manager.update_device(dev_id, **update_kwargs):
                                    success = True; self.clear_results(); self._update_port_totals_display(); self._update_manual_port_options()
                                else: item.setText(str(old_count)); QMessageBox.warning(self, "更新失败", f"更新 {port_type_name} 端口数量失败。")
                            else: print("用户取消修改端口数量。"); item.setText(str(old_count))
                        else: success = True
                    except (ValueError, AssertionError): QMessageBox.warning(self, "输入错误", f"{port_type_name} 端口数量必须是非负整数。"); item.setText(str(old_count))
        finally: self.device_tablewidget.blockSignals(False) # !! 修改: 使用 self.xxx !!

    @Slot(QTableWidgetItem)
    def show_device_details_from_table(self, item: QTableWidgetItem):
        """处理设备表格双击事件，显示设备详情。"""
        # !! 修改: 使用 self.xxx !!
        if not item: return
        row = item.row(); name_item = self.device_tablewidget.item(row, COL_NAME)
        if not name_item: return
        dev_id = name_item.data(Qt.ItemDataRole.UserRole)
        device = self.network_manager.get_device_by_id(dev_id)
        if not device: QMessageBox.critical(self, "错误", "无法找到所选设备的详细信息。"); return
        self._display_device_details_popup(device)

    @Slot()
    def on_layout_change(self):
        """处理布局下拉框选择变化事件。"""
        if self.network_manager.get_all_devices():
            print("布局选择已更改，重置节点位置并重绘。")
            self.topology_controller.reset_layout_state()

    @Slot(str)
    def filter_device_table(self, text: str):
        """根据输入过滤设备表格。"""
        # !! 修改: 使用 self.xxx !!
        filter_text = text.lower()
        for row in range(self.device_tablewidget.rowCount()):
            match = False; name_item = self.device_tablewidget.item(row, COL_NAME); type_item = self.device_tablewidget.item(row, COL_TYPE)
            if name_item and filter_text in name_item.text().lower(): match = True
            if not match and type_item and filter_text in type_item.text().lower(): match = True
            self.device_tablewidget.setRowHidden(row, not match)

    @Slot()
    def filter_connection_list(self):
        """根据下拉框和输入框过滤手动编辑中的连接列表。"""
        # !! 修改: 使用 self.xxx !!
        selected_type = self.conn_filter_type_combo.currentText(); filter_device_text = self.conn_filter_device_entry.text().strip().lower()
        type_filter_active = selected_type != "所有类型"
        for i in range(self.manual_connection_list.count()):
            item = self.manual_connection_list.item(i); conn_data = item.data(Qt.ItemDataRole.UserRole)
            if not conn_data or not isinstance(conn_data, tuple) or len(conn_data) != 5: item.setHidden(False); continue
            dev1, _, dev2, _, conn_type = conn_data; item_conn_type = conn_type
            dev1_name_lower = dev1.name.lower(); dev2_name_lower = dev2.name.lower()
            type_match = (not type_filter_active) or (item_conn_type == selected_type)
            device_match = (not filter_device_text) or (filter_device_text in dev1_name_lower or filter_device_text in dev2_name_lower)
            item.setHidden(not (type_match and device_match))

    # --- UI 更新辅助方法 ---

    def _add_device_to_table(self, device: Device):
        """将设备对象添加到 UI 表格中。"""
        # !! 修改: 使用 self.xxx !!
        row_position = self.device_tablewidget.rowCount(); self.device_tablewidget.insertRow(row_position)
        editable_flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable; non_editable_flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled
        name_item = QTableWidgetItem(device.name); name_item.setData(Qt.ItemDataRole.UserRole, device.id); name_item.setFlags(editable_flags)
        type_item = QTableWidgetItem(device.type); type_item.setFlags(non_editable_flags)
        mpo_item = NumericTableWidgetItem(str(device.mpo_total)); mpo_item.setData(Qt.ItemDataRole.UserRole + 1, device.mpo_total); mpo_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter); mpo_item.setFlags(editable_flags if device.type in UHD_TYPES else non_editable_flags)
        lc_item = NumericTableWidgetItem(str(device.lc_total)); lc_item.setData(Qt.ItemDataRole.UserRole + 1, device.lc_total); lc_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter); lc_item.setFlags(editable_flags if device.type in UHD_TYPES else non_editable_flags)
        sfp_item = NumericTableWidgetItem(str(device.sfp_total)); sfp_item.setData(Qt.ItemDataRole.UserRole + 1, device.sfp_total); sfp_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter); sfp_item.setFlags(editable_flags if device.type == DEV_MN else non_editable_flags)
        conn_val = float(f"{device.connections:.2f}"); conn_item = NumericTableWidgetItem(f"{conn_val:.2f}"); conn_item.setData(Qt.ItemDataRole.UserRole + 1, conn_val); conn_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter); conn_item.setFlags(non_editable_flags)
        self.device_tablewidget.setItem(row_position, COL_NAME, name_item); self.device_tablewidget.setItem(row_position, COL_TYPE, type_item); self.device_tablewidget.setItem(row_position, COL_MPO, mpo_item); self.device_tablewidget.setItem(row_position, COL_LC, lc_item); self.device_tablewidget.setItem(row_position, COL_SFP, sfp_item); self.device_tablewidget.setItem(row_position, COL_CONN, conn_item)

    def _update_device_table_connections(self):
        """更新设备表格中的“连接数”列。"""
        # !! 修改: 使用 self.xxx !!
        for row in range(self.device_tablewidget.rowCount()):
            name_item = self.device_tablewidget.item(row, COL_NAME)
            if name_item:
                dev_id = name_item.data(Qt.ItemDataRole.UserRole); device = self.network_manager.get_device_by_id(dev_id)
                if device:
                    conn_val = float(f"{device.connections:.2f}"); conn_item = self.device_tablewidget.item(row, COL_CONN)
                    if conn_item: conn_item.setText(f"{conn_val:.2f}"); conn_item.setData(Qt.ItemDataRole.UserRole + 1, conn_val)
                    else: conn_item = NumericTableWidgetItem(f"{conn_val:.2f}"); conn_item.setData(Qt.ItemDataRole.UserRole + 1, conn_val); conn_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter); conn_item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled); self.device_tablewidget.setItem(row, COL_CONN, conn_item)

    def _update_device_combos(self):
        """更新手动编辑区域的设备下拉框选项。"""
        # !! 修改: 使用 self.xxx !!
        self.edit_dev1_combo.blockSignals(True); self.edit_dev2_combo.blockSignals(True)
        current_dev1_id = self.edit_dev1_combo.currentData(); current_dev2_id = self.edit_dev2_combo.currentData()
        self.edit_dev1_combo.clear(); self.edit_dev2_combo.clear(); self.edit_dev1_combo.addItem("选择设备 1...", userData=None); self.edit_dev2_combo.addItem("选择设备 2...", userData=None)
        sorted_devices = sorted(self.network_manager.get_all_devices(), key=lambda dev: dev.name)
        idx1_to_select = 0; idx2_to_select = 0
        for i, dev in enumerate(sorted_devices):
            item_text = f"{dev.name} ({dev.type})"; self.edit_dev1_combo.addItem(item_text, userData=dev.id); self.edit_dev2_combo.addItem(item_text, userData=dev.id)
            if dev.id == current_dev1_id: idx1_to_select = i + 1
            if dev.id == current_dev2_id: idx2_to_select = i + 1
        self.edit_dev1_combo.setCurrentIndex(idx1_to_select); self.edit_dev2_combo.setCurrentIndex(idx2_to_select)
        self.edit_dev1_combo.blockSignals(False); self.edit_dev2_combo.blockSignals(False)
        self._update_manual_port_options()

    def _populate_edit_port_combos(self, device_combo_to_populate: QComboBox, port_combo_to_populate: QComboBox, other_device_combo: QComboBox, other_port_combo: QComboBox):
        """动态填充指定的端口下拉列表，并根据另一侧的选择进行过滤。"""
        # (此方法逻辑不变)
        port_combo_to_populate.blockSignals(True); current_port_selection = port_combo_to_populate.currentText(); port_combo_to_populate.clear(); port_combo_to_populate.addItem("选择端口..."); port_combo_to_populate.setEnabled(False)
        dev_id = device_combo_to_populate.currentData()
        if dev_id is not None:
            device = self.network_manager.get_device_by_id(dev_id)
            if device:
                available_ports = self.network_manager.get_available_ports(dev_id); ports_to_add = []
                other_dev_id = other_device_combo.currentData(); other_port_name = other_port_combo.currentText()
                if other_dev_id is not None and other_port_name != "选择端口...":
                    compatible_types_here = self.network_manager.get_compatible_port_types(other_dev_id, other_port_name)
                    for port in available_ports:
                        port_type_here = get_port_type_from_name(port);
                        if port_type_here in compatible_types_here: ports_to_add.append(port)
                else: ports_to_add = available_ports
                if ports_to_add:
                    port_combo_to_populate.addItems(ports_to_add); index_to_select = port_combo_to_populate.findText(current_port_selection); port_combo_to_populate.setCurrentIndex(index_to_select if index_to_select != -1 else 0); port_combo_to_populate.setEnabled(True)
                else: port_combo_to_populate.addItem("无兼容/可用端口"); port_combo_to_populate.setCurrentIndex(1)
        port_combo_to_populate.blockSignals(False)

    @Slot()
    def _update_manual_port_options(self):
        """统一更新手动添加连接中的两个端口下拉列表的选项。"""
        # !! 修改: 使用 self.xxx !!
        self._populate_edit_port_combos(self.edit_dev1_combo, self.edit_port1_combo, self.edit_dev2_combo, self.edit_port2_combo)
        self._populate_edit_port_combos(self.edit_dev2_combo, self.edit_port2_combo, self.edit_dev1_combo, self.edit_port1_combo)

    def _update_port_totals_display(self):
        """更新显示端口总数的标签。"""
        # !! 修改: 使用 self.xxx !!
        totals = self.network_manager.calculate_port_totals(); self.port_totals_label.setText(f"总计: {PORT_MPO}: {totals['mpo']}, {PORT_LC}: {totals['lc']}, {PORT_SFP}+: {totals['sfp']}")

    def _update_connection_views(self):
        """更新连接列表文本框、手动编辑列表和拓扑图。"""
        print("DEBUG: _update_connection_views called")
        # !! 修改: 使用 self.xxx !!
        # 1. 更新连接列表文本框 (QTextEdit)
        self.connections_textedit.clear(); connections = self.network_manager.get_all_connections()
        if connections:
            self.connections_textedit.append("<b>连接列表:</b><hr>")
            for i, conn in enumerate(connections): dev1, port1, dev2, port2, conn_type = conn; self.connections_textedit.append(f"{i+1}. {dev1.name} [{port1}] &lt;-&gt; {dev2.name} [{port2}] ({conn_type})")
        else: self.connections_textedit.append("无连接。")
        # 2. 更新手动编辑中的连接列表 (QListWidget)
        self.manual_connection_list.clear()
        if connections:
            for i, conn in enumerate(connections): dev1, port1, dev2, port2, conn_type = conn; item_text = f"{i+1}. {dev1.name} [{port1}] <-> {dev2.name} [{port2}] ({conn_type})"; item = QListWidgetItem(item_text); item.setData(Qt.ItemDataRole.UserRole, conn); self.manual_connection_list.addItem(item)
        self.remove_manual_button.setEnabled(bool(connections)); self.filter_connection_list()
        # 3. 更新拓扑图
        selected_layout = self.layout_combo.currentText().lower()
        devices_for_plot = self.network_manager.get_all_devices(); connections_for_plot = self.network_manager.get_all_connections(); port_totals = self.network_manager.calculate_port_totals()
        current_node_positions = self.topology_controller.get_node_positions(); current_selected_node_id = self.topology_controller.get_selected_node_id()
        print(f"DEBUG: Plotting with positions: {current_node_positions}"); print(f"DEBUG: Plotting with selected node: {current_selected_node_id}")
        # !! 修改: 不再需要 MainWindow 持有 fig 引用 !!
        figure, calculated_pos = self.mpl_canvas.plot_topology(devices_for_plot, connections_for_plot, layout_algorithm=selected_layout, fixed_pos=current_node_positions, selected_node_id=current_selected_node_id, port_totals_dict=port_totals)
        # if figure: self.fig = figure # 移除
        if calculated_pos is not None and self.topology_controller.dragged_node_id is None:
            if self.topology_controller.node_positions is None or selected_layout != getattr(self, '_last_layout_used', None):
                print(f"DEBUG: Updating controller positions due to new layout '{selected_layout}'")
                self.topology_controller.node_positions = calculated_pos
                setattr(self, '_last_layout_used', selected_layout)
        # 4. 更新导出按钮状态
        has_connections = bool(connections); has_devices = bool(devices_for_plot); has_figure = figure is not None and has_devices
        self.export_list_button.setEnabled(has_connections); self.export_topo_button.setEnabled(has_figure); self.export_report_button.setEnabled(has_connections and has_figure)

    @Slot(object)
    def _display_device_details_popup(self, dev: Device):
        """显示包含设备详细信息的弹出窗口。(作为槽函数被 Controller 调用)"""
        # (此方法逻辑不变)
        details = f"ID: {dev.id}\n名称: {dev.name}\n类型: {dev.type}\n"
        avail_ports = dev.get_all_available_ports(); avail_lc_count = sum(1 for p in avail_ports if p.startswith(PORT_LC)); avail_sfp_count = sum(1 for p in avail_ports if p.startswith(PORT_SFP)); avail_mpo_ch_count = sum(1 for p in avail_ports if p.startswith(PORT_MPO))
        if dev.type in UHD_TYPES: details += f"{PORT_MPO} 端口总数: {dev.mpo_total}\n{PORT_LC} 端口总数: {dev.lc_total}\n可用 {PORT_MPO} 子通道: {avail_mpo_ch_count}\n可用 {PORT_LC} 端口: {avail_lc_count}\n"
        elif dev.type == DEV_MN: details += f"{PORT_SFP}+ 端口总数: {dev.sfp_total}\n可用 {PORT_SFP}+ 端口: {avail_sfp_count}\n"
        details += f"当前连接数 (估算): {dev.connections:.2f}\n"
        if dev.port_connections:
            details += "\n端口连接详情:\n"; lc_conns = {p: t for p, t in dev.port_connections.items() if p.startswith(PORT_LC)}; sfp_conns = {p: t for p, t in dev.port_connections.items() if p.startswith(PORT_SFP)}; mpo_conns_grouped = defaultdict(dict)
            for p, t in dev.port_connections.items():
                if p.startswith(PORT_MPO): mpo_base = p.split('-')[0]; mpo_conns_grouped[mpo_base][p] = t
            if lc_conns: details += f"  {PORT_LC} 连接:\n";
            for port in sorted(lc_conns.keys(), key=lambda x: int(x[len(PORT_LC):])): details += f"    {port} -> {lc_conns[port]}\n"
            if sfp_conns: details += f"  {PORT_SFP}+ 连接:\n";
            for port in sorted(sfp_conns.keys(), key=lambda x: int(x[len(PORT_SFP):])): details += f"    {port} -> {sfp_conns[port]}\n"
            if mpo_conns_grouped: details += f"  {PORT_MPO} 连接 (Breakout):\n";
            for base_port in sorted(mpo_conns_grouped.keys(), key=lambda x: int(x[len(PORT_MPO):])):
                details += f"    {base_port}:\n";
                for port in sorted(mpo_conns_grouped[base_port].keys(), key=lambda x: int(x.split('-Ch')[-1])): details += f"      {port} -> {mpo_conns_grouped[base_port][port]}\n"
        QMessageBox.information(self, f"设备详情 - {dev.name}", details)

    def _get_matplotlib_font_prop(self):
         """获取用于 Matplotlib 的 FontProperties 对象"""
         # (此方法逻辑不变)
         try:
             if hasattr(self, 'chinese_font') and self.chinese_font.family():
                 if self.chinese_font.family() in [f.name for f in font_manager.fontManager.ttflist]: return font_manager.FontProperties(family=self.chinese_font.family())
                 else:
                      db = QFontDatabase(); families = db.families()
                      try:
                           font_index = families.index(self.chinese_font.family()); font_paths = db.applicationFontFiles(font_index)
                           if font_paths:
                                font_manager.fontManager.addfont(font_paths[0])
                                if self.chinese_font.family() in [f.name for f in font_manager.fontManager.ttflist]: return font_manager.FontProperties(family=self.chinese_font.family())
                      except ValueError: print(f"警告: 字体家族 '{self.chinese_font.family()}' 未在 QFontDatabase 中找到。")
                      print(f"警告: Qt 字体 '{self.chinese_font.family()}' 未在 Matplotlib 字体列表中找到。")
         except Exception as e: print(f"获取 Matplotlib 字体属性时出错: {e}")
         return font_manager.FontProperties(family='sans-serif')

