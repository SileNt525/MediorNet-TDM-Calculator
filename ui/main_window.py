# -*- coding: utf-8 -*-
"""
ui/main_window.py

定义主应用程序窗口 MainWindow 类。
包含 UI 布局、控件、事件处理 (槽函数) 和与 NetworkManager 的交互。
"""

import sys
import os
import copy
import json
import base64
import io
import datetime
import random # 用于 _display_device_details_popup 中的兼容性
from collections import defaultdict
from typing import Optional, List, Dict, Tuple, Set, Any # <--- **修复: 添加了 typing 导入**

# PySide6 imports
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QTextEdit,
    QTabWidget, QFrame, QFileDialog, QMessageBox, QSpacerItem, QSizePolicy,
    QGridLayout, QListWidgetItem, QAbstractItemView, QTableWidget, QTableWidgetItem,
    QHeaderView, QListWidget, QSplitter, QCheckBox
)
from PySide6.QtCore import Slot, Qt
from PySide6.QtGui import QFont, QGuiApplication, QFontDatabase # 导入 QFontDatabase

# Matplotlib imports (for canvas interaction logic)
from matplotlib import font_manager
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D # 导入 Line2D 用于图例

# NetworkX import (needed for graph interaction logic if kept here)
import networkx as nx

# --- 从项目模块导入 ---
# 注意：这里的导入路径假设 main.py 与 core, ui, utils 目录在同一级
try:
    from core.network_manager import NetworkManager, ConnectionType
    from core.device import (
        Device,
        DEV_UHD, DEV_HORIZON, DEV_MN, UHD_TYPES,
        PORT_MPO, PORT_LC, PORT_SFP, PORT_UNKNOWN,
        get_port_type_from_name
    )
    # 假设这些文件已存在于相应的包中
    from .topology_canvas import MplCanvas
    from .widgets import NumericTableWidgetItem
    from utils.export_utils import export_connections_to_file, export_topology_to_file, export_report_to_html
    from utils.misc_utils import resource_path
except ImportError as e:
    print(f"导入错误 (main_window.py): {e} - 请确保 core, ui, utils 包及其模块已正确创建。")
    # 提供一些回退，以便代码至少可以被解析，尽管运行时会失败
    NetworkManager = object
    Device = object
    ConnectionType = tuple
    DEV_UHD, DEV_HORIZON, DEV_MN, UHD_TYPES = '', '', '', []
    PORT_MPO, PORT_LC, PORT_SFP, PORT_UNKNOWN = '', '', '', ''
    get_port_type_from_name = lambda x: ''
    MplCanvas = QWidget # Fallback
    NumericTableWidgetItem = QTableWidgetItem # Fallback
    export_connections_to_file = lambda *args, **kwargs: None
    export_topology_to_file = lambda *args, **kwargs: None
    export_report_to_html = lambda *args, **kwargs: None
    resource_path = lambda x: x # Fallback


# --- UI 常量 (表格列索引) ---
COL_NAME = 0
COL_TYPE = 1
COL_MPO = 2
COL_LC = 3
COL_SFP = 4
COL_CONN = 5

# --- QSS 样式定义 (从 main.py 移入) ---
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
        self.setWindowTitle("MediorNet TDM 连接计算器 V1.1 (Refactored)")
        self.setGeometry(100, 100, 1100, 800)

        # --- 核心数据管理器 ---
        self.network_manager = NetworkManager()

        # --- UI 状态变量 ---
        self.fig = None # Matplotlib Figure 对象 (由 MplCanvas 管理)
        self.node_positions: Optional[Dict[int, Tuple[float, float]]] = None # 存储节点布局 {node_id: (x, y)}
        self.selected_node_id: Optional[int] = None # 当前选中的节点 ID
        self.dragged_node_id: Optional[int] = None # 当前正在拖动的节点 ID
        self.drag_offset: Tuple[float, float] = (0, 0)   # 拖动时的偏移量
        self.connecting_node_id: Optional[int] = None # Shift 拖动时起始节点的 ID
        self.connection_line: Optional[Line2D] = None # Shift 拖动时绘制的临时线对象
        self.suppress_confirmations: bool = False # 是否跳过确认弹窗

        # --- 字体加载 ---
        self.chinese_font = self._setup_fonts()

        # --- UI 布局与控件初始化 ---
        self._setup_ui()

        # --- 初始化 UI 状态 ---
        self.update_port_entries()
        self._update_port_totals_display()
        self._update_connection_views()
        self._update_device_combos()

    def _setup_fonts(self) -> QFont:
        """加载并设置字体，返回 QFont 对象。"""
        default_font = QFont()
        default_font.setPointSize(10)
        try:
            # 尝试加载 Noto Sans CJK SC 字体
            font_relative_path = os.path.join('assets', 'NotoSansCJKsc-Regular.otf')
            font_path = resource_path(font_relative_path) # 使用 utils 中的函数
            if os.path.exists(font_path):
                font_id = QFontDatabase.addApplicationFont(font_path)
                if font_id != -1:
                    families = QFontDatabase.applicationFontFamilies(font_id)
                    if families:
                        family = families[0]
                        print(f"成功加载并设置字体: {family} (路径: {font_path})")
                        # 设置 Matplotlib 字体
                        font_manager.fontManager.addfont(font_path)
                        plt.rcParams['font.sans-serif'] = [family] + plt.rcParams.get('font.sans-serif', [])
                        plt.rcParams['axes.unicode_minus'] = False
                        return QFont(family, 10)
                    else:
                        print(f"警告: 无法从字体文件获取 family name {font_path}")

                else:
                    print(f"警告: 添加字体失败 {font_path}")
            else:
                print(f"警告: 未在路径 {font_path} 找到字体文件。将使用默认字体。")

        except ImportError:
             print("警告: 无法导入 QFontDatabase。将使用默认字体。")
        except Exception as e:
            print(f"加载或设置字体时出错: {e}")

        # 回退到默认字体
        plt.rcParams['font.sans-serif'] = ['sans-serif']
        plt.rcParams['axes.unicode_minus'] = False
        return default_font


    def _setup_ui(self):
        """初始化和布局所有 UI 控件。"""
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        main_splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(main_splitter)

        # --- 左侧面板 ---
        left_panel = QFrame()
        left_panel.setFrameShape(QFrame.Shape.StyledPanel)
        left_layout = QVBoxLayout(left_panel)
        main_splitter.addWidget(left_panel)

        # 添加设备组
        add_group = QFrame()
        add_group.setObjectName("addDeviceGroup")
        add_group_layout = QGridLayout(add_group)
        add_group_layout.setContentsMargins(10, 15, 10, 10)
        add_group_layout.setVerticalSpacing(8)
        add_title = QLabel("<b>添加新设备</b>")
        add_title.setFont(QFont(self.chinese_font.family(), 11))
        add_group_layout.addWidget(add_title, 0, 0, 1, 2, alignment=Qt.AlignmentFlag.AlignCenter)

        add_group_layout.addWidget(QLabel("类型:"), 1, 0)
        self.device_type_combo = QComboBox()
        self.device_type_combo.addItems([DEV_UHD, DEV_HORIZON, DEV_MN])
        self.device_type_combo.setFont(self.chinese_font)
        self.device_type_combo.currentIndexChanged.connect(self.update_port_entries)
        add_group_layout.addWidget(self.device_type_combo, 1, 1)

        add_group_layout.addWidget(QLabel("名称:"), 2, 0)
        self.device_name_entry = QLineEdit()
        self.device_name_entry.setFont(self.chinese_font)
        add_group_layout.addWidget(self.device_name_entry, 2, 1)

        self.mpo_label = QLabel(f"{PORT_MPO} 端口:")
        add_group_layout.addWidget(self.mpo_label, 3, 0)
        self.mpo_entry = QLineEdit("4") # 默认值
        self.mpo_entry.setFont(self.chinese_font)
        add_group_layout.addWidget(self.mpo_entry, 3, 1)

        self.lc_label = QLabel(f"{PORT_LC} 端口:")
        add_group_layout.addWidget(self.lc_label, 4, 0)
        self.lc_entry = QLineEdit("2") # 默认值
        self.lc_entry.setFont(self.chinese_font)
        add_group_layout.addWidget(self.lc_entry, 4, 1)

        self.sfp_label = QLabel(f"{PORT_SFP}+ 端口:")
        self.sfp_entry = QLineEdit("8") # 默认值
        self.sfp_entry.setFont(self.chinese_font)
        add_group_layout.addWidget(self.sfp_label, 5, 0)
        add_group_layout.addWidget(self.sfp_entry, 5, 1)
        self.sfp_label.hide() # 初始隐藏 SFP
        self.sfp_entry.hide()

        self.add_button = QPushButton("添加设备")
        self.add_button.setFont(self.chinese_font)
        self.add_button.clicked.connect(self.add_device) # 连接槽函数
        add_group_layout.addWidget(self.add_button, 6, 0, 1, 2, alignment=Qt.AlignmentFlag.AlignCenter)
        left_layout.addWidget(add_group)

        # 设备列表组
        list_group = QFrame()
        list_group.setObjectName("listGroup")
        list_group_layout = QVBoxLayout(list_group)
        list_group_layout.setContentsMargins(10, 15, 10, 10)

        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("过滤:", font=self.chinese_font))
        self.device_filter_entry = QLineEdit()
        self.device_filter_entry.setFont(self.chinese_font)
        self.device_filter_entry.setPlaceholderText("按名称或类型过滤...")
        self.device_filter_entry.textChanged.connect(self.filter_device_table) # 连接槽函数
        filter_layout.addWidget(self.device_filter_entry)
        list_group_layout.addLayout(filter_layout)

        self.device_tablewidget = QTableWidget()
        self.device_tablewidget.setFont(self.chinese_font)
        self.device_tablewidget.setColumnCount(6)
        self.device_tablewidget.setHorizontalHeaderLabels(["名称", "类型", PORT_MPO, PORT_LC, f"{PORT_SFP}+", "连接数(估)"])
        self.device_tablewidget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.device_tablewidget.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.device_tablewidget.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.SelectedClicked | QAbstractItemView.EditTrigger.EditKeyPressed)
        self.device_tablewidget.setSortingEnabled(True)
        header = self.device_tablewidget.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(COL_NAME, QHeaderView.ResizeMode.Stretch) # 名称列拉伸
        self.device_tablewidget.setColumnWidth(COL_TYPE, 90)
        self.device_tablewidget.setColumnWidth(COL_MPO, 50)
        self.device_tablewidget.setColumnWidth(COL_LC, 50)
        self.device_tablewidget.setColumnWidth(COL_SFP, 50)
        self.device_tablewidget.setColumnWidth(COL_CONN, 80)
        self.device_tablewidget.itemDoubleClicked.connect(self.show_device_details_from_table) # 连接槽函数
        self.device_tablewidget.itemChanged.connect(self.on_device_item_changed) # 连接槽函数
        list_group_layout.addWidget(self.device_tablewidget)

        device_op_layout = QHBoxLayout()
        self.remove_button = QPushButton("移除选中")
        self.remove_button.setFont(self.chinese_font)
        self.remove_button.clicked.connect(self.remove_device) # 连接槽函数
        self.clear_button = QPushButton("清空所有")
        self.clear_button.setFont(self.chinese_font)
        self.clear_button.clicked.connect(self.clear_all_devices) # 连接槽函数
        device_op_layout.addWidget(self.remove_button)
        device_op_layout.addWidget(self.clear_button)
        list_group_layout.addLayout(device_op_layout)

        self.port_totals_label = QLabel("总计: MPO: 0, LC: 0, SFP+: 0")
        font = self.port_totals_label.font()
        font.setBold(True)
        self.port_totals_label.setFont(font)
        self.port_totals_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.port_totals_label.setStyleSheet("padding-top: 5px; padding-right: 5px;")
        list_group_layout.addWidget(self.port_totals_label)
        left_layout.addWidget(list_group)

        # 文件操作组
        file_group = QFrame()
        file_group.setObjectName("fileGroup")
        file_group_layout = QGridLayout(file_group)
        file_group_layout.setContentsMargins(10, 15, 10, 10)
        file_title = QLabel("<b>文件操作</b>")
        file_title.setFont(QFont(self.chinese_font.family(), 11))
        file_group_layout.addWidget(file_title, 0, 0, 1, 2, alignment=Qt.AlignmentFlag.AlignCenter)

        self.save_button = QPushButton("保存配置")
        self.save_button.setFont(self.chinese_font)
        self.save_button.clicked.connect(self.save_config) # 连接槽函数
        self.load_button = QPushButton("加载配置")
        self.load_button.setFont(self.chinese_font)
        self.load_button.clicked.connect(self.load_config) # 连接槽函数
        file_group_layout.addWidget(self.save_button, 1, 0)
        file_group_layout.addWidget(self.load_button, 1, 1)

        self.export_list_button = QPushButton("导出列表")
        self.export_list_button.setFont(self.chinese_font)
        self.export_list_button.clicked.connect(self.export_connections) # 连接槽函数
        self.export_list_button.setEnabled(False)
        file_group_layout.addWidget(self.export_list_button, 2, 0)

        self.export_topo_button = QPushButton("导出拓扑图")
        self.export_topo_button.setFont(self.chinese_font)
        self.export_topo_button.clicked.connect(self.export_topology) # 连接槽函数
        self.export_topo_button.setEnabled(False)
        file_group_layout.addWidget(self.export_topo_button, 2, 1)

        self.export_report_button = QPushButton("导出报告 (HTML)")
        self.export_report_button.setFont(self.chinese_font)
        self.export_report_button.clicked.connect(self.export_html_report) # 连接槽函数
        self.export_report_button.setEnabled(False)
        file_group_layout.addWidget(self.export_report_button, 3, 0, 1, 2)
        left_layout.addWidget(file_group)

        # 跳过确认弹窗设置
        suppress_frame = QFrame()
        suppress_frame.setFrameShape(QFrame.Shape.NoFrame)
        suppress_layout = QHBoxLayout(suppress_frame)
        suppress_layout.setContentsMargins(10, 0, 10, 5)
        self.suppress_confirm_checkbox = QCheckBox("跳过确认弹窗")
        self.suppress_confirm_checkbox.setFont(self.chinese_font)
        self.suppress_confirm_checkbox.stateChanged.connect(self._toggle_suppress_confirmations) # 连接槽函数
        suppress_layout.addWidget(self.suppress_confirm_checkbox)
        suppress_layout.addStretch()
        left_layout.addWidget(suppress_frame)

        left_layout.addStretch()

        # --- 右侧面板与 Tab 页 ---
        right_panel = QFrame()
        right_panel.setFrameShape(QFrame.Shape.StyledPanel)
        right_layout = QVBoxLayout(right_panel)
        main_splitter.addWidget(right_panel)

        # 计算控制栏
        calculate_control_frame = QFrame()
        calculate_control_frame.setObjectName("calculateControlFrame")
        calculate_control_layout = QHBoxLayout(calculate_control_frame)
        calculate_control_layout.setContentsMargins(10, 5, 10, 5)

        calculate_control_layout.addWidget(QLabel("计算模式:", font=self.chinese_font))
        self.topology_mode_combo = QComboBox()
        self.topology_mode_combo.setFont(self.chinese_font)
        self.topology_mode_combo.addItems(["Mesh", "环形"])
        calculate_control_layout.addWidget(self.topology_mode_combo)

        calculate_control_layout.addWidget(QLabel("布局:", font=self.chinese_font))
        self.layout_combo = QComboBox()
        self.layout_combo.setFont(self.chinese_font)
        self.layout_combo.addItems(["Spring", "Circular", "Kamada-Kawai", "Random", "Shell"])
        self.layout_combo.currentIndexChanged.connect(self.on_layout_change) # 连接槽函数
        calculate_control_layout.addWidget(self.layout_combo)

        self.calculate_button = QPushButton("计算连接")
        self.calculate_button.setFont(self.chinese_font)
        self.calculate_button.clicked.connect(self.calculate_and_display) # 连接槽函数
        calculate_control_layout.addWidget(self.calculate_button)

        self.fill_mesh_button = QPushButton("填充 (Mesh)")
        self.fill_mesh_button.setFont(self.chinese_font)
        self.fill_mesh_button.setEnabled(False)
        self.fill_mesh_button.clicked.connect(self.fill_remaining_mesh) # 连接槽函数
        calculate_control_layout.addWidget(self.fill_mesh_button)

        self.fill_ring_button = QPushButton("填充 (环形)")
        self.fill_ring_button.setFont(self.chinese_font)
        self.fill_ring_button.setEnabled(False)
        self.fill_ring_button.clicked.connect(self.fill_remaining_ring) # 连接槽函数
        calculate_control_layout.addWidget(self.fill_ring_button)

        calculate_control_layout.addStretch()
        right_layout.addWidget(calculate_control_frame)

        # Tab 控件
        self.tab_widget = QTabWidget()
        self.tab_widget.setFont(self.chinese_font)
        right_layout.addWidget(self.tab_widget)

        # 连接列表 Tab
        self.connections_tab = QWidget()
        connections_layout = QVBoxLayout(self.connections_tab)
        self.connections_textedit = QTextEdit()
        self.connections_textedit.setFont(self.chinese_font)
        self.connections_textedit.setReadOnly(True)
        connections_layout.addWidget(self.connections_textedit)
        self.tab_widget.addTab(self.connections_tab, "连接列表")

        # 拓扑图 Tab
        self.topology_tab = QWidget()
        topology_layout = QVBoxLayout(self.topology_tab)
        # 实例化 MplCanvas
        self.mpl_canvas = MplCanvas(self.topology_tab, width=8, height=6, dpi=100)
        topology_layout.addWidget(self.mpl_canvas)
        self.tab_widget.addTab(self.topology_tab, "拓扑图")
        # 连接画布事件
        self.mpl_canvas.mpl_connect('button_press_event', self.on_canvas_press)
        self.mpl_canvas.mpl_connect('motion_notify_event', self.on_canvas_motion)
        self.mpl_canvas.mpl_connect('button_release_event', self.on_canvas_release)

        # 手动编辑 Tab
        self.edit_tab = QWidget()
        edit_main_layout = QVBoxLayout(self.edit_tab)

        # 添加手动连接组
        add_manual_group = QFrame()
        add_manual_group.setObjectName("addManualGroup")
        add_manual_group.setFrameShape(QFrame.Shape.StyledPanel)
        add_manual_layout = QGridLayout(add_manual_group)
        add_manual_layout.setContentsMargins(10,10,10,10)
        add_manual_layout.addWidget(QLabel("<b>添加手动连接</b>", font=self.chinese_font), 0, 0, 1, 4, alignment=Qt.AlignmentFlag.AlignCenter)

        add_manual_layout.addWidget(QLabel("设备 1:", font=self.chinese_font), 1, 0)
        self.edit_dev1_combo = QComboBox()
        self.edit_dev1_combo.setFont(self.chinese_font)
        add_manual_layout.addWidget(self.edit_dev1_combo, 1, 1)
        add_manual_layout.addWidget(QLabel("端口 1:", font=self.chinese_font), 1, 2)
        self.edit_port1_combo = QComboBox()
        self.edit_port1_combo.setFont(self.chinese_font)
        add_manual_layout.addWidget(self.edit_port1_combo, 1, 3)

        add_manual_layout.addWidget(QLabel("设备 2:", font=self.chinese_font), 2, 0)
        self.edit_dev2_combo = QComboBox()
        self.edit_dev2_combo.setFont(self.chinese_font)
        add_manual_layout.addWidget(self.edit_dev2_combo, 2, 1)
        add_manual_layout.addWidget(QLabel("端口 2:", font=self.chinese_font), 2, 2)
        self.edit_port2_combo = QComboBox()
        self.edit_port2_combo.setFont(self.chinese_font)
        add_manual_layout.addWidget(self.edit_port2_combo, 2, 3)

        self.add_manual_button = QPushButton("添加连接")
        self.add_manual_button.setFont(self.chinese_font)
        self.add_manual_button.clicked.connect(self.add_manual_connection) # 连接槽函数
        add_manual_layout.addWidget(self.add_manual_button, 3, 0, 1, 4, alignment=Qt.AlignmentFlag.AlignCenter)
        # 连接设备下拉框信号以更新端口选项
        self.edit_dev1_combo.currentIndexChanged.connect(self._update_manual_port_options)
        self.edit_port1_combo.currentIndexChanged.connect(self._update_manual_port_options) # 端口变化也可能影响另一侧
        self.edit_dev2_combo.currentIndexChanged.connect(self._update_manual_port_options)
        self.edit_port2_combo.currentIndexChanged.connect(self._update_manual_port_options) # 端口变化也可能影响另一侧
        edit_main_layout.addWidget(add_manual_group)

        # 移除现有连接组
        remove_manual_group = QFrame()
        remove_manual_group.setObjectName("removeManualGroup")
        remove_manual_group.setFrameShape(QFrame.Shape.StyledPanel)
        remove_manual_layout = QVBoxLayout(remove_manual_group)
        remove_manual_layout.setContentsMargins(10,10,10,10)
        remove_title = QLabel("<b>移除现有连接</b> (选中下方列表中的连接进行移除)")
        remove_title.setFont(self.chinese_font)
        remove_manual_layout.addWidget(remove_title)

        filter_conn_layout = QHBoxLayout()
        filter_conn_layout.addWidget(QLabel("类型过滤:", font=self.chinese_font))
        self.conn_filter_type_combo = QComboBox()
        self.conn_filter_type_combo.setFont(self.chinese_font)
        self.conn_filter_type_combo.addItems(["所有类型", "LC-LC (100G)", "MPO-MPO (25G)", "SFP-SFP (10G)", "MPO-SFP (10G)"])
        self.conn_filter_type_combo.currentIndexChanged.connect(self.filter_connection_list) # 连接槽函数
        filter_conn_layout.addWidget(self.conn_filter_type_combo)
        filter_conn_layout.addSpacing(15)
        filter_conn_layout.addWidget(QLabel("设备过滤:", font=self.chinese_font))
        self.conn_filter_device_entry = QLineEdit()
        self.conn_filter_device_entry.setFont(self.chinese_font)
        self.conn_filter_device_entry.setPlaceholderText("按设备名称过滤...")
        self.conn_filter_device_entry.textChanged.connect(self.filter_connection_list) # 连接槽函数
        filter_conn_layout.addWidget(self.conn_filter_device_entry)
        remove_manual_layout.insertLayout(1, filter_conn_layout) # 插入到标题下方

        self.manual_connection_list = QListWidget()
        self.manual_connection_list.setFont(self.chinese_font)
        self.manual_connection_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection) # 允许多选
        remove_manual_layout.addWidget(self.manual_connection_list)

        self.remove_manual_button = QPushButton("移除选中连接")
        self.remove_manual_button.setFont(self.chinese_font)
        self.remove_manual_button.clicked.connect(self.remove_manual_connection) # 连接槽函数
        self.remove_manual_button.setEnabled(False) # 初始禁用
        remove_manual_layout.addWidget(self.remove_manual_button, alignment=Qt.AlignmentFlag.AlignCenter)
        edit_main_layout.addWidget(remove_manual_group)

        self.tab_widget.addTab(self.edit_tab, "手动编辑")

        # 设置 Splitter 初始大小
        main_splitter.setSizes([400, 700])
        main_splitter.setStretchFactor(1, 1) # 让右侧面板随窗口拉伸

    # --- 槽函数 (Slot Methods) ---

    @Slot()
    def update_port_entries(self):
        """根据设备类型选择更新端口输入框的可见性。"""
        selected_type = self.device_type_combo.currentText()
        is_micron = selected_type == DEV_MN
        is_uhd_horizon = selected_type in UHD_TYPES

        self.mpo_label.setVisible(is_uhd_horizon)
        self.mpo_entry.setVisible(is_uhd_horizon)
        self.lc_label.setVisible(is_uhd_horizon)
        self.lc_entry.setVisible(is_uhd_horizon)
        self.sfp_label.setVisible(is_micron)
        self.sfp_entry.setVisible(is_micron)

    @Slot()
    def add_device(self):
        """处理“添加设备”按钮点击事件。"""
        dtype = self.device_type_combo.currentText()
        name = self.device_name_entry.text().strip()

        if not dtype:
            QMessageBox.critical(self, "错误", "请选择设备类型。")
            return
        if not name:
            QMessageBox.critical(self, "错误", "请输入设备名称。")
            return

        # 从 NetworkManager 添加设备
        mpo_ports_str = self.mpo_entry.text() or "0"
        lc_ports_str = self.lc_entry.text() or "0"
        sfp_ports_str = self.sfp_entry.text() or "0"
        mpo_ports, lc_ports, sfp_ports = 0, 0, 0

        try:
            if dtype in UHD_TYPES:
                mpo_ports = int(mpo_ports_str)
                lc_ports = int(lc_ports_str)
                if mpo_ports < 0 or lc_ports < 0: raise ValueError("端口数不能为负")
            elif dtype == DEV_MN:
                sfp_ports = int(sfp_ports_str)
                if sfp_ports < 0: raise ValueError("端口数不能为负")
            else:
                 raise ValueError("无效的设备类型") # 理论上不会发生

            new_device = self.network_manager.add_device(name, dtype, mpo_ports, lc_ports, sfp_ports)

            if new_device:
                self._add_device_to_table(new_device) # 更新 UI 表格
                self.device_name_entry.clear()
                self._update_device_combos() # 更新手动编辑下拉框
                self.clear_results() # 清除旧的计算结果和连接
                self.node_positions = None # 重置拓扑图布局
                self.selected_node_id = None
                self._update_port_totals_display() # 更新端口总数显示
                self._update_connection_views() # 更新视图
            else:
                # NetworkManager 内部会打印名称冲突错误
                QMessageBox.critical(self, "错误", f"无法添加设备 '{name}' (可能名称已存在)。")

        except ValueError as e:
            QMessageBox.critical(self, "输入错误", f"端口数量必须是非负整数: {e}")
        except Exception as e:
             QMessageBox.critical(self, "添加失败", f"添加设备时发生未知错误: {e}")


    @Slot()
    def remove_device(self):
        """处理“移除选中”按钮点击事件。"""
        selected_rows = sorted(list(set(index.row() for index in self.device_tablewidget.selectedIndexes())), reverse=True)
        if not selected_rows:
            QMessageBox.warning(self, "提示", "请先在表格中选择要移除的设备行。")
            return

        ids_to_remove = set()
        names_to_remove = []
        for row_index in selected_rows:
            name_item = self.device_tablewidget.item(row_index, COL_NAME)
            if name_item:
                dev_id = name_item.data(Qt.ItemDataRole.UserRole)
                ids_to_remove.add(dev_id)
                names_to_remove.append(name_item.text())

        if not ids_to_remove:
            return

        user_confirmed = True
        if not self.suppress_confirmations:
            reply = QMessageBox.question(self, "确认移除",
                                         f"确定要移除选中的设备及其所有连接吗？\n({', '.join(names_to_remove)})",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            user_confirmed = (reply == QMessageBox.StandardButton.Yes)

        if user_confirmed:
            removed_count = 0
            for dev_id in ids_to_remove:
                if self.network_manager.remove_device(dev_id):
                    removed_count += 1

            if removed_count > 0:
                # 从 UI 表格中移除行
                # 需要在循环外移除，或者从后往前移除以避免索引错误
                for row_index in selected_rows: # 之前已排序 reverse=True
                    self.device_tablewidget.removeRow(row_index)

                self._update_device_combos() # 更新下拉框
                # self.clear_results() # remove_device 内部已清除连接，这里只需更新视图
                self.node_positions = None # 重置布局
                self.selected_node_id = None
                self._update_port_totals_display()
                self._update_connection_views() # 重新绘制拓扑图和列表
                print(f"成功移除了 {removed_count} 个设备及其连接。")
            else:
                 print("没有设备被移除（可能在管理器中移除失败）。")
        else:
            print("用户取消移除设备。")


    @Slot()
    def clear_all_devices(self):
        """处理“清空所有”按钮点击事件。"""
        if not self.network_manager.get_all_devices():
            # 如果列表已空，无需操作
            return

        user_confirmed = True
        if not self.suppress_confirmations:
            reply = QMessageBox.question(self, "确认清空", "确定要清空所有设备和连接吗？",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            user_confirmed = (reply == QMessageBox.StandardButton.Yes)

        if user_confirmed:
            self.network_manager.clear_all_devices_and_connections()
            self.device_tablewidget.setRowCount(0) # 清空 UI 表格
            self._update_device_combos()
            self.node_positions = None
            self.selected_node_id = None
            self._update_port_totals_display()
            self._update_connection_views() # 清空显示
            # 禁用填充按钮
            self.fill_mesh_button.setEnabled(False)
            self.fill_ring_button.setEnabled(False)
            print("所有设备和连接已清空。")
        else:
            print("用户取消清空所有设备。")

    @Slot()
    def clear_results(self):
        """清除计算结果和连接，重置设备端口状态。"""
        self.network_manager.clear_connections() # 清除核心数据
        self.fig = None # 重置 Matplotlib Figure
        self.node_positions = None
        self.selected_node_id = None
        self.connections_textedit.clear()
        self.manual_connection_list.clear()
        self.mpl_canvas.axes.cla()
        self.mpl_canvas.axes.text(0.5, 0.5, '点击“计算”或手动添加连接', ha='center', va='center', fontproperties=self._get_matplotlib_font_prop())
        self.mpl_canvas.draw()
        # 禁用导出和填充按钮
        self.export_list_button.setEnabled(False)
        self.export_topo_button.setEnabled(False)
        self.export_report_button.setEnabled(False)
        self.remove_manual_button.setEnabled(False)
        self.fill_mesh_button.setEnabled(False)
        self.fill_ring_button.setEnabled(False)
        # 更新 UI 表格中的连接数列
        self._update_device_table_connections()
        # 更新端口总数
        self._update_port_totals_display()
        # 更新手动编辑下拉框
        self._update_manual_port_options()
        print("计算结果和连接已清除。")

    @Slot()
    def calculate_and_display(self):
        """处理“计算连接”按钮点击事件。"""
        devices = self.network_manager.get_all_devices()
        if not devices:
            QMessageBox.information(self, "提示", "请先添加设备。")
            return

        # 1. 清除 NetworkManager 中现有的连接和设备端口状态
        self.network_manager.clear_connections()

        # 2. 从 NetworkManager 获取计算结果 (纯数据，不修改状态)
        mode = self.topology_mode_combo.currentText()
        calculated_connections_data: List[ConnectionType] = []
        error_message = None
        if mode == "Mesh":
            print("使用 Mesh 算法进行计算...")
            calculated_connections_data = self.network_manager.calculate_mesh()
        elif mode == "环形":
            print("使用环形算法进行计算...")
            calculated_connections_data, error_message = self.network_manager.calculate_ring()
        else:
            QMessageBox.critical(self, "错误", f"未知的计算模式: {mode}")
            return

        if error_message:
            QMessageBox.warning(self, f"{mode} 计算警告", error_message)

        # 3. 将计算结果添加到 NetworkManager 的状态中
        added_count = 0
        if calculated_connections_data:
            print(f"计算得到 {len(calculated_connections_data)} 条连接，正在添加到管理器...")
            for conn_tuple_data in calculated_connections_data:
                # conn_tuple_data 包含的是原始 Device 对象引用
                dev1, port1, dev2, port2, _ = conn_tuple_data
                if self.network_manager.add_connection(dev1.id, port1, dev2.id, port2):
                    added_count += 1
                else:
                    # 理论上不应失败，因为是从干净状态开始计算的
                    print(f"警告: 将计算出的连接 {dev1.name}[{port1}]<->{dev2.name}[{port2}] 添加到管理器时失败。")
            print(f"成功添加了 {added_count} 条计算出的连接到管理器。")
        else:
            print("计算未产生任何连接。")


        # 4. 更新所有 UI 视图
        self.node_positions = None # 重置布局
        self.selected_node_id = None
        self._update_connection_views() # 从 manager 获取新状态并显示
        self._update_device_table_connections() # 更新表格连接数
        self._update_device_combos() # 更新手动编辑下拉框
        self._update_manual_port_options() # 更新手动编辑端口

        # 5. 更新按钮状态
        has_connections = bool(self.network_manager.get_all_connections())
        can_fill = has_connections or any(bool(dev.get_all_available_ports()) for dev in self.network_manager.get_all_devices())
        self.fill_mesh_button.setEnabled(can_fill)
        self.fill_ring_button.setEnabled(can_fill)
        self.export_list_button.setEnabled(has_connections)
        self.export_topo_button.setEnabled(bool(devices)) # 只要有设备就能导出空图
        self.export_report_button.setEnabled(has_connections and bool(devices))
        self.remove_manual_button.setEnabled(has_connections)


    @Slot()
    def fill_remaining_mesh(self):
        """处理“填充 (Mesh)”按钮点击事件。"""
        if not self.network_manager.get_all_devices():
            QMessageBox.information(self, "提示", "请先添加设备。")
            return

        print("开始填充剩余连接 (Mesh)...")
        new_connections = self.network_manager.fill_connections_mesh()

        if new_connections:
            # 更新 UI
            self.selected_node_id = None # 清除选择
            self._update_connection_views()
            self._update_device_table_connections()
            self._update_manual_port_options()
            self._update_port_totals_display()
            QMessageBox.information(self, "填充完成", f"成功添加了 {len(new_connections)} 条新 Mesh 连接。")
        else:
            QMessageBox.information(self, "填充完成", "没有找到更多可以建立的 Mesh 连接。")

        # 填充后禁用填充按钮，因为端口可能已用完
        self.fill_mesh_button.setEnabled(False)
        self.fill_ring_button.setEnabled(False)

    @Slot()
    def fill_remaining_ring(self):
        """处理“填充 (环形)”按钮点击事件。"""
        if not self.network_manager.get_all_devices():
            QMessageBox.information(self, "提示", "请先添加设备。")
            return

        print("开始填充剩余连接 (环形)...")
        new_connections = self.network_manager.fill_connections_ring()

        if new_connections:
             # 更新 UI
            self.selected_node_id = None # 清除选择
            self._update_connection_views()
            self._update_device_table_connections()
            self._update_manual_port_options()
            self._update_port_totals_display()
            QMessageBox.information(self, "填充完成", f"成功添加了 {len(new_connections)} 条新环形连接段。")
        else:
            QMessageBox.information(self, "填充完成", "没有找到更多可以建立的环形连接段。")

        # 填充后禁用填充按钮
        self.fill_mesh_button.setEnabled(False)
        self.fill_ring_button.setEnabled(False)

    @Slot()
    def save_config(self):
        """处理“保存配置”按钮点击事件。"""
        if not self.network_manager.get_all_devices():
            QMessageBox.warning(self, "提示", "设备列表为空，无需保存。")
            return

        filepath, _ = QFileDialog.getSaveFileName(self, "保存项目配置", "", "JSON 文件 (*.json);;所有文件 (*)")
        if not filepath:
            return

        if self.network_manager.save_project(filepath):
            QMessageBox.information(self, "成功", f"项目配置已保存到:\n{filepath}")
        else:
            QMessageBox.critical(self, "保存失败", f"无法保存项目配置文件。")

    @Slot()
    def load_config(self):
        """处理“加载配置”按钮点击事件。"""
        user_confirmed_load = True
        if self.network_manager.get_all_devices(): # 检查是否有设备
            if not self.suppress_confirmations:
                reply = QMessageBox.question(self, "确认加载", "加载配置将覆盖当前所有设备和连接，确定吗？",
                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
                user_confirmed_load = (reply == QMessageBox.StandardButton.Yes)

        if not user_confirmed_load:
            print("用户取消加载配置。")
            return

        filepath, _ = QFileDialog.getOpenFileName(self, "加载项目配置", "", "JSON 文件 (*.json);;所有文件 (*)")
        if not filepath:
            return

        if self.network_manager.load_project(filepath):
            # 加载成功后，需要完全刷新 UI
            self.device_tablewidget.setRowCount(0) # 清空表格
            for dev in self.network_manager.get_all_devices():
                self._add_device_to_table(dev) # 重新填充表格

            self._update_device_combos()
            self.node_positions = None # 重置布局
            self.selected_node_id = None
            self._update_port_totals_display()
            self._update_connection_views() # 重新绘制所有视图
            # 更新填充按钮状态
            has_connections = bool(self.network_manager.get_all_connections())
            can_fill = has_connections or any(bool(dev.get_all_available_ports()) for dev in self.network_manager.get_all_devices())
            self.fill_mesh_button.setEnabled(can_fill)
            self.fill_ring_button.setEnabled(can_fill)

            QMessageBox.information(self, "成功", f"项目配置已从以下文件加载:\n{filepath}")
        else:
            # 加载失败时，NetworkManager 内部会尝试清空状态
            self.device_tablewidget.setRowCount(0)
            self._update_device_combos()
            self.node_positions = None
            self.selected_node_id = None
            self._update_port_totals_display()
            self._update_connection_views()
            self.fill_mesh_button.setEnabled(False)
            self.fill_ring_button.setEnabled(False)
            QMessageBox.critical(self, "加载失败", f"无法加载项目配置文件:\n{filepath}")


    @Slot()
    def export_connections(self):
        """处理“导出列表”按钮点击事件。"""
        connections = self.network_manager.get_all_connections()
        if not connections:
            QMessageBox.warning(self, "提示", "没有连接结果可导出。")
            return
        # 调用 utils 中的导出函数
        export_connections_to_file(self, connections) # 传递 self 作为父窗口

    @Slot()
    def export_topology(self):
        """处理“导出拓扑图”按钮点击事件。"""
        if not self.mpl_canvas.fig or not self.network_manager.get_all_devices():
            QMessageBox.warning(self, "提示", "没有拓扑图可导出（请先添加设备并计算）。")
            return
        # 调用 utils 中的导出函数
        export_topology_to_file(self, self.mpl_canvas.fig) # 传递 self 和 figure 对象

    @Slot()
    def export_html_report(self):
        """处理“导出报告 (HTML)”按钮点击事件。"""
        devices = self.network_manager.get_all_devices()
        connections = self.network_manager.get_all_connections()
        if not devices or not self.mpl_canvas.fig:
             QMessageBox.warning(self, "无法导出", "请先添加设备并生成拓扑图。")
             return
        # 调用 utils 中的导出函数
        export_report_to_html(self, self.mpl_canvas.fig, connections)


    @Slot()
    def add_manual_connection(self):
        """处理手动编辑标签页中的“添加连接”按钮。"""
        dev1_id = self.edit_dev1_combo.currentData()
        dev2_id = self.edit_dev2_combo.currentData()
        port1_text = self.edit_port1_combo.currentText()
        port2_text = self.edit_port2_combo.currentText()

        if dev1_id is None or dev2_id is None or port1_text == "选择端口..." or port2_text == "选择端口...":
            QMessageBox.warning(self, "选择不完整", "请选择两个设备和它们各自的端口。")
            return
        if dev1_id == dev2_id:
            QMessageBox.warning(self, "选择错误", "不能将设备连接到自身。")
            return

        # 调用 NetworkManager 添加连接
        added_connection = self.network_manager.add_connection(dev1_id, port1_text, dev2_id, port2_text)

        if added_connection:
            # 更新 UI
            self.node_positions = None # 可能需要重新布局
            self.selected_node_id = None
            self._update_connection_views() # 更新列表和拓扑图
            self._update_device_table_connections() # 更新表格连接数
            self._update_manual_port_options() # 更新可用端口下拉框
            self._update_port_totals_display()
            # 更新填充按钮状态
            has_connections = True
            can_fill = True # 添加了连接，总能尝试填充
            self.fill_mesh_button.setEnabled(can_fill)
            self.fill_ring_button.setEnabled(can_fill)
            print(f"成功添加手动连接: {added_connection[0].name}[{port1_text}] <-> {added_connection[2].name}[{port2_text}]")
        else:
            # NetworkManager 内部会打印错误原因
            QMessageBox.warning(self, "添加失败", "无法添加手动连接，请检查端口兼容性、可用性或查看控制台输出。")
            self._update_manual_port_options() # 刷新端口列表以防状态不同步


    @Slot()
    def remove_manual_connection(self):
        """处理手动编辑标签页中的“移除选中连接”按钮。"""
        selected_items = self.manual_connection_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "提示", "请在下方列表中选择要移除的连接。")
            return

        connections_to_remove_data = []
        items_to_remove_from_widget = []
        for item in selected_items:
            conn_data = item.data(Qt.ItemDataRole.UserRole)
            if conn_data:
                connections_to_remove_data.append(conn_data)
                items_to_remove_from_widget.append(item)

        if not connections_to_remove_data:
            return

        removed_count = 0
        for conn_data in connections_to_remove_data:
            dev1, port1, dev2, port2, _ = conn_data
            # 调用 NetworkManager 移除连接
            if self.network_manager.remove_connection(dev1.id, port1, dev2.id, port2):
                removed_count += 1
            else:
                print(f"警告: 尝试从管理器移除连接时失败: {dev1.name}[{port1}] <-> {dev2.name}[{port2}]")

        if removed_count > 0:
            # 更新 UI
            self.node_positions = None # 可能需要重新布局
            self.selected_node_id = None
            # 从 QListWidget 中移除项 (从后往前)
            # 获取需要移除的行号
            rows_to_remove = sorted([self.manual_connection_list.row(item) for item in items_to_remove_from_widget], reverse=True)
            for row in rows_to_remove:
                 if row != -1:
                      self.manual_connection_list.takeItem(row) # 按行号移除

            self._update_connection_views() # 更新文本列表和拓扑图
            self._update_device_table_connections() # 更新表格连接数
            self._update_manual_port_options() # 更新可用端口
            self._update_port_totals_display()
            print(f"成功移除了 {removed_count} 条连接。")
            # 更新填充按钮状态
            has_connections = bool(self.network_manager.get_all_connections())
            can_fill = has_connections or any(bool(dev.get_all_available_ports()) for dev in self.network_manager.get_all_devices())
            self.fill_mesh_button.setEnabled(can_fill)
            self.fill_ring_button.setEnabled(can_fill)
            self.remove_manual_button.setEnabled(has_connections)
        else:
            print("没有连接被移除。")


    @Slot(int)
    def _toggle_suppress_confirmations(self, state):
        """更新是否跳过确认弹窗的状态。"""
        self.suppress_confirmations = (state == Qt.CheckState.Checked.value)
        print(f"跳过确认弹窗: {'已启用' if self.suppress_confirmations else '已禁用'}")

    @Slot(QTableWidgetItem)
    def on_device_item_changed(self, item: QTableWidgetItem):
        """处理设备表格中项目被编辑后的事件。"""
        if not item: return
        row = item.row()
        col = item.column()
        name_item = self.device_tablewidget.item(row, COL_NAME)
        if not name_item: return
        dev_id = name_item.data(Qt.ItemDataRole.UserRole)
        if dev_id is None: return # 无效行

        self.device_tablewidget.blockSignals(True) # 阻止信号循环

        success = False
        try:
            if col == COL_NAME:
                new_name = item.text().strip()
                if self.network_manager.update_device(dev_id, new_name=new_name):
                     success = True
                     self._update_device_combos() # 更新下拉框名称
                     self._update_connection_views() # 更新图标签
                else:
                     # 更新失败，恢复旧名称
                     device = self.network_manager.get_device_by_id(dev_id)
                     if device: item.setText(device.name)
                     QMessageBox.warning(self, "重命名失败", f"无法将设备重命名为 '{new_name}' (可能名称冲突或为空)。")

            elif col in [COL_MPO, COL_LC, COL_SFP]:
                port_attr_map = {COL_MPO: 'mpo', COL_LC: 'lc', COL_SFP: 'sfp'}
                port_name_map = {COL_MPO: 'MPO', COL_LC: 'LC', COL_SFP: 'SFP+'}
                attr_suffix = port_attr_map.get(col)
                port_type_name = port_name_map.get(col)
                device = self.network_manager.get_device_by_id(dev_id)
                if not device: raise ValueError("找不到设备")

                # 检查是否允许修改此类型的端口
                is_uhd_horizon = device.type in UHD_TYPES
                is_micron = device.type == DEV_MN
                can_edit_this_port = (attr_suffix in ['mpo', 'lc'] and is_uhd_horizon) or \
                                     (attr_suffix == 'sfp' and is_micron)

                if not can_edit_this_port:
                     old_count = getattr(device, f"{attr_suffix}_total", 0)
                     print(f"不允许修改设备类型 '{device.type}' 的 '{port_type_name}' 端口数量。")
                     item.setText(str(old_count))
                else:
                    old_count = getattr(device, f"{attr_suffix}_total", 0)
                    try:
                        new_count = int(item.text().strip())
                        if new_count < 0: raise ValueError("端口数不能为负")
                        if new_count != old_count:
                            user_confirmed_port_change = True
                            if not self.suppress_confirmations:
                                reply = QMessageBox.question(self, "确认修改端口数量",
                                                             f"修改设备 '{device.name}' 的 {port_type_name} 端口数量将清除所有现有连接并可能需要重新计算布局。\n是否继续？",
                                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
                                user_confirmed_port_change = (reply == QMessageBox.StandardButton.Yes)

                            if user_confirmed_port_change:
                                update_kwargs = {f"new_{attr_suffix}": new_count}
                                if self.network_manager.update_device(dev_id, **update_kwargs):
                                    success = True
                                    # NetworkManager 内部已清除连接，这里只需更新 UI
                                    self.clear_results() # 会更新所有视图
                                    self._update_port_totals_display()
                                    self._update_manual_port_options()
                                else:
                                     # 更新失败，恢复旧值
                                     item.setText(str(old_count))
                                     QMessageBox.warning(self, "更新失败", f"更新 {port_type_name} 端口数量失败。")
                            else:
                                print("用户取消修改端口数量。")
                                item.setText(str(old_count)) # 恢复旧值
                        else:
                             success = True # 值未变，也算成功

                    except ValueError:
                        QMessageBox.warning(self, "输入错误", f"{port_type_name} 端口数量必须是非负整数。")
                        item.setText(str(old_count)) # 恢复旧值
        finally:
            self.device_tablewidget.blockSignals(False)


    @Slot(QTableWidgetItem)
    def show_device_details_from_table(self, item: QTableWidgetItem):
        """处理设备表格双击事件，显示设备详情。"""
        if not item: return
        row = item.row()
        name_item = self.device_tablewidget.item(row, COL_NAME)
        if not name_item: return
        dev_id = name_item.data(Qt.ItemDataRole.UserRole)
        device = self.network_manager.get_device_by_id(dev_id)
        if not device:
            QMessageBox.critical(self, "错误", "无法找到所选设备的详细信息。")
            return
        self._display_device_details_popup(device)

    @Slot()
    def on_layout_change(self):
        """处理布局下拉框选择变化事件。"""
        # 只要有设备或连接，改变布局就重置位置并重绘
        if self.network_manager.get_all_devices():
            print("布局选择已更改，重置节点位置并重绘。")
            self.node_positions = None
            self.selected_node_id = None
            self._update_connection_views()

    @Slot(str)
    def filter_device_table(self, text: str):
        """根据输入过滤设备表格。"""
        filter_text = text.lower()
        for row in range(self.device_tablewidget.rowCount()):
            match = False
            name_item = self.device_tablewidget.item(row, COL_NAME)
            type_item = self.device_tablewidget.item(row, COL_TYPE)
            if name_item and filter_text in name_item.text().lower():
                match = True
            if not match and type_item and filter_text in type_item.text().lower():
                match = True
            self.device_tablewidget.setRowHidden(row, not match)

    @Slot()
    def filter_connection_list(self):
        """根据下拉框和输入框过滤手动编辑中的连接列表。"""
        selected_type = self.conn_filter_type_combo.currentText()
        filter_device_text = self.conn_filter_device_entry.text().strip().lower()
        type_filter_active = selected_type != "所有类型"

        for i in range(self.manual_connection_list.count()):
            item = self.manual_connection_list.item(i)
            conn_data = item.data(Qt.ItemDataRole.UserRole)

            if not conn_data or not isinstance(conn_data, tuple) or len(conn_data) != 5:
                item.setHidden(False) # 显示无效项？或者隐藏？暂时显示
                continue

            dev1, _, dev2, _, conn_type = conn_data
            item_conn_type = conn_type
            dev1_name_lower = dev1.name.lower()
            dev2_name_lower = dev2.name.lower()

            type_match = True
            if type_filter_active:
                type_match = (item_conn_type == selected_type)

            device_match = True
            if filter_device_text:
                device_match = (filter_device_text in dev1_name_lower or filter_device_text in dev2_name_lower)

            item.setHidden(not (type_match and device_match))


    # --- UI 更新辅助方法 ---

    def _add_device_to_table(self, device: Device):
        """将设备对象添加到 UI 表格中。"""
        row_position = self.device_tablewidget.rowCount()
        self.device_tablewidget.insertRow(row_position)

        # 设置单元格的可编辑和不可编辑标志
        editable_flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable
        non_editable_flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled

        # 创建单元格项
        name_item = QTableWidgetItem(device.name)
        name_item.setData(Qt.ItemDataRole.UserRole, device.id) # 存储 ID 以便后续查找
        name_item.setFlags(editable_flags) # 名称可编辑

        type_item = QTableWidgetItem(device.type)
        type_item.setFlags(non_editable_flags) # 类型不可编辑

        mpo_item = NumericTableWidgetItem(str(device.mpo_total))
        mpo_item.setData(Qt.ItemDataRole.UserRole + 1, device.mpo_total) # 存储数值用于排序
        mpo_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        mpo_item.setFlags(editable_flags if device.type in UHD_TYPES else non_editable_flags) # 仅 UHD/HorizoN 可编辑 MPO

        lc_item = NumericTableWidgetItem(str(device.lc_total))
        lc_item.setData(Qt.ItemDataRole.UserRole + 1, device.lc_total)
        lc_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        lc_item.setFlags(editable_flags if device.type in UHD_TYPES else non_editable_flags) # 仅 UHD/HorizoN 可编辑 LC

        sfp_item = NumericTableWidgetItem(str(device.sfp_total))
        sfp_item.setData(Qt.ItemDataRole.UserRole + 1, device.sfp_total)
        sfp_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        sfp_item.setFlags(editable_flags if device.type == DEV_MN else non_editable_flags) # 仅 MicroN 可编辑 SFP

        conn_val = float(f"{device.connections:.2f}") # 获取当前连接数
        conn_item = NumericTableWidgetItem(f"{conn_val:.2f}")
        conn_item.setData(Qt.ItemDataRole.UserRole + 1, conn_val)
        conn_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        conn_item.setFlags(non_editable_flags) # 连接数不可编辑

        # 将单元格项放入表格
        self.device_tablewidget.setItem(row_position, COL_NAME, name_item)
        self.device_tablewidget.setItem(row_position, COL_TYPE, type_item)
        self.device_tablewidget.setItem(row_position, COL_MPO, mpo_item)
        self.device_tablewidget.setItem(row_position, COL_LC, lc_item)
        self.device_tablewidget.setItem(row_position, COL_SFP, sfp_item)
        self.device_tablewidget.setItem(row_position, COL_CONN, conn_item)

    def _update_device_table_connections(self):
        """更新设备表格中的“连接数”列。"""
        for row in range(self.device_tablewidget.rowCount()):
            name_item = self.device_tablewidget.item(row, COL_NAME)
            if name_item:
                dev_id = name_item.data(Qt.ItemDataRole.UserRole)
                device = self.network_manager.get_device_by_id(dev_id)
                if device:
                    conn_val = float(f"{device.connections:.2f}")
                    conn_item = self.device_tablewidget.item(row, COL_CONN)
                    if conn_item:
                        # 更新现有项
                        conn_item.setText(f"{conn_val:.2f}")
                        conn_item.setData(Qt.ItemDataRole.UserRole + 1, conn_val)
                    else:
                        # 创建新项 (理论上不应发生，除非表格不完整)
                        conn_item = NumericTableWidgetItem(f"{conn_val:.2f}")
                        conn_item.setData(Qt.ItemDataRole.UserRole + 1, conn_val)
                        conn_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                        conn_item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                        self.device_tablewidget.setItem(row, COL_CONN, conn_item)

    def _update_device_combos(self):
        """更新手动编辑区域的设备下拉框选项。"""
        self.edit_dev1_combo.blockSignals(True)
        self.edit_dev2_combo.blockSignals(True)

        current_dev1_id = self.edit_dev1_combo.currentData()
        current_dev2_id = self.edit_dev2_combo.currentData()

        self.edit_dev1_combo.clear()
        self.edit_dev2_combo.clear()
        self.edit_dev1_combo.addItem("选择设备 1...", userData=None)
        self.edit_dev2_combo.addItem("选择设备 2...", userData=None)

        # 从 NetworkManager 获取设备列表并排序
        sorted_devices = sorted(self.network_manager.get_all_devices(), key=lambda dev: dev.name)

        idx1_to_select = 0
        idx2_to_select = 0
        for i, dev in enumerate(sorted_devices):
            item_text = f"{dev.name} ({dev.type})"
            self.edit_dev1_combo.addItem(item_text, userData=dev.id)
            self.edit_dev2_combo.addItem(item_text, userData=dev.id)
            # 检查是否需要恢复之前的选择
            if dev.id == current_dev1_id:
                idx1_to_select = i + 1
            if dev.id == current_dev2_id:
                idx2_to_select = i + 1

        self.edit_dev1_combo.setCurrentIndex(idx1_to_select)
        self.edit_dev2_combo.setCurrentIndex(idx2_to_select)

        self.edit_dev1_combo.blockSignals(False)
        self.edit_dev2_combo.blockSignals(False)

        # 设备变化后需要更新端口选项
        self._update_manual_port_options()

    def _populate_edit_port_combos(self, device_combo_to_populate: QComboBox, port_combo_to_populate: QComboBox,
                                   other_device_combo: QComboBox, other_port_combo: QComboBox):
        """
        动态填充指定的端口下拉列表，并根据另一侧的选择进行过滤。

        Args:
            device_combo_to_populate: 需要为其填充端口的设备下拉框。
            port_combo_to_populate: 需要被填充的端口下拉框。
            other_device_combo: 对端的设备下拉框。
            other_port_combo: 对端的端口下拉框。
        """
        port_combo_to_populate.blockSignals(True)
        current_port_selection = port_combo_to_populate.currentText() # 记录当前选择
        port_combo_to_populate.clear()
        port_combo_to_populate.addItem("选择端口...")
        port_combo_to_populate.setEnabled(False) # 默认禁用

        dev_id = device_combo_to_populate.currentData()
        if dev_id is not None:
            device = self.network_manager.get_device_by_id(dev_id)
            if device:
                # 从 NetworkManager 获取可用端口
                available_ports = self.network_manager.get_available_ports(dev_id)
                ports_to_add = []

                # 获取对端信息以进行兼容性过滤
                other_dev_id = other_device_combo.currentData()
                other_port_name = other_port_combo.currentText()

                if other_dev_id is not None and other_port_name != "选择端口...":
                    # 获取兼容的端口类型
                    compatible_types_here = self.network_manager.get_compatible_port_types(other_dev_id, other_port_name)
                    # 过滤可用端口
                    for port in available_ports:
                        port_type_here = get_port_type_from_name(port)
                        if port_type_here in compatible_types_here:
                            ports_to_add.append(port)
                else:
                    # 如果对端未完全选定，显示所有可用端口
                    ports_to_add = available_ports

                # 填充下拉框
                if ports_to_add:
                    port_combo_to_populate.addItems(ports_to_add)
                    # 尝试恢复之前的选择
                    index_to_select = port_combo_to_populate.findText(current_port_selection)
                    port_combo_to_populate.setCurrentIndex(index_to_select if index_to_select != -1 else 0)
                    port_combo_to_populate.setEnabled(True) # 启用下拉框
                else:
                    port_combo_to_populate.addItem("无兼容/可用端口")
                    port_combo_to_populate.setCurrentIndex(1) # 显示提示信息

        port_combo_to_populate.blockSignals(False)

    @Slot()
    def _update_manual_port_options(self):
        """统一更新手动添加连接中的两个端口下拉列表的选项。"""
        # 更新 Port1 下拉框，根据 Dev2 和 Port2 的选择进行过滤
        self._populate_edit_port_combos(self.edit_dev1_combo, self.edit_port1_combo,
                                        self.edit_dev2_combo, self.edit_port2_combo)
        # 更新 Port2 下拉框，根据 Dev1 和 Port1 的选择进行过滤
        self._populate_edit_port_combos(self.edit_dev2_combo, self.edit_port2_combo,
                                        self.edit_dev1_combo, self.edit_port1_combo)

    def _update_port_totals_display(self):
        """更新显示端口总数的标签。"""
        # 从 NetworkManager 获取总数
        totals = self.network_manager.calculate_port_totals()
        self.port_totals_label.setText(f"总计: {PORT_MPO}: {totals['mpo']}, {PORT_LC}: {totals['lc']}, {PORT_SFP}+: {totals['sfp']}")


    def _update_connection_views(self):
        """更新连接列表文本框、手动编辑列表和拓扑图。"""
        # 1. 更新连接列表文本框 (QTextEdit)
        self.connections_textedit.clear()
        connections = self.network_manager.get_all_connections()
        if connections:
            self.connections_textedit.append("<b>连接列表:</b><hr>")
            for i, conn in enumerate(connections):
                dev1, port1, dev2, port2, conn_type = conn
                self.connections_textedit.append(f"{i+1}. {dev1.name} [{port1}] &lt;-&gt; {dev2.name} [{port2}] ({conn_type})")
        else:
            self.connections_textedit.append("无连接。")

        # 2. 更新手动编辑中的连接列表 (QListWidget)
        self.manual_connection_list.clear()
        if connections:
            for i, conn in enumerate(connections):
                dev1, port1, dev2, port2, conn_type = conn
                item_text = f"{i+1}. {dev1.name} [{port1}] <-> {dev2.name} [{port2}] ({conn_type})"
                item = QListWidgetItem(item_text)
                item.setData(Qt.ItemDataRole.UserRole, conn) # 存储连接元组数据
                self.manual_connection_list.addItem(item)
        # 根据是否有连接更新移除按钮状态
        self.remove_manual_button.setEnabled(bool(connections))
        self.filter_connection_list() # 应用当前过滤器

        # 3. 更新拓扑图
        selected_layout = self.layout_combo.currentText().lower()
        # 从 NetworkManager 获取绘图所需数据
        devices_for_plot = self.network_manager.get_all_devices()
        connections_for_plot = self.network_manager.get_all_connections()
        port_totals = self.network_manager.calculate_port_totals()

        # 调用 MplCanvas 的绘图方法
        self.fig, calculated_pos = self.mpl_canvas.plot_topology(
            devices_for_plot,
            connections_for_plot,
            layout_algorithm=selected_layout,
            fixed_pos=self.node_positions,
            selected_node_id=self.selected_node_id,
            port_totals_dict=port_totals
        )

        # 如果布局是新计算的，则存储它 (除非正在拖动)
        if calculated_pos is not None and self.dragged_node_id is None:
            # 仅在首次计算或布局更改后存储
            # 或者在拖动结束后更新 (在 on_canvas_release 中处理可能更好)
            if self.node_positions is None or selected_layout != getattr(self, '_last_layout_used', None):
                self.node_positions = calculated_pos
                setattr(self, '_last_layout_used', selected_layout) # 记录当前使用的布局


        # 4. 更新导出按钮状态
        has_connections = bool(connections)
        has_devices = bool(devices_for_plot)
        has_figure = self.fig is not None and has_devices

        self.export_list_button.setEnabled(has_connections)
        self.export_topo_button.setEnabled(has_figure)
        self.export_report_button.setEnabled(has_connections and has_figure)


    def _display_device_details_popup(self, dev: Device):
        """显示包含设备详细信息的弹出窗口。"""
        details = f"ID: {dev.id}\n名称: {dev.name}\n类型: {dev.type}\n"
        # 从设备对象获取可用端口信息
        avail_ports = dev.get_all_available_ports()
        avail_lc_count = sum(1 for p in avail_ports if p.startswith(PORT_LC))
        avail_sfp_count = sum(1 for p in avail_ports if p.startswith(PORT_SFP))
        avail_mpo_ch_count = sum(1 for p in avail_ports if p.startswith(PORT_MPO))

        if dev.type in UHD_TYPES:
            details += f"{PORT_MPO} 端口总数: {dev.mpo_total}\n"
            details += f"{PORT_LC} 端口总数: {dev.lc_total}\n"
            details += f"可用 {PORT_MPO} 子通道: {avail_mpo_ch_count}\n"
            details += f"可用 {PORT_LC} 端口: {avail_lc_count}\n"
        elif dev.type == DEV_MN:
            details += f"{PORT_SFP}+ 端口总数: {dev.sfp_total}\n"
            details += f"可用 {PORT_SFP}+ 端口: {avail_sfp_count}\n"

        details += f"当前连接数 (估算): {dev.connections:.2f}\n"

        if dev.port_connections:
            details += "\n端口连接详情:\n"
            # 分类并排序显示
            lc_conns = {p: t for p, t in dev.port_connections.items() if p.startswith(PORT_LC)}
            sfp_conns = {p: t for p, t in dev.port_connections.items() if p.startswith(PORT_SFP)}
            mpo_conns_grouped = defaultdict(dict)
            for p, t in dev.port_connections.items():
                if p.startswith(PORT_MPO):
                    mpo_base = p.split('-')[0]
                    mpo_conns_grouped[mpo_base][p] = t

            if lc_conns:
                details += f"  {PORT_LC} 连接:\n"
                # 按端口号排序
                for port in sorted(lc_conns.keys(), key=lambda x: int(x[len(PORT_LC):])):
                    details += f"    {port} -> {lc_conns[port]}\n"
            if sfp_conns:
                details += f"  {PORT_SFP}+ 连接:\n"
                # 按端口号排序
                for port in sorted(sfp_conns.keys(), key=lambda x: int(x[len(PORT_SFP):])):
                    details += f"    {port} -> {sfp_conns[port]}\n"
            if mpo_conns_grouped:
                details += f"  {PORT_MPO} 连接 (Breakout):\n"
                # 按 MPO 端口号排序
                for base_port in sorted(mpo_conns_grouped.keys(), key=lambda x: int(x[len(PORT_MPO):])):
                    details += f"    {base_port}:\n"
                    # 按通道号排序
                    for port in sorted(mpo_conns_grouped[base_port].keys(), key=lambda x: int(x.split('-Ch')[-1])):
                        details += f"      {port} -> {mpo_conns_grouped[base_port][port]}\n"

        QMessageBox.information(self, f"设备详情 - {dev.name}", details)

    def _get_matplotlib_font_prop(self):
         """获取用于 Matplotlib 的 FontProperties 对象"""
         # 尝试使用加载的字体，否则回退
         try:
             if hasattr(self, 'chinese_font') and self.chinese_font.family():
                 # 检查字体是否真的被 Matplotlib 识别
                 if self.chinese_font.family() in [f.name for f in font_manager.fontManager.ttflist]:
                      return font_manager.FontProperties(family=self.chinese_font.family())
                 else:
                      # 如果 Qt 加载了但 Matplotlib 没找到，尝试添加路径
                      # 注意：QFontDatabase 可能需要 QApplication 实例存在
                      db = QFontDatabase()
                      # 查找字体家族对应的文件路径
                      # 注意：QFontDatabase.applicationFontFiles 可能需要 font ID 而不是 family name
                      # 这是一个复杂的问题，取决于字体如何被 Qt 注册
                      # 简化处理：如果 Matplotlib 找不到，就用默认
                      print(f"警告: Qt 字体 '{self.chinese_font.family()}' 未在 Matplotlib 字体列表中找到。")

         except Exception as e:
              print(f"获取 Matplotlib 字体属性时出错: {e}")
         # 回退到默认 sans-serif
         return font_manager.FontProperties(family='sans-serif')


    # --- 画布事件处理方法 (Canvas Event Handlers & Helpers) ---
    # 这些方法现在操作 self.network_manager 来修改状态

    def _get_node_at_event(self, event) -> Optional[int]:
        """辅助函数：查找鼠标事件位置下的节点 ID。"""
        # 检查事件是否发生在 MplCanvas 的 axes 上，并且有有效的节点位置信息
        if event.inaxes != self.mpl_canvas.axes or not self.node_positions:
            return None
        # 获取事件的坐标数据
        x, y = event.xdata, event.ydata
        # 如果坐标无效（例如，点击在图形区域外），则返回 None
        if x is None or y is None:
            return None

        clicked_node_id = None
        min_dist_sq = float('inf') # 初始化最小距离平方为无穷大

        # 使用动态阈值，基于当前视图范围，避免固定阈值在缩放时失效
        xlim = self.mpl_canvas.axes.get_xlim()
        ylim = self.mpl_canvas.axes.get_ylim()
        # 阈值设为视图对角线长度的一小部分 (例如 3%) 的平方
        # 这个比例可以根据需要调整
        threshold_dist_sq = ((xlim[1]-xlim[0])**2 + (ylim[1]-ylim[0])**2) * (0.03**2)

        # 遍历所有节点位置
        for node_id, (nx, ny) in self.node_positions.items():
            dist_sq = (x - nx)**2 + (y - ny)**2 # 计算点击位置到节点的距离平方
            # 如果距离小于当前最小距离，并且小于阈值
            if dist_sq < min_dist_sq and dist_sq < threshold_dist_sq:
                min_dist_sq = dist_sq # 更新最小距离
                clicked_node_id = node_id # 记录潜在的点击节点 ID
        # 返回找到的最近且在阈值内的节点 ID，否则返回 None
        return clicked_node_id

    def _start_node_drag(self, node_id: int, event_xdata: float, event_ydata: float):
        """辅助函数：开始节点拖动。"""
        # 确认节点位置信息存在
        if self.node_positions is None or node_id not in self.node_positions:
            print(f"警告: 尝试拖动未找到位置的节点 ID {node_id}")
            return

        self.dragged_node_id = node_id # 记录正在拖动的节点
        self.connecting_node_id = None # 确保不在连接模式
        nx, ny = self.node_positions[node_id] # 获取节点当前位置
        # 计算鼠标按下位置相对于节点位置的偏移量
        self.drag_offset = (event_xdata - nx, event_ydata - ny)

        # 如果点击的节点不是当前选中的，则选中它并重绘以更新高亮
        if self.selected_node_id != node_id:
            self.selected_node_id = node_id
            print(f"选中节点 (准备拖动): ID={self.selected_node_id}")
            self._update_connection_views() # 更新视图以显示高亮
        else:
            # 如果已经是选中状态，直接开始拖动
            print(f"开始拖动节点: ID={self.selected_node_id}")

    def _start_connection_drag(self, node_id: int):
        """辅助函数：开始连接拖动 (Shift+Click)。"""
        # 确认节点位置信息存在
        if self.node_positions is None or node_id not in self.node_positions:
             print(f"警告: 尝试从未知位置的节点 ID {node_id} 开始连接拖动")
             return

        print(f"开始连接拖动: 从 ID={node_id}")
        self.connecting_node_id = node_id # 记录连接的起始节点
        self.dragged_node_id = None # 确保不在普通节点拖动模式

        # 如果点击的节点不是当前选中的，则选中它并重绘以更新高亮
        if self.selected_node_id != node_id:
            self.selected_node_id = node_id
            self._update_connection_views() # 更新视图以显示高亮

    def _handle_background_press(self):
        """辅助函数：处理画布背景点击，清除选择和拖动状态。"""
        # 检查是否有状态需要清除（选中、拖动节点、拖动连接线）
        needs_redraw = self.selected_node_id is not None or \
                       self.dragged_node_id is not None or \
                       self.connecting_node_id is not None

        # 重置所有状态变量
        self.selected_node_id = None
        self.dragged_node_id = None
        self.connecting_node_id = None

        # 如果之前有状态，需要重绘以清除高亮或临时线
        if needs_redraw:
            print("清除选中/状态 (点击背景)")
            self._update_connection_views() # 更新视图

    def _end_node_drag(self):
        """辅助函数：结束节点拖动。"""
        # 确认确实有节点在被拖动
        if self.dragged_node_id is not None:
            print(f"结束拖动节点: ID={self.dragged_node_id}")
            # 记录最终位置（self.node_positions 已在 motion 事件中更新）
            # 可以考虑在这里触发一次保存布局的操作（如果需要持久化）
            self.dragged_node_id = None # 重置拖动状态
            # 可以选择性地触发一次绘制以确保最终状态
            # self.mpl_canvas.draw_idle()
        else:
             # 这种情况理论上不应发生，但加上日志以防万一
             print("调试: _end_node_drag 被调用但 self.dragged_node_id 为 None")


    def _end_connection_drag(self, event):
        """辅助函数：结束连接拖动并尝试通过 NetworkManager 创建连接。"""
        start_node_id = self.connecting_node_id
        self.connecting_node_id = None # 重置连接拖动状态

        # 清理画布上的临时连接线
        if self.connection_line:
            try:
                self.connection_line.remove() # 从 axes 移除
            except ValueError: # 可能已经被移除或无效
                pass
            finally:
                 self.connection_line = None # 重置引用
            self.mpl_canvas.draw_idle() # 更新画布以移除线条

        # 检查起始节点 ID 是否有效
        if start_node_id is None:
            print("调试: _end_connection_drag 启动节点 ID 为 None")
            return

        # 查找鼠标释放位置的节点
        target_node_id = self._get_node_at_event(event)

        # 检查目标节点是否有效且不是起始节点本身
        if target_node_id is not None and target_node_id != start_node_id:
            # **修复: 调用 NetworkManager 的 add_best_connection 方法**
            print(f"尝试通过拖拽连接: ID {start_node_id} -> ID {target_node_id}")
            added_connection = self.network_manager.add_best_connection(start_node_id, target_node_id)

            if added_connection:
                 # 连接成功，更新 UI
                 self.node_positions = None # 重置布局，让 NetworkX 重新计算
                 self.selected_node_id = None # 清除选择
                 self._update_connection_views()
                 self._update_device_table_connections()
                 self._update_manual_port_options()
                 self._update_port_totals_display()
                 # 更新填充按钮状态
                 has_connections = bool(self.network_manager.get_all_connections())
                 can_fill = has_connections or any(bool(dev.get_all_available_ports()) for dev in self.network_manager.get_all_devices())
                 self.fill_mesh_button.setEnabled(can_fill)
                 self.fill_ring_button.setEnabled(can_fill)
                 print(f"成功通过拖拽添加连接: {added_connection[0].name}[{added_connection[1]}] <-> {added_connection[2].name}[{added_connection[3]}]")
            else:
                 # 连接失败 (NetworkManager 内部已打印原因)
                 dev1 = self.network_manager.get_device_by_id(start_node_id)
                 dev2 = self.network_manager.get_device_by_id(target_node_id)
                 dev1_name = dev1.name if dev1 else f"ID {start_node_id}"
                 dev2_name = dev2.name if dev2 else f"ID {target_node_id}"
                 QMessageBox.warning(self, "连接失败", f"无法在 {dev1_name} 和 {dev2_name} 之间自动添加连接（可能无可用兼容端口）。")
                 self._update_manual_port_options() # 刷新端口列表
        else:
            # 目标无效或与起始节点相同
            print("连接拖动取消或目标无效/相同。")


    @Slot(object)
    def on_canvas_press(self, event):
        """处理画布上的鼠标按下事件。"""
        # 清理可能残留的临时连接线
        if self.connection_line:
            try: self.connection_line.remove(); self.connection_line = None
            except ValueError: pass
            # 不需要立即 draw_idle，后续操作会触发

        # 首先检查点击是否在坐标轴内且坐标有效
        if event.inaxes != self.mpl_canvas.axes or event.xdata is None or event.ydata is None:
            self._handle_background_press() # 点击在画布外或无效区域
            return

        clicked_node_id = self._get_node_at_event(event) # 查找点击位置的节点
        modifiers = QGuiApplication.keyboardModifiers()
        is_shift_pressed = modifiers == Qt.KeyboardModifier.ShiftModifier

        if event.dblclick:
            # 双击事件
            self.dragged_node_id = None # 取消任何拖动状态
            self.connecting_node_id = None
            if clicked_node_id is not None:
                device = self.network_manager.get_device_by_id(clicked_node_id)
                if device:
                    self._display_device_details_popup(device) # 显示详情
            # 双击背景无操作
        elif event.button == 1: # 鼠标左键单击
            if clicked_node_id is not None: # 点击在节点上
                if is_shift_pressed:
                    self._start_connection_drag(clicked_node_id) # Shift+单击 开始连接拖动
                else:
                    self._start_node_drag(clicked_node_id, event.xdata, event.ydata) # 普通单击 开始节点拖动
            else: # 点击在背景上
                self._handle_background_press() # 清除选择等

    @Slot(object)
    def on_canvas_motion(self, event):
        """处理画布上的鼠标移动事件。"""
        if event.inaxes != self.mpl_canvas.axes or event.xdata is None or event.ydata is None:
            return # 忽略画布外的移动

        x, y = event.xdata, event.ydata

        # 处理节点拖动
        if self.dragged_node_id is not None and event.button == 1 and self.node_positions:
            # 检查 self.node_positions 是否包含该节点
            if self.dragged_node_id in self.node_positions:
                 new_x = x - self.drag_offset[0]
                 new_y = y - self.drag_offset[1]
                 self.node_positions[self.dragged_node_id] = (new_x, new_y)
                 self._update_connection_views() # 实时重绘以更新节点位置
            else:
                 print(f"警告: 尝试拖动节点 {self.dragged_node_id} 但其不在 node_positions 中")
                 self.dragged_node_id = None # 重置拖动状态

        # 处理连接线拖动
        elif self.connecting_node_id is not None and event.button == 1 and self.node_positions:
            start_pos = self.node_positions.get(self.connecting_node_id)
            # **修复: 明确检查 start_pos 是否为 None**
            if start_pos is not None:
                # 移除旧的临时线
                if self.connection_line:
                    try: self.connection_line.remove(); self.connection_line = None
                    except ValueError: pass
                    except AttributeError: self.connection_line = None # 防御性
                # 绘制新的临时线
                self.connection_line = Line2D([start_pos[0], x], [start_pos[1], y],
                                              ls='--', c='gray', lw=1.5,
                                              transform=self.mpl_canvas.axes.transData,
                                              zorder=10) # 确保线在顶层
                self.mpl_canvas.axes.add_line(self.connection_line)
                self.mpl_canvas.draw_idle() # 使用 idle 提高性能
            else:
                 print(f"警告: 尝试绘制连接线但起始节点 {self.connecting_node_id} 位置未知")
                 self.connecting_node_id = None # 重置状态

    @Slot(object)
    def on_canvas_release(self, event):
        """处理画布上的鼠标释放事件。"""
        if event.button == 1: # 左键释放
            if self.dragged_node_id is not None:
                self._end_node_drag() # 结束节点拖动
            elif self.connecting_node_id is not None:
                self._end_connection_drag(event) # 结束连接拖动并尝试创建连接

# 注意：这个文件现在只包含 MainWindow 类及其直接相关的 UI 设置和槽函数。
# MplCanvas, NumericTableWidgetItem, export 函数, resource_path 函数
# 已被（或将被）移动到它们各自的文件中。
# APP_STYLE 常量也留在这里，因为它直接被 MainWindow 使用。
