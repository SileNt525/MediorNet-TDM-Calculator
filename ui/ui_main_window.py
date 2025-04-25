# -*- coding: utf-8 -*-
"""
ui/ui_main_window.py

包含由 MainWindow 使用的 UI 定义类 Ui_MainWindow。
这个类负责创建和布局所有的 UI 控件。
"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit,
    QComboBox, QTextEdit, QTabWidget, QFrame, QSpacerItem, QSizePolicy,
    QGridLayout, QListWidgetItem, QAbstractItemView, QTableWidget,
    QTableWidgetItem, QHeaderView, QListWidget, QSplitter, QCheckBox,
    QMainWindow # 导入 QMainWindow 以便类型提示
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

# 从项目模块导入 (仅用于常量和类型提示)
# 注意：避免从 ui.main_window 导入，以防循环
try:
    from core.device import (
        DEV_UHD, DEV_HORIZON, DEV_MN,
        PORT_MPO, PORT_LC, PORT_SFP
    )
    # 导入自定义控件
    from .widgets import NumericTableWidgetItem
    # 导入画布类 (虽然不由 setupUi 创建，但可能需要知道类型)
    from .topology_canvas import MplCanvas
except ImportError as e:
    print(f"导入错误 (ui_main_window.py): {e}")
    # Fallbacks
    DEV_UHD, DEV_HORIZON, DEV_MN = 'MicroN UHD', 'HorizoN', 'MicroN'
    PORT_MPO, PORT_LC, PORT_SFP = 'MPO', 'LC', 'SFP'
    NumericTableWidgetItem = QTableWidgetItem
    MplCanvas = QWidget

# UI 常量
COL_NAME = 0
COL_TYPE = 1
COL_MPO = 2
COL_LC = 3
COL_SFP = 4
COL_CONN = 5

class Ui_MainWindow(object):
    """此类包含主窗口的 UI 定义和布局。"""

    def setupUi(self, MainWindow: QMainWindow):
        """
        创建和布局主窗口的 UI 控件。

        Args:
            MainWindow (QMainWindow): 要在其上设置 UI 的主窗口实例。
        """
        MainWindow.setObjectName("MainWindow")

        if not MainWindow.centralWidget():
             main_widget = QWidget(MainWindow)
             MainWindow.setCentralWidget(main_widget)
        else:
             main_widget = MainWindow.centralWidget()

        main_layout = QHBoxLayout(main_widget)
        main_splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(main_splitter)

        left_panel = QFrame()
        left_panel.setFrameShape(QFrame.Shape.StyledPanel)
        left_layout = QVBoxLayout(left_panel)
        main_splitter.addWidget(left_panel)

        # !! 获取字体，并确保在整个方法中使用这个局部变量 !!
        chinese_font = getattr(MainWindow, 'chinese_font', QFont())

        # --- 添加设备组 ---
        add_group = QFrame()
        add_group.setObjectName("addDeviceGroup")
        add_group_layout = QGridLayout(add_group)
        add_group_layout.setContentsMargins(10, 15, 10, 10)
        add_group_layout.setVerticalSpacing(8)
        add_title = QLabel("<b>添加新设备</b>")
        # !! 使用局部变量 chinese_font !!
        add_title.setFont(QFont(chinese_font.family(), 11))
        add_group_layout.addWidget(add_title, 0, 0, 1, 2, alignment=Qt.AlignmentFlag.AlignCenter)

        add_group_layout.addWidget(QLabel("类型:"), 1, 0)
        MainWindow.device_type_combo = QComboBox()
        MainWindow.device_type_combo.addItems([DEV_UHD, DEV_HORIZON, DEV_MN])
        MainWindow.device_type_combo.setFont(chinese_font) # !! 使用局部变量 !!
        add_group_layout.addWidget(MainWindow.device_type_combo, 1, 1)

        add_group_layout.addWidget(QLabel("名称:"), 2, 0)
        MainWindow.device_name_entry = QLineEdit()
        MainWindow.device_name_entry.setFont(chinese_font) # !! 使用局部变量 !!
        add_group_layout.addWidget(MainWindow.device_name_entry, 2, 1)

        MainWindow.mpo_label = QLabel(f"{PORT_MPO} 端口:")
        add_group_layout.addWidget(MainWindow.mpo_label, 3, 0)
        MainWindow.mpo_entry = QLineEdit("2")
        MainWindow.mpo_entry.setFont(chinese_font) # !! 使用局部变量 !!
        add_group_layout.addWidget(MainWindow.mpo_entry, 3, 1)a

        MainWindow.lc_label = QLabel(f"{PORT_LC} 端口:")
        add_group_layout.addWidget(MainWindow.lc_label, 4, 0)
        MainWindow.lc_entry = QLineEdit("2")
        MainWindow.lc_entry.setFont(chinese_font) # !! 使用局部变量 !!
        add_group_layout.addWidget(MainWindow.lc_entry, 4, 1)

        MainWindow.sfp_label = QLabel(f"{PORT_SFP}+ 端口:")
        MainWindow.sfp_entry = QLineEdit("8")
        MainWindow.sfp_entry.setFont(chinese_font) # !! 使用局部变量 !!
        add_group_layout.addWidget(MainWindow.sfp_label, 5, 0)
        add_group_layout.addWidget(MainWindow.sfp_entry, 5, 1)
        MainWindow.sfp_label.hide()
        MainWindow.sfp_entry.hide()

        MainWindow.add_button = QPushButton("添加设备")
        MainWindow.add_button.setFont(chinese_font) # !! 使用局部变量 !!
        add_group_layout.addWidget(MainWindow.add_button, 6, 0, 1, 2, alignment=Qt.AlignmentFlag.AlignCenter)
        left_layout.addWidget(add_group)

        # --- 设备列表组 ---
        list_group = QFrame()
        list_group.setObjectName("listGroup")
        list_group_layout = QVBoxLayout(list_group)
        list_group_layout.setContentsMargins(10, 15, 10, 10)

        filter_layout = QHBoxLayout()
        filter_label = QLabel("过滤:") # 创建实例
        filter_label.setFont(chinese_font) # !! 使用局部变量 !!
        filter_layout.addWidget(filter_label)
        MainWindow.device_filter_entry = QLineEdit()
        MainWindow.device_filter_entry.setFont(chinese_font) # !! 使用局部变量 !!
        MainWindow.device_filter_entry.setPlaceholderText("按名称或类型过滤...")
        filter_layout.addWidget(MainWindow.device_filter_entry)
        list_group_layout.addLayout(filter_layout)

        MainWindow.device_tablewidget = QTableWidget()
        MainWindow.device_tablewidget.setFont(chinese_font) # !! 使用局部变量 !!
        MainWindow.device_tablewidget.setColumnCount(6)
        MainWindow.device_tablewidget.setHorizontalHeaderLabels(["名称", "类型", PORT_MPO, PORT_LC, f"{PORT_SFP}+", "连接数(估)"])
        MainWindow.device_tablewidget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        MainWindow.device_tablewidget.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        MainWindow.device_tablewidget.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.SelectedClicked | QAbstractItemView.EditTrigger.EditKeyPressed)
        MainWindow.device_tablewidget.setSortingEnabled(True)
        header = MainWindow.device_tablewidget.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(COL_NAME, QHeaderView.ResizeMode.Stretch)
        MainWindow.device_tablewidget.setColumnWidth(COL_TYPE, 90)
        MainWindow.device_tablewidget.setColumnWidth(COL_MPO, 50)
        MainWindow.device_tablewidget.setColumnWidth(COL_LC, 50)
        MainWindow.device_tablewidget.setColumnWidth(COL_SFP, 50)
        MainWindow.device_tablewidget.setColumnWidth(COL_CONN, 80)
        list_group_layout.addWidget(MainWindow.device_tablewidget)

        device_op_layout = QHBoxLayout()
        MainWindow.remove_button = QPushButton("移除选中")
        MainWindow.remove_button.setFont(chinese_font) # !! 使用局部变量 !!
        MainWindow.clear_button = QPushButton("清空所有")
        MainWindow.clear_button.setFont(chinese_font) # !! 使用局部变量 !!
        device_op_layout.addWidget(MainWindow.remove_button)
        device_op_layout.addWidget(MainWindow.clear_button)
        list_group_layout.addLayout(device_op_layout)

        MainWindow.port_totals_label = QLabel("总计: MPO: 0, LC: 0, SFP+: 0")
        font = MainWindow.port_totals_label.font(); font.setBold(True)
        MainWindow.port_totals_label.setFont(font) # 使用 QLabel 自己的 font 对象
        MainWindow.port_totals_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        MainWindow.port_totals_label.setStyleSheet("padding-top: 5px; padding-right: 5px;")
        list_group_layout.addWidget(MainWindow.port_totals_label)
        left_layout.addWidget(list_group)

        # --- 文件操作组 ---
        file_group = QFrame()
        file_group.setObjectName("fileGroup")
        file_group_layout = QGridLayout(file_group)
        file_group_layout.setContentsMargins(10, 15, 10, 10)
        file_title = QLabel("<b>文件操作</b>")
        file_title.setFont(QFont(chinese_font.family(), 11)) # !! 使用局部变量 !!
        file_group_layout.addWidget(file_title, 0, 0, 1, 2, alignment=Qt.AlignmentFlag.AlignCenter)
        MainWindow.save_button = QPushButton("保存配置")
        MainWindow.save_button.setFont(chinese_font) # !! 使用局部变量 !!
        MainWindow.load_button = QPushButton("加载配置")
        MainWindow.load_button.setFont(chinese_font) # !! 使用局部变量 !!
        file_group_layout.addWidget(MainWindow.save_button, 1, 0)
        file_group_layout.addWidget(MainWindow.load_button, 1, 1)
        MainWindow.export_list_button = QPushButton("导出列表")
        MainWindow.export_list_button.setFont(chinese_font) # !! 使用局部变量 !!
        MainWindow.export_list_button.setEnabled(False)
        file_group_layout.addWidget(MainWindow.export_list_button, 2, 0)
        MainWindow.export_topo_button = QPushButton("导出拓扑图")
        MainWindow.export_topo_button.setFont(chinese_font) # !! 使用局部变量 !!
        MainWindow.export_topo_button.setEnabled(False)
        file_group_layout.addWidget(MainWindow.export_topo_button, 2, 1)
        MainWindow.export_report_button = QPushButton("导出报告 (HTML)")
        MainWindow.export_report_button.setFont(chinese_font) # !! 使用局部变量 !!
        MainWindow.export_report_button.setEnabled(False)
        file_group_layout.addWidget(MainWindow.export_report_button, 3, 0, 1, 2)
        left_layout.addWidget(file_group)

        # --- 跳过确认弹窗设置 ---
        suppress_frame = QFrame()
        suppress_frame.setFrameShape(QFrame.Shape.NoFrame)
        suppress_layout = QHBoxLayout(suppress_frame)
        suppress_layout.setContentsMargins(10, 0, 10, 5)
        MainWindow.suppress_confirm_checkbox = QCheckBox("跳过确认弹窗")
        MainWindow.suppress_confirm_checkbox.setFont(chinese_font) # !! 使用局部变量 !!
        suppress_layout.addWidget(MainWindow.suppress_confirm_checkbox)
        suppress_layout.addStretch()
        left_layout.addWidget(suppress_frame)

        left_layout.addStretch()

        # --- 右侧面板与 Tab 页 ---
        right_panel = QFrame()
        right_panel.setFrameShape(QFrame.Shape.StyledPanel)
        right_layout = QVBoxLayout(right_panel)
        main_splitter.addWidget(right_panel)

        # --- 计算控制栏 ---
        calculate_control_frame = QFrame()
        calculate_control_frame.setObjectName("calculateControlFrame")
        calculate_control_layout = QHBoxLayout(calculate_control_frame)
        calculate_control_layout.setContentsMargins(10, 5, 10, 5)
        calculate_label1 = QLabel("计算模式:") # 创建实例
        calculate_label1.setFont(chinese_font) # !! 使用局部变量 !!
        calculate_control_layout.addWidget(calculate_label1)
        MainWindow.topology_mode_combo = QComboBox()
        MainWindow.topology_mode_combo.setFont(chinese_font) # !! 使用局部变量 !!
        MainWindow.topology_mode_combo.addItems(["Mesh", "环形"])
        calculate_control_layout.addWidget(MainWindow.topology_mode_combo)
        calculate_label2 = QLabel("布局:") # 创建实例
        calculate_label2.setFont(chinese_font) # !! 使用局部变量 !!
        calculate_control_layout.addWidget(calculate_label2)
        MainWindow.layout_combo = QComboBox()
        MainWindow.layout_combo.setFont(chinese_font) # !! 使用局部变量 !!
        MainWindow.layout_combo.addItems(["Spring", "Circular", "Kamada-Kawai", "Random", "Shell"])
        calculate_control_layout.addWidget(MainWindow.layout_combo)
        MainWindow.calculate_button = QPushButton("计算连接")
        MainWindow.calculate_button.setFont(chinese_font) # !! 使用局部变量 !!
        calculate_control_layout.addWidget(MainWindow.calculate_button)
        MainWindow.fill_mesh_button = QPushButton("填充 (Mesh)")
        MainWindow.fill_mesh_button.setFont(chinese_font) # !! 使用局部变量 !!
        MainWindow.fill_mesh_button.setEnabled(False)
        calculate_control_layout.addWidget(MainWindow.fill_mesh_button)
        MainWindow.fill_ring_button = QPushButton("填充 (环形)")
        MainWindow.fill_ring_button.setFont(chinese_font) # !! 使用局部变量 !!
        MainWindow.fill_ring_button.setEnabled(False)
        calculate_control_layout.addWidget(MainWindow.fill_ring_button)
        calculate_control_layout.addStretch()
        right_layout.addWidget(calculate_control_frame)

        # --- Tab 控件 ---
        MainWindow.tab_widget = QTabWidget()
        MainWindow.tab_widget.setFont(chinese_font) # !! 使用局部变量 !!
        right_layout.addWidget(MainWindow.tab_widget)

        # --- 连接列表 Tab ---
        MainWindow.connections_tab = QWidget()
        connections_layout = QVBoxLayout(MainWindow.connections_tab)
        MainWindow.connections_textedit = QTextEdit()
        MainWindow.connections_textedit.setFont(chinese_font) # !! 使用局部变量 !!
        MainWindow.connections_textedit.setReadOnly(True)
        connections_layout.addWidget(MainWindow.connections_textedit)
        MainWindow.tab_widget.addTab(MainWindow.connections_tab, "连接列表")

        # --- 拓扑图 Tab ---
        MainWindow.topology_tab = QWidget()
        topology_layout = QVBoxLayout(MainWindow.topology_tab)
        if hasattr(MainWindow, 'mpl_canvas') and MainWindow.mpl_canvas:
             MainWindow.mpl_canvas.setParent(MainWindow.topology_tab)
             topology_layout.addWidget(MainWindow.mpl_canvas)
        else:
             print("错误: MainWindow 实例缺少 mpl_canvas 属性！")
             placeholder = QLabel("拓扑图画布加载失败")
             placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
             topology_layout.addWidget(placeholder)
        MainWindow.tab_widget.addTab(MainWindow.topology_tab, "拓扑图")

        # --- 手动编辑 Tab ---
        MainWindow.edit_tab = QWidget()
        edit_main_layout = QVBoxLayout(MainWindow.edit_tab)
        add_manual_group = QFrame(); add_manual_group.setObjectName("addManualGroup"); add_manual_group.setFrameShape(QFrame.Shape.StyledPanel); add_manual_layout = QGridLayout(add_manual_group); add_manual_layout.setContentsMargins(10,10,10,10);
        add_manual_title = QLabel("<b>添加手动连接</b>") # 创建实例
        add_manual_title.setFont(chinese_font) # !! 使用局部变量 !!
        add_manual_layout.addWidget(add_manual_title, 0, 0, 1, 4, alignment=Qt.AlignmentFlag.AlignCenter);
        add_manual_label1 = QLabel("设备 1:") # 创建实例
        add_manual_label1.setFont(chinese_font) # !! 使用局部变量 !!
        add_manual_layout.addWidget(add_manual_label1, 1, 0);
        MainWindow.edit_dev1_combo = QComboBox(); MainWindow.edit_dev1_combo.setFont(chinese_font); add_manual_layout.addWidget(MainWindow.edit_dev1_combo, 1, 1); # !! 使用局部变量 !!
        add_manual_label2 = QLabel("端口 1:") # 创建实例
        add_manual_label2.setFont(chinese_font) # !! 使用局部变量 !!
        add_manual_layout.addWidget(add_manual_label2, 1, 2);
        MainWindow.edit_port1_combo = QComboBox(); MainWindow.edit_port1_combo.setFont(chinese_font); add_manual_layout.addWidget(MainWindow.edit_port1_combo, 1, 3); # !! 使用局部变量 !!
        add_manual_label3 = QLabel("设备 2:") # 创建实例
        add_manual_label3.setFont(chinese_font) # !! 使用局部变量 !!
        add_manual_layout.addWidget(add_manual_label3, 2, 0);
        MainWindow.edit_dev2_combo = QComboBox(); MainWindow.edit_dev2_combo.setFont(chinese_font); add_manual_layout.addWidget(MainWindow.edit_dev2_combo, 2, 1); # !! 使用局部变量 !!
        add_manual_label4 = QLabel("端口 2:") # 创建实例
        add_manual_label4.setFont(chinese_font) # !! 使用局部变量 !!
        add_manual_layout.addWidget(add_manual_label4, 2, 2);
        MainWindow.edit_port2_combo = QComboBox(); MainWindow.edit_port2_combo.setFont(chinese_font); add_manual_layout.addWidget(MainWindow.edit_port2_combo, 2, 3); # !! 使用局部变量 !!
        MainWindow.add_manual_button = QPushButton("添加连接"); MainWindow.add_manual_button.setFont(chinese_font); add_manual_layout.addWidget(MainWindow.add_manual_button, 3, 0, 1, 4, alignment=Qt.AlignmentFlag.AlignCenter); edit_main_layout.addWidget(add_manual_group) # !! 使用局部变量 !!
        remove_manual_group = QFrame(); remove_manual_group.setObjectName("removeManualGroup"); remove_manual_group.setFrameShape(QFrame.Shape.StyledPanel); remove_manual_layout = QVBoxLayout(remove_manual_group); remove_manual_layout.setContentsMargins(10,10,10,10);
        remove_title = QLabel("<b>移除现有连接</b> (选中下方列表中的连接进行移除)"); remove_title.setFont(chinese_font); remove_manual_layout.addWidget(remove_title); # !! 使用局部变量 !!
        filter_conn_layout = QHBoxLayout();
        filter_conn_label1 = QLabel("类型过滤:") # 创建实例
        filter_conn_label1.setFont(chinese_font) # !! 使用局部变量 !!
        filter_conn_layout.addWidget(filter_conn_label1);
        MainWindow.conn_filter_type_combo = QComboBox(); MainWindow.conn_filter_type_combo.setFont(chinese_font); MainWindow.conn_filter_type_combo.addItems(["所有类型", "LC-LC (100G)", "MPO-MPO (25G)", "SFP-SFP (10G)", "MPO-SFP (10G)"]); filter_conn_layout.addWidget(MainWindow.conn_filter_type_combo); filter_conn_layout.addSpacing(15); # !! 使用局部变量 !!
        filter_conn_label2 = QLabel("设备过滤:") # 创建实例
        filter_conn_label2.setFont(chinese_font) # !! 使用局部变量 !!
        filter_conn_layout.addWidget(filter_conn_label2);
        MainWindow.conn_filter_device_entry = QLineEdit(); MainWindow.conn_filter_device_entry.setFont(chinese_font); MainWindow.conn_filter_device_entry.setPlaceholderText("按设备名称过滤..."); filter_conn_layout.addWidget(MainWindow.conn_filter_device_entry); remove_manual_layout.insertLayout(1, filter_conn_layout); # !! 使用局部变量 !!
        MainWindow.manual_connection_list = QListWidget(); MainWindow.manual_connection_list.setFont(chinese_font); MainWindow.manual_connection_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection); remove_manual_layout.addWidget(MainWindow.manual_connection_list); # !! 使用局部变量 !!
        MainWindow.remove_manual_button = QPushButton("移除选中连接"); MainWindow.remove_manual_button.setFont(chinese_font); MainWindow.remove_manual_button.setEnabled(False); remove_manual_layout.addWidget(MainWindow.remove_manual_button, alignment=Qt.AlignmentFlag.AlignCenter); edit_main_layout.addWidget(remove_manual_group) # !! 使用局部变量 !!
        MainWindow.tab_widget.addTab(MainWindow.edit_tab, "手动编辑")

        # --- 设置 Splitter 初始大小 ---
        main_splitter.setSizes([400, 700])
        main_splitter.setStretchFactor(1, 1)

