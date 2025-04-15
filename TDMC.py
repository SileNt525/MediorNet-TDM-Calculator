# -*- coding: utf-8 -*-
# MediorNet TDM 连接计算器 V36
# 主要变更:
# - 修正了 plot_topology 函数中因 current_node_ids 未提前赋值导致的 UnboundLocalError。
# - 保留 V35 的填充逻辑修正及其他功能。

import sys
import tkinter as tk
from tkinter import messagebox
import networkx as nx
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import itertools
import random
import platform
import copy
import json
from collections import defaultdict# -*- coding: utf-8 -*-
# MediorNet TDM 连接计算器 V37
# 主要变更:
# - 新增拓扑图节点拖动功能（基于 Matplotlib 事件处理，性能和流畅度可能有限）。
# - 保留 V36 的错误修复及其他功能。

import sys
import tkinter as tk
from tkinter import messagebox
import networkx as nx
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import itertools
import random
import platform
import copy
import json
from collections import defaultdict

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QTextEdit,
    QTabWidget, QFrame, QFileDialog, QMessageBox, QSpacerItem, QSizePolicy,
    QGridLayout, QListWidgetItem, QAbstractItemView, QTableWidget, QTableWidgetItem,
    QHeaderView, QListWidget, QSplitter
)
from PySide6.QtCore import Slot, Qt
from PySide6.QtGui import QFont

# --- QSS 样式定义 (与 V36 相同) ---
APP_STYLE = """
QMainWindow, QDialog, QMessageBox { /* 应用于主窗口、对话框和消息框 */
    background-color: #f0f0f0; /* 浅灰色背景 */
}
QFrame {
    background-color: transparent; /* 让 Frame 透明，除非特别指定 */
}
/* 为特定的分组 Frame 添加边框和圆角 */
QFrame#addDeviceGroup, QFrame#listGroup, QFrame#fileGroup,
QFrame#calculateControlFrame, QFrame#addManualGroup, QFrame#removeManualGroup {
    border: 1px solid #c8c8c8;
    border-radius: 5px;
    background-color: #f8f8f8; /* 轻微区分背景 */
    margin-bottom: 5px; /* 组之间增加一点间距 */
}
QPushButton {
    background-color: #e1e1e1;
    border: 1px solid #adadad;
    padding: 5px 10px;
    border-radius: 4px;
    min-height: 20px;
    min-width: 75px; /* 按钮最小宽度 */
}
QPushButton:hover {
    background-color: #cacaca;
    border: 1px solid #999999;
}
QPushButton:pressed {
    background-color: #b0b0b0;
    border: 1px solid #777777;
}
QPushButton:disabled {
    background-color: #d3d3d3;
    color: #a0a0a0;
    border: 1px solid #c0c0c0;
}
QLineEdit, QComboBox, QTextEdit, QListWidget, QTableWidget {
    background-color: white;
    border: 1px solid #c0c0c0;
    border-radius: 3px;
    padding: 3px;
    selection-background-color: #a8cce4; /* 选中项背景色 */
    selection-color: black; /* 选中项文字颜色 */
}
QComboBox::drop-down { /* 下拉箭头样式 */
    border: none;
    background: transparent;
    width: 15px; /* 增加箭头区域宽度 */
    padding-right: 5px;
}
QComboBox::down-arrow { /* 箭头图标 */
     width: 12px;
     height: 12px;
}
QTableWidget {
    alternate-background-color: #f8f8f8; /* 表格斑马纹 */
    gridline-color: #d0d0d0; /* 网格线颜色 */
}
QHeaderView::section { /* 表头样式 */
    background-color: #e8e8e8;
    padding: 4px;
    border: 1px solid #d0d0d0;
    border-left: none; /* 移除左边框避免双线 */
    font-weight: bold;
}
QHeaderView::section:first { /* 第一个表头单元格 */
     border-left: 1px solid #d0d0d0; /* 补回第一个左边框 */
}
QTabWidget::pane { /* 标签页主体框架 */
    border: 1px solid #c0c0c0;
    border-top: none; /* 顶部边框由标签栏提供 */
    background-color: white; /* 标签页内容区域背景 */
}
QTabBar::tab { /* 标签样式 */
    background: #e1e1e1;
    border: 1px solid #adadad;
    border-bottom: none;
    padding: 6px 12px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    margin-right: 2px; /* 标签间距 */
}
QTabBar::tab:selected {
    background: white; /* 选中标签页背景与内容区域一致 */
    margin-bottom: -1px; /* 轻微上移，与边框融合 */
    border-bottom: 1px solid white; /* 覆盖 pane 的顶部边框 */
}
QTabBar::tab:!selected:hover {
    background: #cacaca;
}
QSplitter::handle { /* 分隔条样式 */
    background-color: #d0d0d0;
    border: none;
}
QSplitter::handle:horizontal {
    width: 3px; /* 水平分隔条宽度 */
}
QSplitter::handle:vertical {
    height: 3px; /* 垂直分隔条高度 */
}
QSplitter::handle:hover {
    background-color: #b0b0b0;
}
"""

# --- 用于数字排序的 QTableWidgetItem 子类 (与 V36 相同) ---
class NumericTableWidgetItem(QTableWidgetItem):
    """自定义 QTableWidgetItem 以支持数字排序。"""
    def __lt__(self, other):
        data_self = self.data(Qt.ItemDataRole.UserRole + 1)
        data_other = other.data(Qt.ItemDataRole.UserRole + 1)
        try:
            num_self = float(data_self)
            num_other = float(data_other)
            return num_self < num_other
        except (TypeError, ValueError):
            return super().__lt__(other)

# --- 数据结构 (与 V36 相同) ---
class Device:
    """代表一个 MediorNet 设备"""
    def __init__(self, id, name, type, mpo_ports=0, lc_ports=0, sfp_ports=0):
        self.id = id
        self.name = name
        self.type = type
        self.mpo_total = mpo_ports
        self.lc_total = lc_ports
        self.sfp_total = sfp_ports
        self.reset_ports()

    def reset_ports(self):
        self.connections = 0.0
        self.port_connections = {}

    def get_all_possible_ports(self):
        ports = []
        ports.extend([f"LC{i+1}" for i in range(self.lc_total)])
        ports.extend([f"SFP{i+1}" for i in range(self.sfp_total)])
        for i in range(self.mpo_total):
            base = f"MPO{i+1}"
            ports.extend([f"{base}-Ch{j+1}" for j in range(4)])
        return ports

    def get_all_available_ports(self):
        all_ports = self.get_all_possible_ports()
        used_ports = set(self.port_connections.keys())
        available = [p for p in all_ports if p not in used_ports]
        return available

    def use_specific_port(self, port_name, target_device_name):
        if port_name in self.get_all_possible_ports() and port_name not in self.port_connections:
            self.port_connections[port_name] = target_device_name
            if port_name.startswith("MPO"): self.connections += 0.25
            else: self.connections += 1
            return True
        return False

    def return_port(self, port_name):
        port_in_use_record = port_name in self.port_connections
        port_already_available = False
        port_type_valid = True
        if port_name.startswith("LC") or port_name.startswith("SFP") or (port_name.startswith("MPO") and "-Ch" in port_name):
             port_already_available = port_name in self.get_all_available_ports()
        else:
            print(f"警告: 尝试归还未知类型的端口 {port_name}")
            port_type_valid = False
        if not port_type_valid: return
        if port_in_use_record:
            target = self.port_connections.pop(port_name)
            if not port_already_available:
                if port_name.startswith("MPO"): self.connections -= 0.25
                elif port_name.startswith("LC") or port_name.startswith("SFP"): self.connections -= 1
                self.connections = max(0.0, self.connections)
        else:
             pass

    def get_available_port(self, port_type, target_device_name):
        possible_ports = []
        if port_type == 'LC': possible_ports = [f"LC{i+1}" for i in range(self.lc_total)]
        elif port_type == 'MPO':
            for i in range(self.mpo_total): base = f"MPO{i+1}"; possible_ports.extend([f"{base}-Ch{j+1}" for j in range(4)])
            random.shuffle(possible_ports)
        elif port_type == 'SFP': possible_ports = [f"SFP{i+1}" for i in range(self.sfp_total)]
        used_ports = set(self.port_connections.keys())
        for port in possible_ports:
            if port not in used_ports:
                self.port_connections[port] = target_device_name
                if port.startswith("MPO"): self.connections += 0.25
                else: self.connections += 1
                return port
        return None

    def get_specific_available_port(self, port_type_prefix):
        possible_ports = []
        if port_type_prefix == 'LC': possible_ports = sorted([f"LC{i+1}" for i in range(self.lc_total)], key=lambda x: int(x[2:]))
        elif port_type_prefix == 'MPO':
            for i in range(self.mpo_total): base = f"MPO{i+1}"; possible_ports.extend(sorted([f"{base}-Ch{j+1}" for j in range(4)], key=lambda x: int(x.split('-Ch')[-1])))
        elif port_type_prefix == 'SFP': possible_ports = sorted([f"SFP{i+1}" for i in range(self.sfp_total)], key=lambda x: int(x[3:]))
        used_ports = set(self.port_connections.keys())
        for port in possible_ports:
            if port not in used_ports: return port
        return None

    def to_dict(self):
        return {'id': self.id, 'name': self.name, 'type': self.type,
                'mpo_ports': self.mpo_total, 'lc_ports': self.lc_total, 'sfp_ports': self.sfp_total}

    @classmethod
    def from_dict(cls, data):
        data.setdefault('id', random.randint(10000, 99999))
        return cls(id=data['id'], name=data['name'], type=data['type'],
                   mpo_ports=data.get('mpo_ports', 0), lc_ports=data.get('lc_ports', 0), sfp_ports=data.get('sfp_ports', 0))

    def __repr__(self):
        return f"{self.name} ({self.type})"

# --- 连接计算逻辑 (与 V36 相同) ---

# 辅助函数
def _find_best_single_link(dev1_copy, dev2_copy):
    """辅助函数：查找两个设备副本之间最高优先级的单个可用连接"""
    if dev1_copy.type in ['MicroN UHD', 'HorizoN'] and dev2_copy.type in ['MicroN UHD', 'HorizoN']:
        port1 = dev1_copy.get_specific_available_port('LC')
        if port1:
            port2 = dev2_copy.get_specific_available_port('LC')
            if port2:
                dev1_copy.use_specific_port(port1, dev2_copy.name); dev2_copy.use_specific_port(port2, dev1_copy.name)
                return port1, port2, 'LC-LC (100G)'
    if dev1_copy.type in ['MicroN UHD', 'HorizoN'] and dev2_copy.type in ['MicroN UHD', 'HorizoN']:
        port1 = dev1_copy.get_specific_available_port('MPO')
        if port1:
            port2 = dev2_copy.get_specific_available_port('MPO')
            if port2:
                dev1_copy.use_specific_port(port1, dev2_copy.name); dev2_copy.use_specific_port(port2, dev1_copy.name)
                return port1, port2, 'MPO-MPO (25G)'
    if dev1_copy.type == 'MicroN' and dev2_copy.type == 'MicroN':
        port1 = dev1_copy.get_specific_available_port('SFP')
        if port1:
            port2 = dev2_copy.get_specific_available_port('SFP')
            if port2:
                dev1_copy.use_specific_port(port1, dev2_copy.name); dev2_copy.use_specific_port(port2, dev1_copy.name)
                return port1, port2, 'SFP-SFP (10G)'
    uhd_dev, micron_dev = (None, None); original_dev1 = dev1_copy
    if dev1_copy.type in ['MicroN UHD', 'HorizoN'] and dev2_copy.type == 'MicroN': uhd_dev, micron_dev = dev1_copy, dev2_copy
    elif dev2_copy.type in ['MicroN UHD', 'HorizoN'] and dev1_copy.type == 'MicroN': uhd_dev, micron_dev = dev2_copy, dev1_copy
    if uhd_dev and micron_dev:
        port_uhd = uhd_dev.get_specific_available_port('MPO')
        if port_uhd:
            port_micron = micron_dev.get_specific_available_port('SFP')
            if port_micron:
                uhd_dev.use_specific_port(port_uhd, micron_dev.name); micron_dev.use_specific_port(port_micron, uhd_dev.name)
                if original_dev1 == uhd_dev: return port_uhd, port_micron, 'MPO-SFP (10G)'
                else: return port_micron, port_uhd, 'MPO-SFP (10G)'
    return None, None, None

# Mesh 计算函数
def calculate_mesh_connections(devices):
    """计算 Mesh 连接 - V34 改进版"""
    if len(devices) < 2: return [], {}
    connections = []; temp_devices = [copy.deepcopy(dev) for dev in devices]
    for d in temp_devices: d.reset_ports()
    device_map = {dev.id: dev for dev in temp_devices}
    all_pairs_ids = list(itertools.combinations([d.id for d in temp_devices], 2))
    connected_once_pairs = set()
    print("Phase 1: 尝试为每个设备对建立第一条连接...")
    made_progress_phase1 = True
    while made_progress_phase1:
        made_progress_phase1 = False
        for dev1_id, dev2_id in all_pairs_ids:
            pair_key = tuple(sorted((dev1_id, dev2_id)))
            if pair_key not in connected_once_pairs:
                dev1_copy = device_map[dev1_id]; dev2_copy = device_map[dev2_id]
                original_dev1 = next(d for d in devices if d.id == dev1_id); original_dev2 = next(d for d in devices if d.id == dev2_id)
                port1, port2, conn_type = _find_best_single_link(dev1_copy, dev2_copy)
                if port1 and port2:
                    if dev1_copy.id == dev1_id: connections.append((original_dev1, port1, original_dev2, port2, conn_type))
                    else: connections.append((original_dev2, port2, original_dev1, port1, conn_type))
                    connected_once_pairs.add(pair_key); made_progress_phase1 = True
        if not made_progress_phase1: break
    print(f"Phase 1 完成. 建立了 {len(connections)} 条初始连接。")
    print("Phase 2: 填充剩余端口...")
    connection_made_in_full_pass_phase2 = True
    while connection_made_in_full_pass_phase2:
        connection_made_in_full_pass_phase2 = False
        random.shuffle(all_pairs_ids)
        for dev1_id, dev2_id in all_pairs_ids:
            dev1_copy = device_map[dev1_id]; dev2_copy = device_map[dev2_id]
            original_dev1 = next(d for d in devices if d.id == dev1_id); original_dev2 = next(d for d in devices if d.id == dev2_id)
            port1, port2, conn_type = _find_best_single_link(dev1_copy, dev2_copy)
            if port1 and port2:
                if dev1_copy.id == dev1_id: connections.append((original_dev1, port1, original_dev2, port2, conn_type))
                else: connections.append((original_dev2, port2, original_dev1, port1, conn_type))
                connection_made_in_full_pass_phase2 = True
    print(f"Phase 2 完成. 总连接数: {len(connections)}")
    return connections, device_map

# 环形计算函数
def calculate_ring_connections(devices):
    if len(devices) < 2: return [], {}
    if len(devices) == 2: return calculate_mesh_connections(devices)
    connections = []; temp_devices = [copy.deepcopy(dev) for dev in devices]
    for d in temp_devices: d.reset_ports()
    device_map = {dev.id: dev for dev in temp_devices}
    sorted_dev_ids = sorted([d.id for d in devices]); num_devices = len(sorted_dev_ids)
    link_established = [False] * num_devices
    for i in range(num_devices):
        dev1_id = sorted_dev_ids[i]; dev2_id = sorted_dev_ids[(i + 1) % num_devices]
        dev1_copy = device_map[dev1_id]; dev2_copy = device_map[dev2_id]
        original_dev1 = next(d for d in devices if d.id == dev1_id); original_dev2 = next(d for d in devices if d.id == dev2_id)
        port1, port2, conn_type = _find_best_single_link(dev1_copy, dev2_copy)
        if port1 and port2:
            if dev1_copy.id == dev1_id: connections.append((original_dev1, port1, original_dev2, port2, conn_type))
            else: connections.append((original_dev2, port2, original_dev1, port1, conn_type))
            link_established[i] = True
        else: print(f"警告: 无法在 {original_dev1.name} 和 {original_dev2.name} 之间建立环形连接段。")
    if not all(link_established): print("警告: 未能完成完整的环形连接。")
    return connections, device_map

# 内部 Mesh 填充辅助函数
def _fill_connections_mesh_style(devices_current_state):
    """辅助函数：使用 Mesh 逻辑填充给定设备状态下的剩余连接。"""
    if len(devices_current_state) < 2: return []
    new_connections = []
    temp_devices = [copy.deepcopy(dev) for dev in devices_current_state]
    device_map = {dev.id: dev for dev in temp_devices}
    all_pairs_ids = list(itertools.combinations([d.id for d in temp_devices], 2))
    connection_made_in_full_pass = True
    while connection_made_in_full_pass:
        connection_made_in_full_pass = False
        random.shuffle(all_pairs_ids)
        for dev1_id, dev2_id in all_pairs_ids:
            dev1_copy = device_map[dev1_id]; dev2_copy = device_map[dev2_id]
            port1, port2, conn_type = _find_best_single_link(dev1_copy, dev2_copy)
            if port1 and port2:
                original_dev1 = next(d for d in devices_current_state if d.id == dev1_id)
                original_dev2 = next(d for d in devices_current_state if d.id == dev2_id)
                if dev1_copy.id == dev1_id: new_connections.append((original_dev1, port1, original_dev2, port2, conn_type))
                else: new_connections.append((original_dev2, port2, original_dev1, port1, conn_type))
                connection_made_in_full_pass = True
    return new_connections

# 内部 Ring 填充辅助函数
def _fill_connections_ring_style(devices_current_state):
    """辅助函数：使用 Ring 逻辑填充给定设备状态下的剩余连接。"""
    if len(devices_current_state) < 2: return []
    new_connections = []
    temp_devices = [copy.deepcopy(dev) for dev in devices_current_state]
    device_map = {dev.id: dev for dev in temp_devices}
    sorted_dev_ids = sorted([d.id for d in temp_devices]); num_devices = len(sorted_dev_ids)
    connection_made_in_full_pass = True
    while connection_made_in_full_pass:
        connection_made_in_full_pass = False
        for i in range(num_devices):
            dev1_id = sorted_dev_ids[i]; dev2_id = sorted_dev_ids[(i + 1) % num_devices]
            dev1_copy = device_map[dev1_id]; dev2_copy = device_map[dev2_id]
            original_dev1 = next(d for d in devices_current_state if d.id == dev1_id)
            original_dev2 = next(d for d in devices_current_state if d.id == dev2_id)
            port1, port2, conn_type = _find_best_single_link(dev1_copy, dev2_copy)
            if port1 and port2:
                if dev1_copy.id == dev1_id: new_connections.append((original_dev1, port1, original_dev2, port2, conn_type))
                else: new_connections.append((original_dev2, port2, original_dev1, port1, conn_type))
                connection_made_in_full_pass = True
    return new_connections
# --- 结束连接计算逻辑 ---


# --- Matplotlib Canvas Widget (与 V36 相同) ---
class MplCanvas(FigureCanvas):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super().__init__(self.fig)
        self.setParent(parent)
        FigureCanvas.setSizePolicy(self, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        FigureCanvas.updateGeometry(self)

    def plot_topology(self, devices, connections, layout_algorithm='spring', fixed_pos=None, selected_node_id=None):
        self.axes.cla()
        if not devices: self.axes.text(0.5, 0.5, '无设备数据', ha='center', va='center'); self.draw(); return None, None
        chinese_font_name = find_chinese_font(); font_prop = None; current_font_family = 'sans-serif'
        if chinese_font_name:
            try: plt.rcParams['font.sans-serif'] = [chinese_font_name] + plt.rcParams.get('font.sans-serif', []); font_prop = font_manager.FontProperties(family=chinese_font_name); current_font_family = chinese_font_name; plt.rcParams['axes.unicode_minus'] = False
            except Exception as e: print(f"警告: 设置 Matplotlib 字体 '{chinese_font_name}' 失败: {e}"); chinese_font_name = None
        if not chinese_font_name: plt.rcParams['font.sans-serif'] = ['sans-serif']; plt.rcParams['axes.unicode_minus'] = False; font_prop = font_manager.FontProperties(); current_font_family = 'sans-serif'
        G = nx.Graph(); node_ids = [dev.id for dev in devices]; node_colors = []; node_labels = {}; node_alphas = []
        highlight_color = 'yellow'; default_alpha = 0.9; dimmed_alpha = 0.3
        for dev in devices:
            G.add_node(dev.id); node_labels[dev.id] = f"{dev.name}\n({dev.type})"
            base_color = 'grey';
            if dev.type == 'MicroN UHD': base_color = 'skyblue'
            elif dev.type == 'HorizoN': base_color = 'lightcoral'
            elif dev.type == 'MicroN': base_color = 'lightgreen'
            if selected_node_id is not None:
                if dev.id == selected_node_id: node_colors.append(highlight_color); node_alphas.append(default_alpha)
                else: is_neighbor = any((conn[0].id == selected_node_id and conn[2].id == dev.id) or (conn[0].id == dev.id and conn[2].id == selected_node_id) for conn in connections); node_colors.append(base_color); node_alphas.append(default_alpha if is_neighbor else dimmed_alpha)
            else: node_colors.append(base_color); node_alphas.append(default_alpha)
        edge_labels, edge_counts = {}, {}; highlighted_edges = set()
        if connections:
            for conn in connections:
                dev1, _, dev2, _, conn_type = conn
                if dev1.id in node_ids and dev2.id in node_ids:
                    edge_key = tuple(sorted((dev1.id, dev2.id))); G.add_edge(dev1.id, dev2.id)
                    if edge_key not in edge_counts: edge_counts[edge_key] = {}
                    base_conn_type = conn_type.split(' ')[0]
                    if base_conn_type not in edge_counts[edge_key]: edge_counts[edge_key][base_conn_type] = {'count': 0, 'details': conn_type}
                    edge_counts[edge_key][base_conn_type]['count'] += 1
                    if selected_node_id is not None and (dev1.id == selected_node_id or dev2.id == selected_node_id): highlighted_edges.add(edge_key)
            for edge_key, type_groups in edge_counts.items(): label_parts = [f"{data['details']} x{data['count']}" for base_type, data in type_groups.items()]; edge_labels[edge_key] = "\n".join(label_parts)
        pos = None
        if not G: print("DIAG (Plot): 图为空，不计算布局。"); self.axes.text(0.5, 0.5, '无连接数据', ha='center', va='center', fontproperties=font_prop); self.draw(); return self.fig, None
        if fixed_pos:
            current_node_ids = set(G.nodes())
            stored_node_ids = set(fixed_pos.keys())
            if current_node_ids == stored_node_ids: pos = fixed_pos
            else: print("DIAG (Plot): 节点已更改，重新计算布局。"); fixed_pos = None
        if not pos:
            try:
                if layout_algorithm == 'circular': pos = nx.circular_layout(G)
                elif layout_algorithm == 'kamada-kawai': pos = nx.kamada_kawai_layout(G)
                elif layout_algorithm == 'random': pos = nx.random_layout(G, seed=42)
                elif layout_algorithm == 'shell':
                    shells = []; types_present = sorted(list(set(dev.type for dev in devices)))
                    for t in types_present: shells.append([dev.id for dev in devices if dev.type == t])
                    if len(shells) < 2: pos = nx.spring_layout(G, seed=42, k=0.8)
                    else: pos = nx.shell_layout(G, nlist=shells)
                else: pos = nx.spring_layout(G, seed=42, k=0.8)
            except Exception as e: print(f"警告: 计算布局 '{layout_algorithm}' 时出错: {e}. 使用 spring 布局回退。"); pos = nx.spring_layout(G, seed=42, k=0.8)
        nx.draw_networkx_nodes(G, pos, node_color=node_colors, node_size=3500, ax=self.axes, alpha=node_alphas)
        nx.draw_networkx_labels(G, pos, labels=node_labels, font_size=9, ax=self.axes, font_family=current_font_family)
        if connections and G.edges():
            unique_edges = list(G.edges()); edge_colors = []; edge_widths = []; edge_alphas = []
            highlight_edge_width = 2.5; default_edge_width = 1.5; dimmed_edge_alpha = 0.15; default_edge_alpha = 0.7
            for u, v in unique_edges:
                edge_key = tuple(sorted((u, v))); is_highlighted = edge_key in highlighted_edges; is_selected_node_present = selected_node_id is not None
                color_found = 'black'
                if edge_key in edge_counts:
                    first_base_type = next(iter(edge_counts[edge_key]))
                    if 'LC-LC' in first_base_type: color_found = 'blue'
                    elif 'MPO-MPO' in first_base_type: color_found = 'red'
                    elif 'MPO-SFP' in first_base_type: color_found = 'orange'
                    elif 'SFP-SFP' in first_base_type: color_found = 'purple'
                edge_colors.append(color_found); edge_widths.append(highlight_edge_width if is_highlighted else default_edge_width)
                edge_alphas.append(default_edge_alpha if (not is_selected_node_present or is_highlighted) else dimmed_edge_alpha)
            nx.draw_networkx_edges(G, pos, edgelist=unique_edges, edge_color=edge_colors, width=edge_widths, alpha=edge_alphas, ax=self.axes, arrows=False)
            edge_label_colors = {}; dimmed_label_color = 'lightgrey'; default_label_color = 'black'
            for edge, label in edge_labels.items():
                 edge_key = tuple(sorted(edge)); is_highlighted = edge_key in highlighted_edges; is_selected_node_present = selected_node_id is not None
                 color = default_label_color if (not is_selected_node_present or is_highlighted) else dimmed_label_color
                 edge_label_colors[edge_key] = color
            nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, font_size=7, ax=self.axes, label_pos=0.5, rotate=False, font_family=current_font_family, font_color=default_label_color)
        self.axes.set_title("网络连接拓扑图", fontproperties=font_prop); self.axes.axis('off')
        legend_elements = [
            plt.Line2D([0], [0], marker='o', color='w', label='MicroN UHD', markerfacecolor='skyblue', markersize=10),
            plt.Line2D([0], [0], marker='o', color='w', label='HorizoN', markerfacecolor='lightcoral', markersize=10),
            plt.Line2D([0], [0], marker='o', color='w', label='MicroN', markerfacecolor='lightgreen', markersize=10),
            plt.Line2D([0], [0], color='blue', lw=2, label='LC-LC (100G)'), plt.Line2D([0], [0], color='red', lw=2, label='MPO-MPO (25G)'),
            plt.Line2D([0], [0], color='orange', lw=2, label='MPO-SFP (10G)'), plt.Line2D([0], [0], color='purple', lw=2, label='SFP-SFP (10G)')
        ]
        legend_prop_small = copy.copy(font_prop); legend_prop_small.set_size('small')
        self.axes.legend(handles=legend_elements, loc='best', prop=legend_prop_small)
        self.fig.tight_layout(); self.draw_idle()
        return self.fig, pos
# --- 结束 Matplotlib Canvas ---

# --- 辅助函数 (查找字体，与 V35 相同) ---
def find_chinese_font():
    font_names = ['SimHei', 'Microsoft YaHei', 'PingFang SC', 'Heiti SC', 'STHeiti', 'Noto Sans CJK SC', 'WenQuanYi Micro Hei', 'sans-serif']
    font_paths = font_manager.findSystemFonts(fontpaths=None, fontext='ttf')
    found_fonts = {}
    for font_path in font_paths:
        try: prop = font_manager.FontProperties(fname=font_path); font_name = prop.get_name()
        except RuntimeError: continue
        except Exception: continue
        if font_name in font_names and font_name not in found_fonts: found_fonts[font_name] = font_path
    for name in ['PingFang SC', 'Microsoft YaHei', 'SimHei']:
         if name in found_fonts: return name
    for name in font_names:
        if name in found_fonts: return name
    print("警告: 未找到特定中文字体。")
    return None

# --- 主窗口 ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MediorNet TDM 连接计算器 V37 (PySide6)") # <--- 版本号更新
        self.setGeometry(100, 100, 1100, 800)
        self.devices = []
        self.connections_result = []
        self.fig = None
        self.node_positions = None
        self.device_id_counter = 0
        self.selected_node_id = None
        self.dragged_node_id = None # <--- 新增：跟踪拖动状态
        self.drag_offset = (0, 0)   # <--- 新增：跟踪拖动偏移
        # 设置字体
        font_families = []; os_system = platform.system()
        if os_system == "Windows": font_families.extend(["Microsoft YaHei", "SimHei"])
        elif os_system == "Darwin": font_families.append("PingFang SC")
        font_families.extend(["Noto Sans CJK SC", "WenQuanYi Micro Hei", "sans-serif"])
        self.chinese_font = QFont(); self.chinese_font.setFamilies(font_families); self.chinese_font.setPointSize(10)
        # --- 主布局与控件 (与 V36 相同) ---
        main_widget = QWidget(); self.setCentralWidget(main_widget); main_layout = QHBoxLayout(main_widget)
        main_splitter = QSplitter(Qt.Orientation.Horizontal); main_layout.addWidget(main_splitter)
        left_panel = QFrame(); left_panel.setFrameShape(QFrame.Shape.StyledPanel);
        left_layout = QVBoxLayout(left_panel); main_splitter.addWidget(left_panel)
        add_group = QFrame(); add_group.setObjectName("addDeviceGroup")
        add_group_layout = QGridLayout(add_group); add_group_layout.setContentsMargins(10, 15, 10, 10); add_group_layout.setVerticalSpacing(8)
        add_title = QLabel("<b>添加新设备</b>"); add_title.setFont(QFont(self.chinese_font.family(), 11)); add_group_layout.addWidget(add_title, 0, 0, 1, 2, alignment=Qt.AlignmentFlag.AlignCenter)
        add_group_layout.addWidget(QLabel("类型:"), 1, 0); self.device_type_combo = QComboBox(); self.device_type_combo.addItems(['MicroN UHD', 'HorizoN', 'MicroN']); self.device_type_combo.setFont(self.chinese_font); self.device_type_combo.currentIndexChanged.connect(self.update_port_entries); add_group_layout.addWidget(self.device_type_combo, 1, 1)
        add_group_layout.addWidget(QLabel("名称:"), 2, 0); self.device_name_entry = QLineEdit(); self.device_name_entry.setFont(self.chinese_font); add_group_layout.addWidget(self.device_name_entry, 2, 1)
        self.mpo_label = QLabel("MPO 端口:"); add_group_layout.addWidget(self.mpo_label, 3, 0); self.mpo_entry = QLineEdit("4"); self.mpo_entry.setFont(self.chinese_font); add_group_layout.addWidget(self.mpo_entry, 3, 1)
        self.lc_label = QLabel("LC 端口:"); add_group_layout.addWidget(self.lc_label, 4, 0); self.lc_entry = QLineEdit("2"); self.lc_entry.setFont(self.chinese_font); add_group_layout.addWidget(self.lc_entry, 4, 1)
        self.sfp_label = QLabel("SFP+ 端口:"); self.sfp_entry = QLineEdit("8"); self.sfp_entry.setFont(self.chinese_font); add_group_layout.addWidget(self.sfp_label, 5, 0); add_group_layout.addWidget(self.sfp_entry, 5, 1); self.sfp_label.hide(); self.sfp_entry.hide()
        self.add_button = QPushButton("添加设备"); self.add_button.setFont(self.chinese_font); self.add_button.clicked.connect(self.add_device); add_group_layout.addWidget(self.add_button, 6, 0, 1, 2, alignment=Qt.AlignmentFlag.AlignCenter)
        left_layout.addWidget(add_group)
        list_group = QFrame(); list_group.setObjectName("listGroup"); list_group_layout = QVBoxLayout(list_group); list_group_layout.setContentsMargins(10, 15, 10, 10)
        filter_layout = QHBoxLayout(); filter_layout.addWidget(QLabel("过滤:", font=self.chinese_font)); self.device_filter_entry = QLineEdit(); self.device_filter_entry.setFont(self.chinese_font); self.device_filter_entry.setPlaceholderText("按名称或类型过滤..."); self.device_filter_entry.textChanged.connect(self.filter_device_table); filter_layout.addWidget(self.device_filter_entry); list_group_layout.addLayout(filter_layout)
        self.device_tablewidget = QTableWidget(); self.device_tablewidget.setFont(self.chinese_font); self.device_tablewidget.setColumnCount(6); self.device_tablewidget.setHorizontalHeaderLabels(["名称", "类型", "MPO", "LC", "SFP+", "连接数(估)"]); self.device_tablewidget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); self.device_tablewidget.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection); self.device_tablewidget.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers);
        self.device_tablewidget.setSortingEnabled(True)
        header = self.device_tablewidget.horizontalHeader(); header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive); header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.device_tablewidget.setColumnWidth(1, 90); self.device_tablewidget.setColumnWidth(2, 50); self.device_tablewidget.setColumnWidth(3, 50); self.device_tablewidget.setColumnWidth(4, 50); self.device_tablewidget.setColumnWidth(5, 80)
        self.device_tablewidget.itemDoubleClicked.connect(self.show_device_details_from_table); list_group_layout.addWidget(self.device_tablewidget)
        device_op_layout = QHBoxLayout(); self.remove_button = QPushButton("移除选中"); self.remove_button.setFont(self.chinese_font); self.remove_button.clicked.connect(self.remove_device); self.clear_button = QPushButton("清空所有"); self.clear_button.setFont(self.chinese_font); self.clear_button.clicked.connect(self.clear_all_devices); device_op_layout.addWidget(self.remove_button); device_op_layout.addWidget(self.clear_button); list_group_layout.addLayout(device_op_layout)
        left_layout.addWidget(list_group)
        file_group = QFrame(); file_group.setObjectName("fileGroup"); file_group_layout = QGridLayout(file_group); file_group_layout.setContentsMargins(10, 15, 10, 10); file_title = QLabel("<b>文件操作</b>"); file_title.setFont(QFont(self.chinese_font.family(), 11)); file_group_layout.addWidget(file_title, 0, 0, 1, 2, alignment=Qt.AlignmentFlag.AlignCenter)
        self.save_button = QPushButton("保存配置"); self.save_button.setFont(self.chinese_font); self.save_button.clicked.connect(self.save_config); self.load_button = QPushButton("加载配置"); self.load_button.setFont(self.chinese_font); self.load_button.clicked.connect(self.load_config); file_group_layout.addWidget(self.save_button, 1, 0); file_group_layout.addWidget(self.load_button, 1, 1)
        self.export_list_button = QPushButton("导出列表"); self.export_list_button.setFont(self.chinese_font); self.export_list_button.clicked.connect(self.export_connections); self.export_list_button.setEnabled(False); self.export_topo_button = QPushButton("导出拓扑图"); self.export_topo_button.setFont(self.chinese_font); self.export_topo_button.clicked.connect(self.export_topology); self.export_topo_button.setEnabled(False); file_group_layout.addWidget(self.export_list_button, 2, 0); file_group_layout.addWidget(self.export_topo_button, 2, 1)
        left_layout.addWidget(file_group); left_layout.addStretch()
        right_panel = QFrame(); right_panel.setFrameShape(QFrame.Shape.StyledPanel)
        right_layout = QVBoxLayout(right_panel); main_splitter.addWidget(right_panel)
        calculate_control_frame = QFrame(); calculate_control_frame.setObjectName("calculateControlFrame")
        calculate_control_layout = QHBoxLayout(calculate_control_frame)
        calculate_control_layout.setContentsMargins(10, 5, 10, 5)
        calculate_control_layout.addWidget(QLabel("计算模式:", font=self.chinese_font))
        self.topology_mode_combo = QComboBox(); self.topology_mode_combo.setFont(self.chinese_font); self.topology_mode_combo.addItems(["Mesh", "环形"]); calculate_control_layout.addWidget(self.topology_mode_combo)
        calculate_control_layout.addWidget(QLabel("布局:", font=self.chinese_font))
        self.layout_combo = QComboBox(); self.layout_combo.setFont(self.chinese_font); self.layout_combo.addItems(["Spring", "Circular", "Kamada-Kawai", "Random", "Shell"]); self.layout_combo.currentIndexChanged.connect(self.on_layout_change); calculate_control_layout.addWidget(self.layout_combo)
        self.calculate_button = QPushButton("计算连接"); self.calculate_button.setFont(self.chinese_font); self.calculate_button.clicked.connect(self.calculate_and_display); calculate_control_layout.addWidget(self.calculate_button)
        self.fill_mesh_button = QPushButton("填充 (Mesh)"); self.fill_mesh_button.setFont(self.chinese_font); self.fill_mesh_button.setEnabled(False); self.fill_mesh_button.clicked.connect(self.fill_remaining_mesh); calculate_control_layout.addWidget(self.fill_mesh_button)
        self.fill_ring_button = QPushButton("填充 (环形)"); self.fill_ring_button.setFont(self.chinese_font); self.fill_ring_button.setEnabled(False); self.fill_ring_button.clicked.connect(self.fill_remaining_ring); calculate_control_layout.addWidget(self.fill_ring_button)
        calculate_control_layout.addStretch(); right_layout.addWidget(calculate_control_frame)
        self.tab_widget = QTabWidget(); self.tab_widget.setFont(self.chinese_font); right_layout.addWidget(self.tab_widget)
        self.connections_tab = QWidget(); connections_layout = QVBoxLayout(self.connections_tab); self.connections_textedit = QTextEdit(); self.connections_textedit.setFont(self.chinese_font); self.connections_textedit.setReadOnly(True); connections_layout.addWidget(self.connections_textedit); self.tab_widget.addTab(self.connections_tab, "连接列表")
        self.topology_tab = QWidget(); topology_layout = QVBoxLayout(self.topology_tab); self.mpl_canvas = MplCanvas(self.topology_tab, width=8, height=6, dpi=100); topology_layout.addWidget(self.mpl_canvas); self.tab_widget.addTab(self.topology_tab, "拓扑图")
        # --- 连接 Matplotlib 事件 (包括拖动所需的) ---
        self.mpl_canvas.mpl_connect('button_press_event', self.on_canvas_press)
        self.mpl_canvas.mpl_connect('motion_notify_event', self.on_canvas_motion)
        self.mpl_canvas.mpl_connect('button_release_event', self.on_canvas_release)
        # --- 结束连接 ---
        self.edit_tab = QWidget(); edit_main_layout = QVBoxLayout(self.edit_tab)
        add_manual_group = QFrame(); add_manual_group.setObjectName("addManualGroup"); add_manual_group.setFrameShape(QFrame.Shape.StyledPanel); add_manual_layout = QGridLayout(add_manual_group); add_manual_layout.setContentsMargins(10,10,10,10); add_manual_layout.addWidget(QLabel("<b>添加手动连接</b>", font=self.chinese_font), 0, 0, 1, 4, alignment=Qt.AlignmentFlag.AlignCenter)
        add_manual_layout.addWidget(QLabel("设备 1:", font=self.chinese_font), 1, 0); self.edit_dev1_combo = QComboBox(); self.edit_dev1_combo.setFont(self.chinese_font); add_manual_layout.addWidget(self.edit_dev1_combo, 1, 1); add_manual_layout.addWidget(QLabel("端口 1:", font=self.chinese_font), 1, 2); self.edit_port1_combo = QComboBox(); self.edit_port1_combo.setFont(self.chinese_font); add_manual_layout.addWidget(self.edit_port1_combo, 1, 3)
        add_manual_layout.addWidget(QLabel("设备 2:", font=self.chinese_font), 2, 0); self.edit_dev2_combo = QComboBox(); self.edit_dev2_combo.setFont(self.chinese_font); add_manual_layout.addWidget(self.edit_dev2_combo, 2, 1); add_manual_layout.addWidget(QLabel("端口 2:", font=self.chinese_font), 2, 2); self.edit_port2_combo = QComboBox(); self.edit_port2_combo.setFont(self.chinese_font); add_manual_layout.addWidget(self.edit_port2_combo, 2, 3)
        self.add_manual_button = QPushButton("添加连接"); self.add_manual_button.setFont(self.chinese_font); self.add_manual_button.clicked.connect(self.add_manual_connection); add_manual_layout.addWidget(self.add_manual_button, 3, 0, 1, 4, alignment=Qt.AlignmentFlag.AlignCenter)
        self.edit_dev1_combo.currentIndexChanged.connect(lambda: self._populate_edit_port_combos(self.edit_dev1_combo, self.edit_port1_combo)); self.edit_dev2_combo.currentIndexChanged.connect(lambda: self._populate_edit_port_combos(self.edit_dev2_combo, self.edit_port2_combo)); edit_main_layout.addWidget(add_manual_group)
        remove_manual_group = QFrame(); remove_manual_group.setObjectName("removeManualGroup"); remove_manual_group.setFrameShape(QFrame.Shape.StyledPanel); remove_manual_layout = QVBoxLayout(remove_manual_group); remove_manual_layout.setContentsMargins(10,10,10,10); remove_title = QLabel("<b>移除现有连接</b> (选中下方列表中的连接进行移除)"); remove_title.setFont(self.chinese_font); remove_manual_layout.addWidget(remove_title)
        filter_conn_layout = QHBoxLayout()
        filter_conn_layout.addWidget(QLabel("类型过滤:", font=self.chinese_font))
        self.conn_filter_type_combo = QComboBox()
        self.conn_filter_type_combo.setFont(self.chinese_font)
        self.conn_filter_type_combo.addItems(["所有类型", "LC-LC (100G)", "MPO-MPO (25G)", "SFP-SFP (10G)", "MPO-SFP (10G)"])
        self.conn_filter_type_combo.currentIndexChanged.connect(self.filter_connection_list)
        filter_conn_layout.addWidget(self.conn_filter_type_combo)
        filter_conn_layout.addSpacing(15)
        filter_conn_layout.addWidget(QLabel("设备过滤:", font=self.chinese_font))
        self.conn_filter_device_entry = QLineEdit()
        self.conn_filter_device_entry.setFont(self.chinese_font)
        self.conn_filter_device_entry.setPlaceholderText("按设备名称过滤...")
        self.conn_filter_device_entry.textChanged.connect(self.filter_connection_list)
        filter_conn_layout.addWidget(self.conn_filter_device_entry)
        remove_manual_layout.insertLayout(1, filter_conn_layout)
        self.manual_connection_list = QListWidget(); self.manual_connection_list.setFont(self.chinese_font); self.manual_connection_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection); remove_manual_layout.addWidget(self.manual_connection_list); self.remove_manual_button = QPushButton("移除选中连接"); self.remove_manual_button.setFont(self.chinese_font); self.remove_manual_button.clicked.connect(self.remove_manual_connection); self.remove_manual_button.setEnabled(False); remove_manual_layout.addWidget(self.remove_manual_button, alignment=Qt.AlignmentFlag.AlignCenter); edit_main_layout.addWidget(remove_manual_group); self.tab_widget.addTab(self.edit_tab, "手动编辑")
        main_splitter.setSizes([400, 700])
        main_splitter.setStretchFactor(1, 1)
        self.update_port_entries()

    # --- 方法 (与 V36 相同，除了事件处理) ---
    @Slot()
    def update_port_entries(self):
        selected_type = self.device_type_combo.currentText()
        is_micron = selected_type == 'MicroN'; is_uhd_horizon = selected_type in ['MicroN UHD', 'HorizoN']
        self.mpo_label.setVisible(is_uhd_horizon); self.mpo_entry.setVisible(is_uhd_horizon)
        self.lc_label.setVisible(is_uhd_horizon); self.lc_entry.setVisible(is_uhd_horizon)
        self.sfp_label.setVisible(is_micron); self.sfp_entry.setVisible(is_micron)

    def _add_device_to_table(self, device):
        row_position = self.device_tablewidget.rowCount(); self.device_tablewidget.insertRow(row_position)
        name_item = QTableWidgetItem(device.name); name_item.setData(Qt.ItemDataRole.UserRole, device.id)
        type_item = QTableWidgetItem(device.type)
        mpo_item = NumericTableWidgetItem(str(device.mpo_total)); mpo_item.setData(Qt.ItemDataRole.UserRole + 1, device.mpo_total); mpo_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        lc_item = NumericTableWidgetItem(str(device.lc_total)); lc_item.setData(Qt.ItemDataRole.UserRole + 1, device.lc_total); lc_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        sfp_item = NumericTableWidgetItem(str(device.sfp_total)); sfp_item.setData(Qt.ItemDataRole.UserRole + 1, device.sfp_total); sfp_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        conn_val = float(f"{device.connections:.2f}")
        conn_item = NumericTableWidgetItem(f"{conn_val:.2f}"); conn_item.setData(Qt.ItemDataRole.UserRole + 1, conn_val); conn_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.device_tablewidget.setItem(row_position, 0, name_item); self.device_tablewidget.setItem(row_position, 1, type_item)
        self.device_tablewidget.setItem(row_position, 2, mpo_item); self.device_tablewidget.setItem(row_position, 3, lc_item)
        self.device_tablewidget.setItem(row_position, 4, sfp_item); self.device_tablewidget.setItem(row_position, 5, conn_item)

    def _update_device_combos(self):
        self.edit_dev1_combo.blockSignals(True); self.edit_dev2_combo.blockSignals(True)
        current_dev1_id = self.edit_dev1_combo.currentData(); current_dev2_id = self.edit_dev2_combo.currentData()
        self.edit_dev1_combo.clear(); self.edit_dev2_combo.clear()
        self.edit_dev1_combo.addItem("选择设备 1...", userData=None); self.edit_dev2_combo.addItem("选择设备 2...", userData=None)
        sorted_devices = sorted(self.devices, key=lambda dev: dev.name)
        idx1_to_select, idx2_to_select = 0, 0
        for i, dev in enumerate(sorted_devices):
            item_text = f"{dev.name} ({dev.type})"; self.edit_dev1_combo.addItem(item_text, userData=dev.id); self.edit_dev2_combo.addItem(item_text, userData=dev.id)
            if dev.id == current_dev1_id: idx1_to_select = i + 1
            if dev.id == current_dev2_id: idx2_to_select = i + 1
        self.edit_dev1_combo.setCurrentIndex(idx1_to_select); self.edit_dev2_combo.setCurrentIndex(idx2_to_select)
        self.edit_dev1_combo.blockSignals(False); self.edit_dev2_combo.blockSignals(False)
        self._populate_edit_port_combos(self.edit_dev1_combo, self.edit_port1_combo)
        self._populate_edit_port_combos(self.edit_dev2_combo, self.edit_port2_combo)

    def _populate_edit_port_combos(self, device_combo, port_combo):
        port_combo.blockSignals(True)
        port_combo.clear(); port_combo.addItem("选择端口...")
        dev_id = device_combo.currentData(); port_combo.setEnabled(False)
        if dev_id is not None:
            device = next((d for d in self.devices if d.id == dev_id), None)
            if device:
                available_ports = device.get_all_available_ports()
                if available_ports: port_combo.addItems(available_ports); port_combo.setEnabled(True)
                else: port_combo.addItem("无可用端口")
        port_combo.blockSignals(False)

    @Slot()
    def add_device(self):
        dtype = self.device_type_combo.currentText(); name = self.device_name_entry.text().strip()
        if not dtype: QMessageBox.critical(self, "错误", "请选择设备类型。"); return
        if not name: QMessageBox.critical(self, "错误", "请输入设备名称。"); return
        if any(dev.name == name for dev in self.devices): QMessageBox.critical(self, "错误", f"设备名称 '{name}' 已存在。"); return
        mpo_ports_str, lc_ports_str, sfp_ports_str = self.mpo_entry.text() or "0", self.lc_entry.text() or "0", self.sfp_entry.text() or "0"
        try:
            if dtype in ['MicroN UHD', 'HorizoN']: mpo_ports, lc_ports, sfp_ports = int(mpo_ports_str), int(lc_ports_str), 0; assert mpo_ports >= 0 and lc_ports >= 0
            elif dtype == 'MicroN': sfp_ports, mpo_ports, lc_ports = int(sfp_ports_str), 0, 0; assert sfp_ports >= 0
            else: raise ValueError("无效类型")
        except (ValueError, AssertionError): QMessageBox.critical(self, "输入错误", "端口数量必须是非负整数。"); return
        self.device_id_counter += 1; new_device = Device(self.device_id_counter, name, dtype, mpo_ports, lc_ports, sfp_ports)
        self.devices.append(new_device); self._add_device_to_table(new_device); self.device_name_entry.clear();
        self._update_device_combos(); self.clear_results()
        self.node_positions = None; self.selected_node_id = None

    @Slot()
    def remove_device(self):
        selected_rows = sorted(list(set(index.row() for index in self.device_tablewidget.selectedIndexes())), reverse=True)
        if not selected_rows: QMessageBox.warning(self, "提示", "请先在表格中选择要移除的设备行。"); return
        ids_to_remove = {self.device_tablewidget.item(row_index, 0).data(Qt.ItemDataRole.UserRole) for row_index in selected_rows if self.device_tablewidget.item(row_index, 0)}
        self.devices = [dev for dev in self.devices if dev.id not in ids_to_remove]
        for row_index in selected_rows: self.device_tablewidget.removeRow(row_index)
        self._update_device_combos(); self.clear_results()
        self.node_positions = None; self.selected_node_id = None

    @Slot()
    def clear_all_devices(self):
        if QMessageBox.question(self, "确认", "确定要清空所有设备吗？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            self.devices = []; self.device_tablewidget.setRowCount(0); self.device_id_counter = 0;
            self._update_device_combos(); self.clear_results()
            self.node_positions = None; self.selected_node_id = None

    @Slot()
    def clear_results(self):
        self.connections_result = []; self.fig = None; self.node_positions = None; self.selected_node_id = None
        self.connections_textedit.clear(); self.manual_connection_list.clear()
        self.mpl_canvas.axes.cla(); self.mpl_canvas.axes.text(0.5, 0.5, '点击“计算”生成图形', ha='center', va='center'); self.mpl_canvas.draw()
        self.export_list_button.setEnabled(False); self.export_topo_button.setEnabled(False); self.remove_manual_button.setEnabled(False)
        self.fill_mesh_button.setEnabled(False); self.fill_ring_button.setEnabled(False)
        for dev in self.devices: dev.reset_ports()
        self._update_device_table_connections()

    @Slot(QTableWidgetItem)
    def show_device_details_from_table(self, item):
         if not item: return
         row = item.row(); name_item = self.device_tablewidget.item(row, 0)
         if not name_item: return
         dev_id = name_item.data(Qt.ItemDataRole.UserRole)
         dev = next((d for d in self.devices if d.id == dev_id), None)
         if not dev: return QMessageBox.critical(self, "错误", "无法找到所选设备的详细信息。")
         self._display_device_details_popup(dev)

    def _display_device_details_popup(self, dev):
        details = f"ID: {dev.id}\n名称: {dev.name}\n类型: {dev.type}\n"
        avail_ports = dev.get_all_available_ports()
        avail_lc_count = sum(1 for p in avail_ports if p.startswith("LC"))
        avail_sfp_count = sum(1 for p in avail_ports if p.startswith("SFP"))
        avail_mpo_ch_count = sum(1 for p in avail_ports if p.startswith("MPO"))
        if dev.type in ['MicroN UHD', 'HorizoN']: details += f"MPO 端口总数: {dev.mpo_total}\nLC 端口总数: {dev.lc_total}\n可用 MPO 子通道: {avail_mpo_ch_count}\n可用 LC 端口: {avail_lc_count}\n"
        elif dev.type == 'MicroN': details += f"SFP+ 端口总数: {dev.sfp_total}\n可用 SFP+ 端口: {avail_sfp_count}\n"
        details += f"当前连接数 (估算): {dev.connections:.2f}\n"
        if dev.port_connections:
            details += "\n端口连接详情:\n"
            lc_conns = {p: t for p, t in dev.port_connections.items() if p.startswith("LC")}
            sfp_conns = {p: t for p, t in dev.port_connections.items() if p.startswith("SFP")}
            mpo_conns_grouped = defaultdict(dict)
            for p, t in dev.port_connections.items():
                if p.startswith("MPO"): mpo_conns_grouped[p.split('-')[0]][p] = t
            if lc_conns: details += "  LC 连接:\n"; [details := details + f"    {port} -> {lc_conns[port]}\n" for port in sorted(lc_conns.keys())]
            if sfp_conns: details += "  SFP+ 连接:\n"; [details := details + f"    {port} -> {sfp_conns[port]}\n" for port in sorted(sfp_conns.keys())]
            if mpo_conns_grouped:
                details += "  MPO 连接 (Breakout):\n"
                for base_port in sorted(mpo_conns_grouped.keys()):
                    details += f"    {base_port}:\n"; [details := details + f"      {port} -> {mpo_conns_grouped[base_port][port]}\n" for port in sorted(mpo_conns_grouped[base_port].keys(), key=lambda x: int(x.split('-Ch')[-1]))]
        QMessageBox.information(self, f"设备详情 - {dev.name}", details)

    @Slot()
    def calculate_and_display(self):
        """计算连接并显示结果"""
        if not self.devices: QMessageBox.information(self, "提示", "请先添加设备。"); return
        self.node_positions = None; self.selected_node_id = None
        mode = self.topology_mode_combo.currentText()
        calculated_connections, final_device_state_map = [], {}
        for dev in self.devices: dev.reset_ports()
        if mode == "Mesh": print("使用改进的 Mesh 算法进行计算..."); calculated_connections, final_device_state_map = calculate_mesh_connections(self.devices)
        elif mode == "环形": print("使用环形算法进行计算..."); calculated_connections, final_device_state_map = calculate_ring_connections(self.devices)
        else: QMessageBox.critical(self, "错误", f"未知的计算模式: {mode}"); return
        self.connections_result = calculated_connections
        for dev in self.devices:
            if dev.id in final_device_state_map:
                final_state_dev = final_device_state_map[dev.id]
                dev.connections = final_state_dev.connections
                dev.port_connections = final_state_dev.port_connections.copy()
            else: dev.reset_ports()
        self._update_connection_views(); self._update_device_table_connections()
        self._update_device_combos()
        enable_fill = bool(self.connections_result) or any(dev.get_all_available_ports() for dev in self.devices)
        self.fill_mesh_button.setEnabled(enable_fill); self.fill_ring_button.setEnabled(enable_fill)

    def _update_connection_views(self):
        """更新连接列表、手动编辑列表和拓扑图 (包含高亮)"""
        self.connections_textedit.clear()
        if self.connections_result:
            self.connections_textedit.append("<b>连接列表:</b><hr>")
            for i, conn in enumerate(self.connections_result): dev1, port1, dev2, port2, conn_type = conn; self.connections_textedit.append(f"{i+1}. {dev1.name} [{port1}] &lt;-&gt; {dev2.name} [{port2}] ({conn_type})")
            self.export_list_button.setEnabled(True)
        else: self.connections_textedit.append("无连接。"); self.export_list_button.setEnabled(False)
        self.manual_connection_list.clear()
        if self.connections_result:
            for i, conn in enumerate(self.connections_result):
                dev1, port1, dev2, port2, conn_type = conn; item_text = f"{i+1}. {dev1.name} [{port1}] <-> {dev2.name} [{port2}] ({conn_type})"
                item = QListWidgetItem(item_text); item.setData(Qt.ItemDataRole.UserRole, conn); self.manual_connection_list.addItem(item)
            self.remove_manual_button.setEnabled(True)
        else: self.remove_manual_button.setEnabled(False)
        self.filter_connection_list()
        selected_layout = self.layout_combo.currentText().lower()
        # 传递 node_positions 和 selected_node_id
        self.fig, calculated_pos = self.mpl_canvas.plot_topology(
            self.devices, self.connections_result,
            layout_algorithm=selected_layout,
            fixed_pos=self.node_positions,
            selected_node_id=self.selected_node_id
        )
        # 只有在布局是新计算出来的时候才更新存储的位置
        if calculated_pos is not None and self.node_positions is None:
            self.node_positions = calculated_pos
        elif calculated_pos is not None and self.node_positions is not None and self.dragged_node_id is None:
             # 如果不是拖动更新，并且布局被重新计算了（例如改变布局算法），也更新位置
             # 检查 fixed_pos 是否在 plot_topology 内部被重置为 None
             # 更好的方法是让 plot_topology 返回是否使用了 fixed_pos
             # 暂时简化：如果 calculated_pos 不是 None，就更新
             # 但这会导致拖动后切换布局再切回来时位置丢失，需要在 on_layout_change 中处理
             # 已经处理：on_layout_change 会将 self.node_positions 设为 None
             pass # 拖动时不更新整体布局

        self.export_topo_button.setEnabled(bool(self.fig))

    def _update_device_table_connections(self):
        """更新设备表格中的连接数估算列，并设置数字排序"""
        for row in range(self.device_tablewidget.rowCount()):
            name_item = self.device_tablewidget.item(row, 0)
            if name_item:
                dev_id = name_item.data(Qt.ItemDataRole.UserRole)
                device = next((d for d in self.devices if d.id == dev_id), None)
                if device:
                    conn_val = float(f"{device.connections:.2f}")
                    conn_item = NumericTableWidgetItem(f"{conn_val:.2f}")
                    conn_item.setData(Qt.ItemDataRole.UserRole + 1, conn_val)
                    conn_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.device_tablewidget.setItem(row, 5, conn_item)

    @Slot()
    def remove_manual_connection(self):
        selected_items = self.manual_connection_list.selectedItems()
        if not selected_items: QMessageBox.warning(self, "提示", "请在下方列表中选择要移除的连接。"); return
        items_to_remove_from_widget, connections_removed_count, connections_to_remove_data = [], 0, []
        for item in selected_items: conn_data = item.data(Qt.ItemDataRole.UserRole); connections_to_remove_data.append(conn_data); items_to_remove_from_widget.append(item)
        for conn_data in connections_to_remove_data:
            dev1_orig, port1, dev2_orig, port2, conn_type = conn_data
            try:
                found_index = -1
                for i, existing_conn in enumerate(self.connections_result):
                    if (existing_conn[0].id == dev1_orig.id and existing_conn[1] == port1 and existing_conn[2].id == dev2_orig.id and existing_conn[3] == port2) or \
                       (existing_conn[0].id == dev2_orig.id and existing_conn[1] == port2 and existing_conn[2].id == dev1_orig.id and existing_conn[3] == port1):
                        found_index = i; break
                if found_index != -1:
                    self.connections_result.pop(found_index)
                    dev1_orig.return_port(port1); dev2_orig.return_port(port2)
                    connections_removed_count += 1
                else: print(f"警告: 移除时未找到匹配项: {conn_data}")
            except Exception as e: print(f"警告: 移除连接时出错: {e} - {conn_data}")
        if connections_removed_count > 0:
             self.node_positions = None; self.selected_node_id = None
             for item in items_to_remove_from_widget:
                 row = self.manual_connection_list.row(item)
                 if row != -1: self.manual_connection_list.takeItem(row)
             self._update_connection_views(); self._update_device_table_connections()
             self._update_device_combos()
             print(f"成功移除了 {connections_removed_count} 条连接。")
             enable_fill = bool(self.connections_result) or any(dev.get_all_available_ports() for dev in self.devices)
             self.fill_mesh_button.setEnabled(enable_fill); self.fill_ring_button.setEnabled(enable_fill)
        else: print("没有连接被移除。")

    @Slot()
    def add_manual_connection(self):
        dev1_id = self.edit_dev1_combo.currentData(); dev2_id = self.edit_dev2_combo.currentData()
        port1_text = self.edit_port1_combo.currentText(); port2_text = self.edit_port2_combo.currentText()
        if dev1_id is None or dev2_id is None or port1_text == "选择端口..." or port2_text == "选择端口...": QMessageBox.warning(self, "选择不完整", "请选择两个设备和它们各自的端口。"); return
        if dev1_id == dev2_id: QMessageBox.warning(self, "选择错误", "不能将设备连接到自身。"); return
        dev1 = next((d for d in self.devices if d.id == dev1_id), None); dev2 = next((d for d in self.devices if d.id == dev2_id), None)
        if not dev1 or not dev2: QMessageBox.critical(self, "内部错误", "找不到选定的设备对象。"); return
        port1_type = "LC" if port1_text.startswith("LC") else ("SFP" if port1_text.startswith("SFP") else ("MPO" if port1_text.startswith("MPO") else "Unknown"))
        port2_type = "LC" if port2_text.startswith("LC") else ("SFP" if port2_text.startswith("SFP") else ("MPO" if port2_text.startswith("MPO") else "Unknown"))
        conn_type_str = ""; valid_connection = False
        dev1_ref, dev2_ref = dev1, dev2; port1_ref, port2_ref = port1_text, port2_text
        if port1_type == "LC" and port2_type == "LC" and dev1_ref.type in ['MicroN UHD', 'HorizoN'] and dev2_ref.type in ['MicroN UHD', 'HorizoN']: valid_connection = True; conn_type_str = "LC-LC (100G)"
        elif port1_type == "MPO" and port2_type == "MPO" and dev1_ref.type in ['MicroN UHD', 'HorizoN'] and dev2_ref.type in ['MicroN UHD', 'HorizoN']: valid_connection = True; conn_type_str = "MPO-MPO (25G)"
        elif port1_type == "SFP" and port2_type == "SFP" and dev1_ref.type == 'MicroN' and dev2_ref.type == 'MicroN': valid_connection = True; conn_type_str = "SFP-SFP (10G)"
        elif port1_type == "MPO" and port2_type == "SFP" and dev1_ref.type in ['MicroN UHD', 'HorizoN'] and dev2_ref.type == 'MicroN': valid_connection = True; conn_type_str = "MPO-SFP (10G)"
        elif port1_type == "SFP" and port2_type == "MPO" and dev1_ref.type == 'MicroN' and dev2_ref.type in ['MicroN UHD', 'HorizoN']:
             dev1, dev2 = dev2_ref, dev1_ref; port1_text, port2_text = port2_ref, port1_ref
             valid_connection = True; conn_type_str = "MPO-SFP (10G)"
        if not valid_connection: QMessageBox.warning(self, "连接无效", f"设备类型 '{dev1_ref.type}' 的端口 '{port1_type}' 与 设备类型 '{dev2_ref.type}' 的端口 '{port2_type}' 之间不允许连接。"); return
        if dev1.use_specific_port(port1_text, dev2.name):
            if dev2.use_specific_port(port2_text, dev1.name):
                self.node_positions = None; self.selected_node_id = None
                self.connections_result.append((dev1_ref, port1_ref, dev2_ref, port2_ref, conn_type_str))
                self._update_connection_views(); self._update_device_table_connections()
                self._populate_edit_port_combos(self.edit_dev1_combo, self.edit_port1_combo)
                self._populate_edit_port_combos(self.edit_dev2_combo, self.edit_port2_combo)
                print(f"成功添加手动连接: {dev1_ref.name}[{port1_ref}] <-> {dev2_ref.name}[{port2_ref}]")
                enable_fill = bool(self.connections_result) or any(dev.get_all_available_ports() for dev in self.devices)
                self.fill_mesh_button.setEnabled(enable_fill); self.fill_ring_button.setEnabled(enable_fill)
            else:
                dev1.return_port(port1_text)
                QMessageBox.critical(self, "错误", f"端口 '{port2_ref}' 在设备 '{dev2_ref.name}' 上不可用或已被占用。")
                self._populate_edit_port_combos(self.edit_dev1_combo if dev1.id == dev1_id else self.edit_dev2_combo, self.edit_port1_combo if dev1.id == dev1_id else self.edit_port2_combo)
        else: QMessageBox.critical(self, "错误", f"端口 '{port1_ref}' 在设备 '{dev1_ref.name}' 上不可用或已被占用。")

    @Slot()
    def fill_remaining_mesh(self):
        """使用 Mesh 算法填充剩余的可用端口 (V35 修正)"""
        if not self.devices: QMessageBox.information(self, "提示", "请先添加设备。"); return
        print("开始填充剩余连接 (Mesh)...")
        new_connections = _fill_connections_mesh_style(self.devices)
        if new_connections:
            print(f"找到 {len(new_connections)} 条新 Mesh 连接可以添加。")
            self.connections_result.extend(new_connections)
            print("更新设备状态...")
            for conn in new_connections:
                dev1, port1, dev2, port2, _ = conn
                if not dev1.use_specific_port(port1, dev2.name): print(f"严重警告: Mesh 填充同步状态时未能使用端口 {dev1.name}[{port1}]")
                if not dev2.use_specific_port(port2, dev1.name): print(f"严重警告: Mesh 填充同步状态时未能使用端口 {dev2.name}[{port2}]")
            print("更新视图...")
            self.selected_node_id = None
            self._update_connection_views(); self._update_device_table_connections(); self._update_device_combos()
            QMessageBox.information(self, "填充完成", f"成功添加了 {len(new_connections)} 条新 Mesh 连接。")
        else: QMessageBox.information(self, "填充完成", "没有找到更多可以建立的 Mesh 连接。"); print("未找到可填充的新 Mesh 连接。")
        self.fill_mesh_button.setEnabled(False); self.fill_ring_button.setEnabled(False)

    @Slot()
    def fill_remaining_ring(self):
        """使用 Ring 算法填充剩余的可用端口 (V35 修正)"""
        if not self.devices: QMessageBox.information(self, "提示", "请先添加设备。"); return
        print("开始填充剩余连接 (环形)...")
        new_connections = _fill_connections_ring_style(self.devices)
        if new_connections:
            print(f"找到 {len(new_connections)} 条新环形连接段可以添加。")
            self.connections_result.extend(new_connections)
            print("更新设备状态...")
            for conn in new_connections:
                dev1, port1, dev2, port2, _ = conn
                if not dev1.use_specific_port(port1, dev2.name): print(f"严重警告: 环形填充同步状态时未能使用端口 {dev1.name}[{port1}]")
                if not dev2.use_specific_port(port2, dev1.name): print(f"严重警告: 环形填充同步状态时未能使用端口 {dev2.name}[{port2}]")
            print("更新视图...")
            self.selected_node_id = None
            self._update_connection_views(); self._update_device_table_connections(); self._update_device_combos()
            QMessageBox.information(self, "填充完成", f"成功添加了 {len(new_connections)} 条新环形连接段。")
        else: QMessageBox.information(self, "填充完成", "没有找到更多可以建立的环形连接段。"); print("未找到可填充的新环形连接段。")
        self.fill_mesh_button.setEnabled(False); self.fill_ring_button.setEnabled(False)

    @Slot()
    def on_layout_change(self):
        if self.connections_result:
            print("DIAG: Layout selection changed, resetting stored positions.")
            self.node_positions = None; self.selected_node_id = None
            self._update_connection_views()

    @Slot()
    def save_config(self):
        if not self.devices: QMessageBox.warning(self, "提示", "设备列表为空。"); return
        filepath, _ = QFileDialog.getSaveFileName(self, "保存设备配置", "", "JSON 文件 (*.json);;所有文件 (*)")
        if not filepath: return
        config_data = [dev.to_dict() for dev in self.devices]
        try:
            with open(filepath, 'w', encoding='utf-8') as f: json.dump(config_data, f, indent=4, ensure_ascii=False)
            QMessageBox.information(self, "成功", f"配置已保存到:\n{filepath}")
        except Exception as e: QMessageBox.critical(self, "保存失败", f"无法保存配置文件:\n{e}")

    @Slot()
    def load_config(self):
        if self.devices and QMessageBox.question(self, "确认", "加载配置将覆盖当前设备列表，确定吗？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No) == QMessageBox.StandardButton.No: return
        filepath, _ = QFileDialog.getOpenFileName(self, "加载设备配置", "", "JSON 文件 (*.json);;所有文件 (*)")
        if not filepath: return
        try:
            with open(filepath, 'r', encoding='utf-8') as f: config_data = json.load(f)
            self.devices = []; self.device_tablewidget.setRowCount(0); self.device_id_counter = 0; self.clear_results(); max_id = 0
            for data in config_data:
                new_device = Device.from_dict(data); self.devices.append(new_device); self._add_device_to_table(new_device)
                if new_device.id > max_id: max_id = new_device.id
            self.device_id_counter = max_id
            self._update_device_combos()
            QMessageBox.information(self, "成功", f"配置已从以下文件加载:\n{filepath}")
            self.node_positions = None; self.selected_node_id = None
        except FileNotFoundError: QMessageBox.critical(self, "加载失败", "找不到文件。")
        except json.JSONDecodeError: QMessageBox.critical(self, "加载失败", "文件格式错误。")
        except Exception as e: QMessageBox.critical(self, "加载失败", f"发生错误:\n{e}")

    @Slot()
    def export_connections(self):
        if not self.connections_result: QMessageBox.warning(self, "提示", "没有连接结果可导出。"); return
        filepath, selected_filter = QFileDialog.getSaveFileName(self, "导出连接列表", "", "文本文件 (*.txt);;CSV 文件 (*.csv);;所有文件 (*)")
        if not filepath: return
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                is_csv = "csv" in selected_filter.lower()
                if is_csv: f.write("序号,设备1,端口1,设备2,端口2,连接类型\n")
                for i, conn in enumerate(self.connections_result):
                    dev1, port1, dev2, port2, conn_type = conn
                    if is_csv: f.write(f"{i+1},{dev1.name},{port1},{dev2.name},{port2},{conn_type}\n")
                    else: f.write(f"{i+1}. {dev1.name} [{port1}] <-> {dev2.name} [{port2}] ({conn_type})\n")
            QMessageBox.information(self, "成功", f"连接列表已导出到:\n{filepath}")
        except Exception as e: QMessageBox.critical(self, "导出失败", f"无法导出连接列表:\n{e}")

    @Slot()
    def export_topology(self):
        if not self.fig: QMessageBox.warning(self, "提示", "没有拓扑图可导出。"); return
        filepath, selected_filter = QFileDialog.getSaveFileName(self, "导出拓扑图", "", "PNG 图像 (*.png);;PDF 文件 (*.pdf);;SVG 文件 (*.svg);;所有文件 (*)")
        if not filepath: return
        try:
            self.fig.savefig(filepath, dpi=300, bbox_inches='tight')
            QMessageBox.information(self, "成功", f"拓扑图已导出到:\n{filepath}")
        except Exception as e: QMessageBox.critical(self, "导出失败", f"无法导出拓扑图:\n{e}")

    @Slot(str)
    def filter_device_table(self, text):
        filter_text = text.lower()
        for row in range(self.device_tablewidget.rowCount()):
            match = False
            name_item = self.device_tablewidget.item(row, 0)
            type_item = self.device_tablewidget.item(row, 1)
            if name_item and filter_text in name_item.text().lower(): match = True
            if not match and type_item and filter_text in type_item.text().lower(): match = True
            self.device_tablewidget.setRowHidden(row, not match)

    @Slot()
    def filter_connection_list(self):
        selected_type = self.conn_filter_type_combo.currentText()
        filter_device_text = self.conn_filter_device_entry.text().strip().lower()
        type_filter_active = selected_type != "所有类型"
        for i in range(self.manual_connection_list.count()):
            item = self.manual_connection_list.item(i)
            conn_data = item.data(Qt.ItemDataRole.UserRole)
            if not conn_data or not isinstance(conn_data, tuple) or len(conn_data) != 5:
                item.setHidden(False); continue
            dev1, _, dev2, _, conn_type = conn_data
            item_conn_type = conn_type
            dev1_name_lower = dev1.name.lower(); dev2_name_lower = dev2.name.lower()
            type_match = True
            if type_filter_active: type_match = (item_conn_type == selected_type)
            device_match = True
            if filter_device_text: device_match = (filter_device_text in dev1_name_lower or filter_device_text in dev2_name_lower)
            item.setHidden(not (type_match and device_match))

    # --- 新增：画布事件处理 (包含拖动逻辑) ---
    @Slot(object)
    def on_canvas_press(self, event):
        """处理画布上的鼠标按下事件 (单击, 双击, 开始拖动)"""
        if event.inaxes != self.mpl_canvas.axes or not self.node_positions:
            needs_redraw = self.selected_node_id is not None or self.dragged_node_id is not None
            self.selected_node_id = None; self.dragged_node_id = None
            if needs_redraw: self._update_connection_views()
            return

        x, y = event.xdata, event.ydata
        if x is None or y is None:
            needs_redraw = self.selected_node_id is not None or self.dragged_node_id is not None
            self.selected_node_id = None; self.dragged_node_id = None
            if needs_redraw: self._update_connection_views()
            return

        clicked_node_id = None; min_dist_sq = float('inf')
        xlim = self.mpl_canvas.axes.get_xlim(); ylim = self.mpl_canvas.axes.get_ylim()
        diagonal_len_sq = (xlim[1]-xlim[0])**2 + (ylim[1]-ylim[0])**2
        threshold_dist_sq = diagonal_len_sq * (0.03**2) # 点击阈值平方
        for node_id, (nx, ny) in self.node_positions.items():
            dist_sq = (x - nx)**2 + (y - ny)**2
            if dist_sq < min_dist_sq and dist_sq < threshold_dist_sq:
                 min_dist_sq = dist_sq; clicked_node_id = node_id

        if event.dblclick:
            self.dragged_node_id = None # 取消拖动状态
            if clicked_node_id is not None:
                print(f"双击节点: {clicked_node_id}")
                device = next((d for d in self.devices if d.id == clicked_node_id), None)
                if device: self._display_device_details_popup(device)
        elif event.button == 1: # 左键按下
            if clicked_node_id is not None:
                self.dragged_node_id = clicked_node_id # 准备开始拖动
                nx, ny = self.node_positions[clicked_node_id]
                self.drag_offset = (x - nx, y - ny) # 计算偏移
                if self.selected_node_id != clicked_node_id:
                    self.selected_node_id = clicked_node_id
                    print(f"选中节点 (准备拖动): {self.selected_node_id}")
                    self._update_connection_views() # 重绘以显示选中高亮
                else:
                     print(f"开始拖动节点: {self.selected_node_id}")
            else:
                self.dragged_node_id = None
                if self.selected_node_id is not None:
                    self.selected_node_id = None
                    print("清除选中 (点击背景)")
                    self._update_connection_views()

    @Slot(object)
    def on_canvas_motion(self, event):
        """处理画布上的鼠标移动事件 (用于拖动节点)"""
        if self.dragged_node_id is None or event.button != 1 or event.inaxes != self.mpl_canvas.axes:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None: return

        new_x = x - self.drag_offset[0]
        new_y = y - self.drag_offset[1]
        self.node_positions[self.dragged_node_id] = (new_x, new_y)
        # 请求重绘 (使用 draw_idle 避免过于频繁的绘制)
        # 注意：这里只更新了 self.node_positions，需要调用 _update_connection_views 来触发使用新位置的重绘
        self._update_connection_views() # 触发重绘

    @Slot(object)
    def on_canvas_release(self, event):
        """处理画布上的鼠标释放事件 (结束拖动)"""
        if event.button == 1 and self.dragged_node_id is not None:
            print(f"结束拖动节点: {self.dragged_node_id}")
            # 最终位置已在 motion 事件中更新并触发了重绘
            self.dragged_node_id = None # 清除拖动状态
            # 可选：触发一次最终的 draw_idle 确保状态
            self.mpl_canvas.draw_idle()
    # --- 结束事件处理 ---

# --- 程序入口 (与 V36 相同) ---
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(APP_STYLE)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QTextEdit,
    QTabWidget, QFrame, QFileDialog, QMessageBox, QSpacerItem, QSizePolicy,
    QGridLayout, QListWidgetItem, QAbstractItemView, QTableWidget, QTableWidgetItem,
    QHeaderView, QListWidget, QSplitter
)
from PySide6.QtCore import Slot, Qt
from PySide6.QtGui import QFont

# --- QSS 样式定义 (与 V35 相同) ---
APP_STYLE = """
QMainWindow, QDialog, QMessageBox { /* 应用于主窗口、对话框和消息框 */
    background-color: #f0f0f0; /* 浅灰色背景 */
}
QFrame {
    background-color: transparent; /* 让 Frame 透明，除非特别指定 */
}
/* 为特定的分组 Frame 添加边框和圆角 */
QFrame#addDeviceGroup, QFrame#listGroup, QFrame#fileGroup,
QFrame#calculateControlFrame, QFrame#addManualGroup, QFrame#removeManualGroup {
    border: 1px solid #c8c8c8;
    border-radius: 5px;
    background-color: #f8f8f8; /* 轻微区分背景 */
    margin-bottom: 5px; /* 组之间增加一点间距 */
}
QPushButton {
    background-color: #e1e1e1;
    border: 1px solid #adadad;
    padding: 5px 10px;
    border-radius: 4px;
    min-height: 20px;
    min-width: 75px; /* 按钮最小宽度 */
}
QPushButton:hover {
    background-color: #cacaca;
    border: 1px solid #999999;
}
QPushButton:pressed {
    background-color: #b0b0b0;
    border: 1px solid #777777;
}
QPushButton:disabled {
    background-color: #d3d3d3;
    color: #a0a0a0;
    border: 1px solid #c0c0c0;
}
QLineEdit, QComboBox, QTextEdit, QListWidget, QTableWidget {
    background-color: white;
    border: 1px solid #c0c0c0;
    border-radius: 3px;
    padding: 3px;
    selection-background-color: #a8cce4; /* 选中项背景色 */
    selection-color: black; /* 选中项文字颜色 */
}
QComboBox::drop-down { /* 下拉箭头样式 */
    border: none;
    background: transparent;
    width: 15px; /* 增加箭头区域宽度 */
    padding-right: 5px;
}
QComboBox::down-arrow { /* 箭头图标 */
     width: 12px;
     height: 12px;
}
QTableWidget {
    alternate-background-color: #f8f8f8; /* 表格斑马纹 */
    gridline-color: #d0d0d0; /* 网格线颜色 */
}
QHeaderView::section { /* 表头样式 */
    background-color: #e8e8e8;
    padding: 4px;
    border: 1px solid #d0d0d0;
    border-left: none; /* 移除左边框避免双线 */
    font-weight: bold;
}
QHeaderView::section:first { /* 第一个表头单元格 */
     border-left: 1px solid #d0d0d0; /* 补回第一个左边框 */
}
QTabWidget::pane { /* 标签页主体框架 */
    border: 1px solid #c0c0c0;
    border-top: none; /* 顶部边框由标签栏提供 */
    background-color: white; /* 标签页内容区域背景 */
}
QTabBar::tab { /* 标签样式 */
    background: #e1e1e1;
    border: 1px solid #adadad;
    border-bottom: none;
    padding: 6px 12px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    margin-right: 2px; /* 标签间距 */
}
QTabBar::tab:selected {
    background: white; /* 选中标签页背景与内容区域一致 */
    margin-bottom: -1px; /* 轻微上移，与边框融合 */
    border-bottom: 1px solid white; /* 覆盖 pane 的顶部边框 */
}
QTabBar::tab:!selected:hover {
    background: #cacaca;
}
QSplitter::handle { /* 分隔条样式 */
    background-color: #d0d0d0;
    border: none;
}
QSplitter::handle:horizontal {
    width: 3px; /* 水平分隔条宽度 */
}
QSplitter::handle:vertical {
    height: 3px; /* 垂直分隔条高度 */
}
QSplitter::handle:hover {
    background-color: #b0b0b0;
}
"""

# --- 用于数字排序的 QTableWidgetItem 子类 (与 V35 相同) ---
class NumericTableWidgetItem(QTableWidgetItem):
    """自定义 QTableWidgetItem 以支持数字排序。"""
    def __lt__(self, other):
        data_self = self.data(Qt.ItemDataRole.UserRole + 1)
        data_other = other.data(Qt.ItemDataRole.UserRole + 1)
        try:
            num_self = float(data_self)
            num_other = float(data_other)
            return num_self < num_other
        except (TypeError, ValueError):
            return super().__lt__(other)

# --- 数据结构 (与 V35 相同) ---
class Device:
    """代表一个 MediorNet 设备"""
    def __init__(self, id, name, type, mpo_ports=0, lc_ports=0, sfp_ports=0):
        self.id = id
        self.name = name
        self.type = type
        self.mpo_total = mpo_ports
        self.lc_total = lc_ports
        self.sfp_total = sfp_ports
        self.reset_ports()

    def reset_ports(self):
        self.connections = 0.0
        self.port_connections = {}

    def get_all_possible_ports(self):
        ports = []
        ports.extend([f"LC{i+1}" for i in range(self.lc_total)])
        ports.extend([f"SFP{i+1}" for i in range(self.sfp_total)])
        for i in range(self.mpo_total):
            base = f"MPO{i+1}"
            ports.extend([f"{base}-Ch{j+1}" for j in range(4)])
        return ports

    def get_all_available_ports(self):
        all_ports = self.get_all_possible_ports()
        used_ports = set(self.port_connections.keys())
        available = [p for p in all_ports if p not in used_ports]
        return available

    def use_specific_port(self, port_name, target_device_name):
        if port_name in self.get_all_possible_ports() and port_name not in self.port_connections:
            self.port_connections[port_name] = target_device_name
            if port_name.startswith("MPO"): self.connections += 0.25
            else: self.connections += 1
            return True
        return False

    def return_port(self, port_name):
        port_in_use_record = port_name in self.port_connections
        port_already_available = False
        port_type_valid = True
        if port_name.startswith("LC") or port_name.startswith("SFP") or (port_name.startswith("MPO") and "-Ch" in port_name):
             port_already_available = port_name in self.get_all_available_ports()
        else:
            print(f"警告: 尝试归还未知类型的端口 {port_name}")
            port_type_valid = False
        if not port_type_valid: return
        if port_in_use_record:
            target = self.port_connections.pop(port_name)
            if not port_already_available:
                if port_name.startswith("MPO"): self.connections -= 0.25
                elif port_name.startswith("LC") or port_name.startswith("SFP"): self.connections -= 1
                self.connections = max(0.0, self.connections)
        else:
             pass

    def get_available_port(self, port_type, target_device_name):
        possible_ports = []
        if port_type == 'LC': possible_ports = [f"LC{i+1}" for i in range(self.lc_total)]
        elif port_type == 'MPO':
            for i in range(self.mpo_total): base = f"MPO{i+1}"; possible_ports.extend([f"{base}-Ch{j+1}" for j in range(4)])
            random.shuffle(possible_ports)
        elif port_type == 'SFP': possible_ports = [f"SFP{i+1}" for i in range(self.sfp_total)]
        used_ports = set(self.port_connections.keys())
        for port in possible_ports:
            if port not in used_ports:
                self.port_connections[port] = target_device_name
                if port.startswith("MPO"): self.connections += 0.25
                else: self.connections += 1
                return port
        return None

    def get_specific_available_port(self, port_type_prefix):
        possible_ports = []
        if port_type_prefix == 'LC': possible_ports = sorted([f"LC{i+1}" for i in range(self.lc_total)], key=lambda x: int(x[2:]))
        elif port_type_prefix == 'MPO':
            for i in range(self.mpo_total): base = f"MPO{i+1}"; possible_ports.extend(sorted([f"{base}-Ch{j+1}" for j in range(4)], key=lambda x: int(x.split('-Ch')[-1])))
        elif port_type_prefix == 'SFP': possible_ports = sorted([f"SFP{i+1}" for i in range(self.sfp_total)], key=lambda x: int(x[3:]))
        used_ports = set(self.port_connections.keys())
        for port in possible_ports:
            if port not in used_ports: return port
        return None

    def to_dict(self):
        return {'id': self.id, 'name': self.name, 'type': self.type,
                'mpo_ports': self.mpo_total, 'lc_ports': self.lc_total, 'sfp_ports': self.sfp_total}

    @classmethod
    def from_dict(cls, data):
        data.setdefault('id', random.randint(10000, 99999))
        return cls(id=data['id'], name=data['name'], type=data['type'],
                   mpo_ports=data.get('mpo_ports', 0), lc_ports=data.get('lc_ports', 0), sfp_ports=data.get('sfp_ports', 0))

    def __repr__(self):
        return f"{self.name} ({self.type})"

# --- 连接计算逻辑 (与 V35 相同) ---

# 辅助函数
def _find_best_single_link(dev1_copy, dev2_copy):
    """辅助函数：查找两个设备副本之间最高优先级的单个可用连接"""
    if dev1_copy.type in ['MicroN UHD', 'HorizoN'] and dev2_copy.type in ['MicroN UHD', 'HorizoN']:
        port1 = dev1_copy.get_specific_available_port('LC')
        if port1:
            port2 = dev2_copy.get_specific_available_port('LC')
            if port2:
                dev1_copy.use_specific_port(port1, dev2_copy.name); dev2_copy.use_specific_port(port2, dev1_copy.name)
                return port1, port2, 'LC-LC (100G)'
    if dev1_copy.type in ['MicroN UHD', 'HorizoN'] and dev2_copy.type in ['MicroN UHD', 'HorizoN']:
        port1 = dev1_copy.get_specific_available_port('MPO')
        if port1:
            port2 = dev2_copy.get_specific_available_port('MPO')
            if port2:
                dev1_copy.use_specific_port(port1, dev2_copy.name); dev2_copy.use_specific_port(port2, dev1_copy.name)
                return port1, port2, 'MPO-MPO (25G)'
    if dev1_copy.type == 'MicroN' and dev2_copy.type == 'MicroN':
        port1 = dev1_copy.get_specific_available_port('SFP')
        if port1:
            port2 = dev2_copy.get_specific_available_port('SFP')
            if port2:
                dev1_copy.use_specific_port(port1, dev2_copy.name); dev2_copy.use_specific_port(port2, dev1_copy.name)
                return port1, port2, 'SFP-SFP (10G)'
    uhd_dev, micron_dev = (None, None); original_dev1 = dev1_copy
    if dev1_copy.type in ['MicroN UHD', 'HorizoN'] and dev2_copy.type == 'MicroN': uhd_dev, micron_dev = dev1_copy, dev2_copy
    elif dev2_copy.type in ['MicroN UHD', 'HorizoN'] and dev1_copy.type == 'MicroN': uhd_dev, micron_dev = dev2_copy, dev1_copy
    if uhd_dev and micron_dev:
        port_uhd = uhd_dev.get_specific_available_port('MPO')
        if port_uhd:
            port_micron = micron_dev.get_specific_available_port('SFP')
            if port_micron:
                uhd_dev.use_specific_port(port_uhd, micron_dev.name); micron_dev.use_specific_port(port_micron, uhd_dev.name)
                if original_dev1 == uhd_dev: return port_uhd, port_micron, 'MPO-SFP (10G)'
                else: return port_micron, port_uhd, 'MPO-SFP (10G)'
    return None, None, None

# Mesh 计算函数
def calculate_mesh_connections(devices):
    """计算 Mesh 连接 - V34 改进版"""
    if len(devices) < 2: return [], {}
    connections = []; temp_devices = [copy.deepcopy(dev) for dev in devices]
    for d in temp_devices: d.reset_ports()
    device_map = {dev.id: dev for dev in temp_devices}
    all_pairs_ids = list(itertools.combinations([d.id for d in temp_devices], 2))
    connected_once_pairs = set()
    print("Phase 1: 尝试为每个设备对建立第一条连接...")
    made_progress_phase1 = True
    while made_progress_phase1:
        made_progress_phase1 = False
        for dev1_id, dev2_id in all_pairs_ids:
            pair_key = tuple(sorted((dev1_id, dev2_id)))
            if pair_key not in connected_once_pairs:
                dev1_copy = device_map[dev1_id]; dev2_copy = device_map[dev2_id]
                original_dev1 = next(d for d in devices if d.id == dev1_id); original_dev2 = next(d for d in devices if d.id == dev2_id)
                port1, port2, conn_type = _find_best_single_link(dev1_copy, dev2_copy)
                if port1 and port2:
                    if dev1_copy.id == dev1_id: connections.append((original_dev1, port1, original_dev2, port2, conn_type))
                    else: connections.append((original_dev2, port2, original_dev1, port1, conn_type))
                    connected_once_pairs.add(pair_key); made_progress_phase1 = True
        if not made_progress_phase1: break
    print(f"Phase 1 完成. 建立了 {len(connections)} 条初始连接。")
    print("Phase 2: 填充剩余端口...")
    connection_made_in_full_pass_phase2 = True
    while connection_made_in_full_pass_phase2:
        connection_made_in_full_pass_phase2 = False
        random.shuffle(all_pairs_ids)
        for dev1_id, dev2_id in all_pairs_ids:
            dev1_copy = device_map[dev1_id]; dev2_copy = device_map[dev2_id]
            original_dev1 = next(d for d in devices if d.id == dev1_id); original_dev2 = next(d for d in devices if d.id == dev2_id)
            port1, port2, conn_type = _find_best_single_link(dev1_copy, dev2_copy)
            if port1 and port2:
                if dev1_copy.id == dev1_id: connections.append((original_dev1, port1, original_dev2, port2, conn_type))
                else: connections.append((original_dev2, port2, original_dev1, port1, conn_type))
                connection_made_in_full_pass_phase2 = True
    print(f"Phase 2 完成. 总连接数: {len(connections)}")
    return connections, device_map

# 环形计算函数
def calculate_ring_connections(devices):
    if len(devices) < 2: return [], {}
    if len(devices) == 2: return calculate_mesh_connections(devices)
    connections = []; temp_devices = [copy.deepcopy(dev) for dev in devices]
    for d in temp_devices: d.reset_ports()
    device_map = {dev.id: dev for dev in temp_devices}
    sorted_dev_ids = sorted([d.id for d in devices]); num_devices = len(sorted_dev_ids)
    link_established = [False] * num_devices
    for i in range(num_devices):
        dev1_id = sorted_dev_ids[i]; dev2_id = sorted_dev_ids[(i + 1) % num_devices]
        dev1_copy = device_map[dev1_id]; dev2_copy = device_map[dev2_id]
        original_dev1 = next(d for d in devices if d.id == dev1_id); original_dev2 = next(d for d in devices if d.id == dev2_id)
        port1, port2, conn_type = _find_best_single_link(dev1_copy, dev2_copy)
        if port1 and port2:
            if dev1_copy.id == dev1_id: connections.append((original_dev1, port1, original_dev2, port2, conn_type))
            else: connections.append((original_dev2, port2, original_dev1, port1, conn_type))
            link_established[i] = True
        else: print(f"警告: 无法在 {original_dev1.name} 和 {original_dev2.name} 之间建立环形连接段。")
    if not all(link_established): print("警告: 未能完成完整的环形连接。")
    return connections, device_map

# 内部 Mesh 填充辅助函数
def _fill_connections_mesh_style(devices_current_state):
    """辅助函数：使用 Mesh 逻辑填充给定设备状态下的剩余连接。"""
    if len(devices_current_state) < 2: return []
    new_connections = []
    temp_devices = [copy.deepcopy(dev) for dev in devices_current_state]
    device_map = {dev.id: dev for dev in temp_devices}
    all_pairs_ids = list(itertools.combinations([d.id for d in temp_devices], 2))
    connection_made_in_full_pass = True
    while connection_made_in_full_pass:
        connection_made_in_full_pass = False
        random.shuffle(all_pairs_ids)
        for dev1_id, dev2_id in all_pairs_ids:
            dev1_copy = device_map[dev1_id]; dev2_copy = device_map[dev2_id]
            port1, port2, conn_type = _find_best_single_link(dev1_copy, dev2_copy)
            if port1 and port2:
                original_dev1 = next(d for d in devices_current_state if d.id == dev1_id)
                original_dev2 = next(d for d in devices_current_state if d.id == dev2_id)
                if dev1_copy.id == dev1_id: new_connections.append((original_dev1, port1, original_dev2, port2, conn_type))
                else: new_connections.append((original_dev2, port2, original_dev1, port1, conn_type))
                connection_made_in_full_pass = True
    return new_connections

# 内部 Ring 填充辅助函数
def _fill_connections_ring_style(devices_current_state):
    """辅助函数：使用 Ring 逻辑填充给定设备状态下的剩余连接。"""
    if len(devices_current_state) < 2: return []
    new_connections = []
    temp_devices = [copy.deepcopy(dev) for dev in devices_current_state]
    device_map = {dev.id: dev for dev in temp_devices}
    sorted_dev_ids = sorted([d.id for d in temp_devices]); num_devices = len(sorted_dev_ids)
    connection_made_in_full_pass = True
    while connection_made_in_full_pass:
        connection_made_in_full_pass = False
        for i in range(num_devices):
            dev1_id = sorted_dev_ids[i]; dev2_id = sorted_dev_ids[(i + 1) % num_devices]
            dev1_copy = device_map[dev1_id]; dev2_copy = device_map[dev2_id]
            original_dev1 = next(d for d in devices_current_state if d.id == dev1_id)
            original_dev2 = next(d for d in devices_current_state if d.id == dev2_id)
            port1, port2, conn_type = _find_best_single_link(dev1_copy, dev2_copy)
            if port1 and port2:
                if dev1_copy.id == dev1_id: new_connections.append((original_dev1, port1, original_dev2, port2, conn_type))
                else: new_connections.append((original_dev2, port2, original_dev1, port1, conn_type))
                connection_made_in_full_pass = True
    return new_connections
# --- 结束连接计算逻辑 ---


# --- Matplotlib Canvas Widget ---
class MplCanvas(FigureCanvas):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super().__init__(self.fig)
        self.setParent(parent)
        FigureCanvas.setSizePolicy(self, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        FigureCanvas.updateGeometry(self)

    def plot_topology(self, devices, connections, layout_algorithm='spring', fixed_pos=None, selected_node_id=None):
        self.axes.cla()
        if not devices: self.axes.text(0.5, 0.5, '无设备数据', ha='center', va='center'); self.draw(); return None, None
        chinese_font_name = find_chinese_font(); font_prop = None; current_font_family = 'sans-serif'
        if chinese_font_name:
            try: plt.rcParams['font.sans-serif'] = [chinese_font_name] + plt.rcParams.get('font.sans-serif', []); font_prop = font_manager.FontProperties(family=chinese_font_name); current_font_family = chinese_font_name; plt.rcParams['axes.unicode_minus'] = False
            except Exception as e: print(f"警告: 设置 Matplotlib 字体 '{chinese_font_name}' 失败: {e}"); chinese_font_name = None
        if not chinese_font_name: plt.rcParams['font.sans-serif'] = ['sans-serif']; plt.rcParams['axes.unicode_minus'] = False; font_prop = font_manager.FontProperties(); current_font_family = 'sans-serif'
        G = nx.Graph(); node_ids = [dev.id for dev in devices]; node_colors = []; node_labels = {}; node_alphas = []
        highlight_color = 'yellow'; default_alpha = 0.9; dimmed_alpha = 0.3
        for dev in devices:
            G.add_node(dev.id); node_labels[dev.id] = f"{dev.name}\n({dev.type})"
            base_color = 'grey';
            if dev.type == 'MicroN UHD': base_color = 'skyblue'
            elif dev.type == 'HorizoN': base_color = 'lightcoral'
            elif dev.type == 'MicroN': base_color = 'lightgreen'
            if selected_node_id is not None:
                if dev.id == selected_node_id: node_colors.append(highlight_color); node_alphas.append(default_alpha)
                else: is_neighbor = any((conn[0].id == selected_node_id and conn[2].id == dev.id) or (conn[0].id == dev.id and conn[2].id == selected_node_id) for conn in connections); node_colors.append(base_color); node_alphas.append(default_alpha if is_neighbor else dimmed_alpha)
            else: node_colors.append(base_color); node_alphas.append(default_alpha)
        edge_labels, edge_counts = {}, {}; highlighted_edges = set()
        if connections:
            for conn in connections:
                dev1, _, dev2, _, conn_type = conn
                if dev1.id in node_ids and dev2.id in node_ids:
                    edge_key = tuple(sorted((dev1.id, dev2.id))); G.add_edge(dev1.id, dev2.id)
                    if edge_key not in edge_counts: edge_counts[edge_key] = {}
                    base_conn_type = conn_type.split(' ')[0]
                    if base_conn_type not in edge_counts[edge_key]: edge_counts[edge_key][base_conn_type] = {'count': 0, 'details': conn_type}
                    edge_counts[edge_key][base_conn_type]['count'] += 1
                    if selected_node_id is not None and (dev1.id == selected_node_id or dev2.id == selected_node_id): highlighted_edges.add(edge_key)
            for edge_key, type_groups in edge_counts.items(): label_parts = [f"{data['details']} x{data['count']}" for base_type, data in type_groups.items()]; edge_labels[edge_key] = "\n".join(label_parts)
        pos = None
        if not G: print("DIAG (Plot): 图为空，不计算布局。"); self.axes.text(0.5, 0.5, '无连接数据', ha='center', va='center', fontproperties=font_prop); self.draw(); return self.fig, None
        if fixed_pos:
            # --- V36 Fix: Assign current_node_ids before comparison ---
            current_node_ids = set(G.nodes()) # <--- 修正：在这里赋值
            stored_node_ids = set(fixed_pos.keys())
            if current_node_ids == stored_node_ids: # 现在比较是安全的
                 pos = fixed_pos
            else:
                 print("DIAG (Plot): 节点已更改，重新计算布局。")
                 fixed_pos = None # 强制重新计算
            # --- End V36 Fix ---
        if not pos:
            try:
                if layout_algorithm == 'circular': pos = nx.circular_layout(G)
                elif layout_algorithm == 'kamada-kawai': pos = nx.kamada_kawai_layout(G)
                elif layout_algorithm == 'random': pos = nx.random_layout(G, seed=42)
                elif layout_algorithm == 'shell':
                    shells = []; types_present = sorted(list(set(dev.type for dev in devices)))
                    for t in types_present: shells.append([dev.id for dev in devices if dev.type == t])
                    if len(shells) < 2: pos = nx.spring_layout(G, seed=42, k=0.8)
                    else: pos = nx.shell_layout(G, nlist=shells)
                else: pos = nx.spring_layout(G, seed=42, k=0.8)
            except Exception as e: print(f"警告: 计算布局 '{layout_algorithm}' 时出错: {e}. 使用 spring 布局回退。"); pos = nx.spring_layout(G, seed=42, k=0.8)
        nx.draw_networkx_nodes(G, pos, node_color=node_colors, node_size=3500, ax=self.axes, alpha=node_alphas)
        nx.draw_networkx_labels(G, pos, labels=node_labels, font_size=9, ax=self.axes, font_family=current_font_family)
        if connections and G.edges():
            unique_edges = list(G.edges()); edge_colors = []; edge_widths = []; edge_alphas = []
            highlight_edge_width = 2.5; default_edge_width = 1.5; dimmed_edge_alpha = 0.15; default_edge_alpha = 0.7
            for u, v in unique_edges:
                edge_key = tuple(sorted((u, v))); is_highlighted = edge_key in highlighted_edges; is_selected_node_present = selected_node_id is not None
                color_found = 'black'
                if edge_key in edge_counts:
                    first_base_type = next(iter(edge_counts[edge_key]))
                    if 'LC-LC' in first_base_type: color_found = 'blue'
                    elif 'MPO-MPO' in first_base_type: color_found = 'red'
                    elif 'MPO-SFP' in first_base_type: color_found = 'orange'
                    elif 'SFP-SFP' in first_base_type: color_found = 'purple'
                edge_colors.append(color_found); edge_widths.append(highlight_edge_width if is_highlighted else default_edge_width)
                edge_alphas.append(default_edge_alpha if (not is_selected_node_present or is_highlighted) else dimmed_edge_alpha)
            nx.draw_networkx_edges(G, pos, edgelist=unique_edges, edge_color=edge_colors, width=edge_widths, alpha=edge_alphas, ax=self.axes, arrows=False)
            edge_label_colors = {}; dimmed_label_color = 'lightgrey'; default_label_color = 'black'
            for edge, label in edge_labels.items():
                 edge_key = tuple(sorted(edge)); is_highlighted = edge_key in highlighted_edges; is_selected_node_present = selected_node_id is not None
                 color = default_label_color if (not is_selected_node_present or is_highlighted) else dimmed_label_color
                 edge_label_colors[edge_key] = color
            nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, font_size=7, ax=self.axes, label_pos=0.5, rotate=False, font_family=current_font_family, font_color=default_label_color)
        self.axes.set_title("网络连接拓扑图", fontproperties=font_prop); self.axes.axis('off')
        legend_elements = [
            plt.Line2D([0], [0], marker='o', color='w', label='MicroN UHD', markerfacecolor='skyblue', markersize=10),
            plt.Line2D([0], [0], marker='o', color='w', label='HorizoN', markerfacecolor='lightcoral', markersize=10),
            plt.Line2D([0], [0], marker='o', color='w', label='MicroN', markerfacecolor='lightgreen', markersize=10),
            plt.Line2D([0], [0], color='blue', lw=2, label='LC-LC (100G)'), plt.Line2D([0], [0], color='red', lw=2, label='MPO-MPO (25G)'),
            plt.Line2D([0], [0], color='orange', lw=2, label='MPO-SFP (10G)'), plt.Line2D([0], [0], color='purple', lw=2, label='SFP-SFP (10G)')
        ]
        legend_prop_small = copy.copy(font_prop); legend_prop_small.set_size('small')
        self.axes.legend(handles=legend_elements, loc='best', prop=legend_prop_small)
        self.fig.tight_layout(); self.draw_idle()
        return self.fig, pos
# --- 结束 Matplotlib Canvas ---

# --- 辅助函数 (查找字体，与 V35 相同) ---
def find_chinese_font():
    font_names = ['SimHei', 'Microsoft YaHei', 'PingFang SC', 'Heiti SC', 'STHeiti', 'Noto Sans CJK SC', 'WenQuanYi Micro Hei', 'sans-serif']
    font_paths = font_manager.findSystemFonts(fontpaths=None, fontext='ttf')
    found_fonts = {}
    for font_path in font_paths:
        try: prop = font_manager.FontProperties(fname=font_path); font_name = prop.get_name()
        except RuntimeError: continue
        except Exception: continue
        if font_name in font_names and font_name not in found_fonts: found_fonts[font_name] = font_path
    for name in ['PingFang SC', 'Microsoft YaHei', 'SimHei']:
         if name in found_fonts: return name
    for name in font_names:
        if name in found_fonts: return name
    print("警告: 未找到特定中文字体。")
    return None

# --- 主窗口 (与 V35 相同) ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MediorNet TDM 连接计算器 V36 (PySide6)") # <--- 版本号更新
        self.setGeometry(100, 100, 1100, 800)
        self.devices = []
        self.connections_result = []
        self.fig = None
        self.node_positions = None
        self.device_id_counter = 0
        self.selected_node_id = None
        font_families = []; os_system = platform.system()
        if os_system == "Windows": font_families.extend(["Microsoft YaHei", "SimHei"])
        elif os_system == "Darwin": font_families.append("PingFang SC")
        font_families.extend(["Noto Sans CJK SC", "WenQuanYi Micro Hei", "sans-serif"])
        self.chinese_font = QFont(); self.chinese_font.setFamilies(font_families); self.chinese_font.setPointSize(10)
        # --- 主布局与控件 (与 V35 相同) ---
        main_widget = QWidget(); self.setCentralWidget(main_widget); main_layout = QHBoxLayout(main_widget)
        main_splitter = QSplitter(Qt.Orientation.Horizontal); main_layout.addWidget(main_splitter)
        left_panel = QFrame(); left_panel.setFrameShape(QFrame.Shape.StyledPanel);
        left_layout = QVBoxLayout(left_panel); main_splitter.addWidget(left_panel)
        add_group = QFrame(); add_group.setObjectName("addDeviceGroup")
        add_group_layout = QGridLayout(add_group); add_group_layout.setContentsMargins(10, 15, 10, 10); add_group_layout.setVerticalSpacing(8)
        add_title = QLabel("<b>添加新设备</b>"); add_title.setFont(QFont(self.chinese_font.family(), 11)); add_group_layout.addWidget(add_title, 0, 0, 1, 2, alignment=Qt.AlignmentFlag.AlignCenter)
        add_group_layout.addWidget(QLabel("类型:"), 1, 0); self.device_type_combo = QComboBox(); self.device_type_combo.addItems(['MicroN UHD', 'HorizoN', 'MicroN']); self.device_type_combo.setFont(self.chinese_font); self.device_type_combo.currentIndexChanged.connect(self.update_port_entries); add_group_layout.addWidget(self.device_type_combo, 1, 1)
        add_group_layout.addWidget(QLabel("名称:"), 2, 0); self.device_name_entry = QLineEdit(); self.device_name_entry.setFont(self.chinese_font); add_group_layout.addWidget(self.device_name_entry, 2, 1)
        self.mpo_label = QLabel("MPO 端口:"); add_group_layout.addWidget(self.mpo_label, 3, 0); self.mpo_entry = QLineEdit("4"); self.mpo_entry.setFont(self.chinese_font); add_group_layout.addWidget(self.mpo_entry, 3, 1)
        self.lc_label = QLabel("LC 端口:"); add_group_layout.addWidget(self.lc_label, 4, 0); self.lc_entry = QLineEdit("2"); self.lc_entry.setFont(self.chinese_font); add_group_layout.addWidget(self.lc_entry, 4, 1)
        self.sfp_label = QLabel("SFP+ 端口:"); self.sfp_entry = QLineEdit("8"); self.sfp_entry.setFont(self.chinese_font); add_group_layout.addWidget(self.sfp_label, 5, 0); add_group_layout.addWidget(self.sfp_entry, 5, 1); self.sfp_label.hide(); self.sfp_entry.hide()
        self.add_button = QPushButton("添加设备"); self.add_button.setFont(self.chinese_font); self.add_button.clicked.connect(self.add_device); add_group_layout.addWidget(self.add_button, 6, 0, 1, 2, alignment=Qt.AlignmentFlag.AlignCenter)
        left_layout.addWidget(add_group)
        list_group = QFrame(); list_group.setObjectName("listGroup"); list_group_layout = QVBoxLayout(list_group); list_group_layout.setContentsMargins(10, 15, 10, 10)
        filter_layout = QHBoxLayout(); filter_layout.addWidget(QLabel("过滤:", font=self.chinese_font)); self.device_filter_entry = QLineEdit(); self.device_filter_entry.setFont(self.chinese_font); self.device_filter_entry.setPlaceholderText("按名称或类型过滤..."); self.device_filter_entry.textChanged.connect(self.filter_device_table); filter_layout.addWidget(self.device_filter_entry); list_group_layout.addLayout(filter_layout)
        self.device_tablewidget = QTableWidget(); self.device_tablewidget.setFont(self.chinese_font); self.device_tablewidget.setColumnCount(6); self.device_tablewidget.setHorizontalHeaderLabels(["名称", "类型", "MPO", "LC", "SFP+", "连接数(估)"]); self.device_tablewidget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); self.device_tablewidget.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection); self.device_tablewidget.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers);
        self.device_tablewidget.setSortingEnabled(True)
        header = self.device_tablewidget.horizontalHeader(); header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive); header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.device_tablewidget.setColumnWidth(1, 90); self.device_tablewidget.setColumnWidth(2, 50); self.device_tablewidget.setColumnWidth(3, 50); self.device_tablewidget.setColumnWidth(4, 50); self.device_tablewidget.setColumnWidth(5, 80)
        self.device_tablewidget.itemDoubleClicked.connect(self.show_device_details_from_table); list_group_layout.addWidget(self.device_tablewidget)
        device_op_layout = QHBoxLayout(); self.remove_button = QPushButton("移除选中"); self.remove_button.setFont(self.chinese_font); self.remove_button.clicked.connect(self.remove_device); self.clear_button = QPushButton("清空所有"); self.clear_button.setFont(self.chinese_font); self.clear_button.clicked.connect(self.clear_all_devices); device_op_layout.addWidget(self.remove_button); device_op_layout.addWidget(self.clear_button); list_group_layout.addLayout(device_op_layout)
        left_layout.addWidget(list_group)
        file_group = QFrame(); file_group.setObjectName("fileGroup"); file_group_layout = QGridLayout(file_group); file_group_layout.setContentsMargins(10, 15, 10, 10); file_title = QLabel("<b>文件操作</b>"); file_title.setFont(QFont(self.chinese_font.family(), 11)); file_group_layout.addWidget(file_title, 0, 0, 1, 2, alignment=Qt.AlignmentFlag.AlignCenter)
        self.save_button = QPushButton("保存配置"); self.save_button.setFont(self.chinese_font); self.save_button.clicked.connect(self.save_config); self.load_button = QPushButton("加载配置"); self.load_button.setFont(self.chinese_font); self.load_button.clicked.connect(self.load_config); file_group_layout.addWidget(self.save_button, 1, 0); file_group_layout.addWidget(self.load_button, 1, 1)
        self.export_list_button = QPushButton("导出列表"); self.export_list_button.setFont(self.chinese_font); self.export_list_button.clicked.connect(self.export_connections); self.export_list_button.setEnabled(False); self.export_topo_button = QPushButton("导出拓扑图"); self.export_topo_button.setFont(self.chinese_font); self.export_topo_button.clicked.connect(self.export_topology); self.export_topo_button.setEnabled(False); file_group_layout.addWidget(self.export_list_button, 2, 0); file_group_layout.addWidget(self.export_topo_button, 2, 1)
        left_layout.addWidget(file_group); left_layout.addStretch()
        right_panel = QFrame(); right_panel.setFrameShape(QFrame.Shape.StyledPanel)
        right_layout = QVBoxLayout(right_panel); main_splitter.addWidget(right_panel)
        calculate_control_frame = QFrame(); calculate_control_frame.setObjectName("calculateControlFrame")
        calculate_control_layout = QHBoxLayout(calculate_control_frame)
        calculate_control_layout.setContentsMargins(10, 5, 10, 5)
        calculate_control_layout.addWidget(QLabel("计算模式:", font=self.chinese_font))
        self.topology_mode_combo = QComboBox(); self.topology_mode_combo.setFont(self.chinese_font); self.topology_mode_combo.addItems(["Mesh", "环形"]); calculate_control_layout.addWidget(self.topology_mode_combo)
        calculate_control_layout.addWidget(QLabel("布局:", font=self.chinese_font))
        self.layout_combo = QComboBox(); self.layout_combo.setFont(self.chinese_font); self.layout_combo.addItems(["Spring", "Circular", "Kamada-Kawai", "Random", "Shell"]); self.layout_combo.currentIndexChanged.connect(self.on_layout_change); calculate_control_layout.addWidget(self.layout_combo)
        self.calculate_button = QPushButton("计算连接"); self.calculate_button.setFont(self.chinese_font); self.calculate_button.clicked.connect(self.calculate_and_display); calculate_control_layout.addWidget(self.calculate_button)
        self.fill_mesh_button = QPushButton("填充 (Mesh)"); self.fill_mesh_button.setFont(self.chinese_font); self.fill_mesh_button.setEnabled(False); self.fill_mesh_button.clicked.connect(self.fill_remaining_mesh); calculate_control_layout.addWidget(self.fill_mesh_button)
        self.fill_ring_button = QPushButton("填充 (环形)"); self.fill_ring_button.setFont(self.chinese_font); self.fill_ring_button.setEnabled(False); self.fill_ring_button.clicked.connect(self.fill_remaining_ring); calculate_control_layout.addWidget(self.fill_ring_button)
        calculate_control_layout.addStretch(); right_layout.addWidget(calculate_control_frame)
        self.tab_widget = QTabWidget(); self.tab_widget.setFont(self.chinese_font); right_layout.addWidget(self.tab_widget)
        self.connections_tab = QWidget(); connections_layout = QVBoxLayout(self.connections_tab); self.connections_textedit = QTextEdit(); self.connections_textedit.setFont(self.chinese_font); self.connections_textedit.setReadOnly(True); connections_layout.addWidget(self.connections_textedit); self.tab_widget.addTab(self.connections_tab, "连接列表")
        self.topology_tab = QWidget(); topology_layout = QVBoxLayout(self.topology_tab); self.mpl_canvas = MplCanvas(self.topology_tab, width=8, height=6, dpi=100); topology_layout.addWidget(self.mpl_canvas); self.tab_widget.addTab(self.topology_tab, "拓扑图")
        self.mpl_canvas.mpl_connect('button_press_event', self.on_canvas_click)
        self.edit_tab = QWidget(); edit_main_layout = QVBoxLayout(self.edit_tab)
        add_manual_group = QFrame(); add_manual_group.setObjectName("addManualGroup"); add_manual_group.setFrameShape(QFrame.Shape.StyledPanel); add_manual_layout = QGridLayout(add_manual_group); add_manual_layout.setContentsMargins(10,10,10,10); add_manual_layout.addWidget(QLabel("<b>添加手动连接</b>", font=self.chinese_font), 0, 0, 1, 4, alignment=Qt.AlignmentFlag.AlignCenter)
        add_manual_layout.addWidget(QLabel("设备 1:", font=self.chinese_font), 1, 0); self.edit_dev1_combo = QComboBox(); self.edit_dev1_combo.setFont(self.chinese_font); add_manual_layout.addWidget(self.edit_dev1_combo, 1, 1); add_manual_layout.addWidget(QLabel("端口 1:", font=self.chinese_font), 1, 2); self.edit_port1_combo = QComboBox(); self.edit_port1_combo.setFont(self.chinese_font); add_manual_layout.addWidget(self.edit_port1_combo, 1, 3)
        add_manual_layout.addWidget(QLabel("设备 2:", font=self.chinese_font), 2, 0); self.edit_dev2_combo = QComboBox(); self.edit_dev2_combo.setFont(self.chinese_font); add_manual_layout.addWidget(self.edit_dev2_combo, 2, 1); add_manual_layout.addWidget(QLabel("端口 2:", font=self.chinese_font), 2, 2); self.edit_port2_combo = QComboBox(); self.edit_port2_combo.setFont(self.chinese_font); add_manual_layout.addWidget(self.edit_port2_combo, 2, 3)
        self.add_manual_button = QPushButton("添加连接"); self.add_manual_button.setFont(self.chinese_font); self.add_manual_button.clicked.connect(self.add_manual_connection); add_manual_layout.addWidget(self.add_manual_button, 3, 0, 1, 4, alignment=Qt.AlignmentFlag.AlignCenter)
        self.edit_dev1_combo.currentIndexChanged.connect(lambda: self._populate_edit_port_combos(self.edit_dev1_combo, self.edit_port1_combo)); self.edit_dev2_combo.currentIndexChanged.connect(lambda: self._populate_edit_port_combos(self.edit_dev2_combo, self.edit_port2_combo)); edit_main_layout.addWidget(add_manual_group)
        remove_manual_group = QFrame(); remove_manual_group.setObjectName("removeManualGroup"); remove_manual_group.setFrameShape(QFrame.Shape.StyledPanel); remove_manual_layout = QVBoxLayout(remove_manual_group); remove_manual_layout.setContentsMargins(10,10,10,10); remove_title = QLabel("<b>移除现有连接</b> (选中下方列表中的连接进行移除)"); remove_title.setFont(self.chinese_font); remove_manual_layout.addWidget(remove_title)
        filter_conn_layout = QHBoxLayout()
        filter_conn_layout.addWidget(QLabel("类型过滤:", font=self.chinese_font))
        self.conn_filter_type_combo = QComboBox()
        self.conn_filter_type_combo.setFont(self.chinese_font)
        self.conn_filter_type_combo.addItems(["所有类型", "LC-LC (100G)", "MPO-MPO (25G)", "SFP-SFP (10G)", "MPO-SFP (10G)"])
        self.conn_filter_type_combo.currentIndexChanged.connect(self.filter_connection_list)
        filter_conn_layout.addWidget(self.conn_filter_type_combo)
        filter_conn_layout.addSpacing(15)
        filter_conn_layout.addWidget(QLabel("设备过滤:", font=self.chinese_font))
        self.conn_filter_device_entry = QLineEdit()
        self.conn_filter_device_entry.setFont(self.chinese_font)
        self.conn_filter_device_entry.setPlaceholderText("按设备名称过滤...")
        self.conn_filter_device_entry.textChanged.connect(self.filter_connection_list)
        filter_conn_layout.addWidget(self.conn_filter_device_entry)
        remove_manual_layout.insertLayout(1, filter_conn_layout)
        self.manual_connection_list = QListWidget(); self.manual_connection_list.setFont(self.chinese_font); self.manual_connection_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection); remove_manual_layout.addWidget(self.manual_connection_list); self.remove_manual_button = QPushButton("移除选中连接"); self.remove_manual_button.setFont(self.chinese_font); self.remove_manual_button.clicked.connect(self.remove_manual_connection); self.remove_manual_button.setEnabled(False); remove_manual_layout.addWidget(self.remove_manual_button, alignment=Qt.AlignmentFlag.AlignCenter); edit_main_layout.addWidget(remove_manual_group); self.tab_widget.addTab(self.edit_tab, "手动编辑")
        main_splitter.setSizes([400, 700])
        main_splitter.setStretchFactor(1, 1)
        self.update_port_entries()

    # --- 方法 (与 V35 相同) ---
    @Slot()
    def update_port_entries(self):
        selected_type = self.device_type_combo.currentText()
        is_micron = selected_type == 'MicroN'; is_uhd_horizon = selected_type in ['MicroN UHD', 'HorizoN']
        self.mpo_label.setVisible(is_uhd_horizon); self.mpo_entry.setVisible(is_uhd_horizon)
        self.lc_label.setVisible(is_uhd_horizon); self.lc_entry.setVisible(is_uhd_horizon)
        self.sfp_label.setVisible(is_micron); self.sfp_entry.setVisible(is_micron)

    def _add_device_to_table(self, device):
        row_position = self.device_tablewidget.rowCount(); self.device_tablewidget.insertRow(row_position)
        name_item = QTableWidgetItem(device.name); name_item.setData(Qt.ItemDataRole.UserRole, device.id)
        type_item = QTableWidgetItem(device.type)
        mpo_item = NumericTableWidgetItem(str(device.mpo_total)); mpo_item.setData(Qt.ItemDataRole.UserRole + 1, device.mpo_total); mpo_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        lc_item = NumericTableWidgetItem(str(device.lc_total)); lc_item.setData(Qt.ItemDataRole.UserRole + 1, device.lc_total); lc_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        sfp_item = NumericTableWidgetItem(str(device.sfp_total)); sfp_item.setData(Qt.ItemDataRole.UserRole + 1, device.sfp_total); sfp_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        conn_val = float(f"{device.connections:.2f}")
        conn_item = NumericTableWidgetItem(f"{conn_val:.2f}"); conn_item.setData(Qt.ItemDataRole.UserRole + 1, conn_val); conn_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.device_tablewidget.setItem(row_position, 0, name_item); self.device_tablewidget.setItem(row_position, 1, type_item)
        self.device_tablewidget.setItem(row_position, 2, mpo_item); self.device_tablewidget.setItem(row_position, 3, lc_item)
        self.device_tablewidget.setItem(row_position, 4, sfp_item); self.device_tablewidget.setItem(row_position, 5, conn_item)

    def _update_device_combos(self):
        self.edit_dev1_combo.blockSignals(True); self.edit_dev2_combo.blockSignals(True)
        current_dev1_id = self.edit_dev1_combo.currentData(); current_dev2_id = self.edit_dev2_combo.currentData()
        self.edit_dev1_combo.clear(); self.edit_dev2_combo.clear()
        self.edit_dev1_combo.addItem("选择设备 1...", userData=None); self.edit_dev2_combo.addItem("选择设备 2...", userData=None)
        sorted_devices = sorted(self.devices, key=lambda dev: dev.name)
        idx1_to_select, idx2_to_select = 0, 0
        for i, dev in enumerate(sorted_devices):
            item_text = f"{dev.name} ({dev.type})"; self.edit_dev1_combo.addItem(item_text, userData=dev.id); self.edit_dev2_combo.addItem(item_text, userData=dev.id)
            if dev.id == current_dev1_id: idx1_to_select = i + 1
            if dev.id == current_dev2_id: idx2_to_select = i + 1
        self.edit_dev1_combo.setCurrentIndex(idx1_to_select); self.edit_dev2_combo.setCurrentIndex(idx2_to_select)
        self.edit_dev1_combo.blockSignals(False); self.edit_dev2_combo.blockSignals(False)
        self._populate_edit_port_combos(self.edit_dev1_combo, self.edit_port1_combo)
        self._populate_edit_port_combos(self.edit_dev2_combo, self.edit_port2_combo)

    def _populate_edit_port_combos(self, device_combo, port_combo):
        port_combo.blockSignals(True)
        port_combo.clear(); port_combo.addItem("选择端口...")
        dev_id = device_combo.currentData(); port_combo.setEnabled(False)
        if dev_id is not None:
            device = next((d for d in self.devices if d.id == dev_id), None)
            if device:
                available_ports = device.get_all_available_ports()
                if available_ports: port_combo.addItems(available_ports); port_combo.setEnabled(True)
                else: port_combo.addItem("无可用端口")
        port_combo.blockSignals(False)

    @Slot()
    def add_device(self):
        dtype = self.device_type_combo.currentText(); name = self.device_name_entry.text().strip()
        if not dtype: QMessageBox.critical(self, "错误", "请选择设备类型。"); return
        if not name: QMessageBox.critical(self, "错误", "请输入设备名称。"); return
        if any(dev.name == name for dev in self.devices): QMessageBox.critical(self, "错误", f"设备名称 '{name}' 已存在。"); return
        mpo_ports_str, lc_ports_str, sfp_ports_str = self.mpo_entry.text() or "0", self.lc_entry.text() or "0", self.sfp_entry.text() or "0"
        try:
            if dtype in ['MicroN UHD', 'HorizoN']: mpo_ports, lc_ports, sfp_ports = int(mpo_ports_str), int(lc_ports_str), 0; assert mpo_ports >= 0 and lc_ports >= 0
            elif dtype == 'MicroN': sfp_ports, mpo_ports, lc_ports = int(sfp_ports_str), 0, 0; assert sfp_ports >= 0
            else: raise ValueError("无效类型")
        except (ValueError, AssertionError): QMessageBox.critical(self, "输入错误", "端口数量必须是非负整数。"); return
        self.device_id_counter += 1; new_device = Device(self.device_id_counter, name, dtype, mpo_ports, lc_ports, sfp_ports)
        self.devices.append(new_device); self._add_device_to_table(new_device); self.device_name_entry.clear();
        self._update_device_combos(); self.clear_results()
        self.node_positions = None; self.selected_node_id = None

    @Slot()
    def remove_device(self):
        selected_rows = sorted(list(set(index.row() for index in self.device_tablewidget.selectedIndexes())), reverse=True)
        if not selected_rows: QMessageBox.warning(self, "提示", "请先在表格中选择要移除的设备行。"); return
        ids_to_remove = {self.device_tablewidget.item(row_index, 0).data(Qt.ItemDataRole.UserRole) for row_index in selected_rows if self.device_tablewidget.item(row_index, 0)}
        self.devices = [dev for dev in self.devices if dev.id not in ids_to_remove]
        for row_index in selected_rows: self.device_tablewidget.removeRow(row_index)
        self._update_device_combos(); self.clear_results()
        self.node_positions = None; self.selected_node_id = None

    @Slot()
    def clear_all_devices(self):
        if QMessageBox.question(self, "确认", "确定要清空所有设备吗？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            self.devices = []; self.device_tablewidget.setRowCount(0); self.device_id_counter = 0;
            self._update_device_combos(); self.clear_results()
            self.node_positions = None; self.selected_node_id = None

    @Slot()
    def clear_results(self):
        self.connections_result = []; self.fig = None; self.node_positions = None; self.selected_node_id = None
        self.connections_textedit.clear(); self.manual_connection_list.clear()
        self.mpl_canvas.axes.cla(); self.mpl_canvas.axes.text(0.5, 0.5, '点击“计算”生成图形', ha='center', va='center'); self.mpl_canvas.draw()
        self.export_list_button.setEnabled(False); self.export_topo_button.setEnabled(False); self.remove_manual_button.setEnabled(False)
        self.fill_mesh_button.setEnabled(False); self.fill_ring_button.setEnabled(False)
        for dev in self.devices: dev.reset_ports()
        self._update_device_table_connections()

    @Slot(QTableWidgetItem)
    def show_device_details_from_table(self, item):
         if not item: return
         row = item.row(); name_item = self.device_tablewidget.item(row, 0)
         if not name_item: return
         dev_id = name_item.data(Qt.ItemDataRole.UserRole)
         dev = next((d for d in self.devices if d.id == dev_id), None)
         if not dev: return QMessageBox.critical(self, "错误", "无法找到所选设备的详细信息。")
         self._display_device_details_popup(dev)

    def _display_device_details_popup(self, dev):
        details = f"ID: {dev.id}\n名称: {dev.name}\n类型: {dev.type}\n"
        avail_ports = dev.get_all_available_ports()
        avail_lc_count = sum(1 for p in avail_ports if p.startswith("LC"))
        avail_sfp_count = sum(1 for p in avail_ports if p.startswith("SFP"))
        avail_mpo_ch_count = sum(1 for p in avail_ports if p.startswith("MPO"))
        if dev.type in ['MicroN UHD', 'HorizoN']: details += f"MPO 端口总数: {dev.mpo_total}\nLC 端口总数: {dev.lc_total}\n可用 MPO 子通道: {avail_mpo_ch_count}\n可用 LC 端口: {avail_lc_count}\n"
        elif dev.type == 'MicroN': details += f"SFP+ 端口总数: {dev.sfp_total}\n可用 SFP+ 端口: {avail_sfp_count}\n"
        details += f"当前连接数 (估算): {dev.connections:.2f}\n"
        if dev.port_connections:
            details += "\n端口连接详情:\n"
            lc_conns = {p: t for p, t in dev.port_connections.items() if p.startswith("LC")}
            sfp_conns = {p: t for p, t in dev.port_connections.items() if p.startswith("SFP")}
            mpo_conns_grouped = defaultdict(dict)
            for p, t in dev.port_connections.items():
                if p.startswith("MPO"): mpo_conns_grouped[p.split('-')[0]][p] = t
            if lc_conns: details += "  LC 连接:\n"; [details := details + f"    {port} -> {lc_conns[port]}\n" for port in sorted(lc_conns.keys())]
            if sfp_conns: details += "  SFP+ 连接:\n"; [details := details + f"    {port} -> {sfp_conns[port]}\n" for port in sorted(sfp_conns.keys())]
            if mpo_conns_grouped:
                details += "  MPO 连接 (Breakout):\n"
                for base_port in sorted(mpo_conns_grouped.keys()):
                    details += f"    {base_port}:\n"; [details := details + f"      {port} -> {mpo_conns_grouped[base_port][port]}\n" for port in sorted(mpo_conns_grouped[base_port].keys(), key=lambda x: int(x.split('-Ch')[-1]))]
        QMessageBox.information(self, f"设备详情 - {dev.name}", details)

    @Slot()
    def calculate_and_display(self):
        """计算连接并显示结果 (使用新的 Mesh 算法)"""
        if not self.devices: QMessageBox.information(self, "提示", "请先添加设备。"); return
        self.node_positions = None; self.selected_node_id = None
        mode = self.topology_mode_combo.currentText()
        calculated_connections, final_device_state_map = [], {}
        for dev in self.devices: dev.reset_ports()
        if mode == "Mesh": print("使用改进的 Mesh 算法进行计算..."); calculated_connections, final_device_state_map = calculate_mesh_connections(self.devices)
        elif mode == "环形": print("使用环形算法进行计算..."); calculated_connections, final_device_state_map = calculate_ring_connections(self.devices)
        else: QMessageBox.critical(self, "错误", f"未知的计算模式: {mode}"); return
        self.connections_result = calculated_connections
        for dev in self.devices:
            if dev.id in final_device_state_map:
                final_state_dev = final_device_state_map[dev.id]
                dev.connections = final_state_dev.connections
                dev.port_connections = final_state_dev.port_connections.copy()
            else: dev.reset_ports()
        self._update_connection_views(); self._update_device_table_connections()
        self._update_device_combos()
        enable_fill = bool(self.connections_result) or any(dev.get_all_available_ports() for dev in self.devices)
        self.fill_mesh_button.setEnabled(enable_fill); self.fill_ring_button.setEnabled(enable_fill)

    def _update_connection_views(self):
        """更新连接列表、手动编辑列表和拓扑图 (包含高亮)"""
        self.connections_textedit.clear()
        if self.connections_result:
            self.connections_textedit.append("<b>连接列表:</b><hr>")
            for i, conn in enumerate(self.connections_result): dev1, port1, dev2, port2, conn_type = conn; self.connections_textedit.append(f"{i+1}. {dev1.name} [{port1}] &lt;-&gt; {dev2.name} [{port2}] ({conn_type})")
            self.export_list_button.setEnabled(True)
        else: self.connections_textedit.append("无连接。"); self.export_list_button.setEnabled(False)
        self.manual_connection_list.clear()
        if self.connections_result:
            for i, conn in enumerate(self.connections_result):
                dev1, port1, dev2, port2, conn_type = conn; item_text = f"{i+1}. {dev1.name} [{port1}] <-> {dev2.name} [{port2}] ({conn_type})"
                item = QListWidgetItem(item_text); item.setData(Qt.ItemDataRole.UserRole, conn); self.manual_connection_list.addItem(item)
            self.remove_manual_button.setEnabled(True)
        else: self.remove_manual_button.setEnabled(False)
        self.filter_connection_list()
        selected_layout = self.layout_combo.currentText().lower()
        self.fig, calculated_pos = self.mpl_canvas.plot_topology(
            self.devices, self.connections_result,
            layout_algorithm=selected_layout,
            fixed_pos=self.node_positions,
            selected_node_id=self.selected_node_id
        )
        if calculated_pos is not None and self.node_positions is None: self.node_positions = calculated_pos
        self.export_topo_button.setEnabled(bool(self.fig))

    def _update_device_table_connections(self):
        """更新设备表格中的连接数估算列，并设置数字排序"""
        for row in range(self.device_tablewidget.rowCount()):
            name_item = self.device_tablewidget.item(row, 0)
            if name_item:
                dev_id = name_item.data(Qt.ItemDataRole.UserRole)
                device = next((d for d in self.devices if d.id == dev_id), None)
                if device:
                    conn_val = float(f"{device.connections:.2f}")
                    conn_item = NumericTableWidgetItem(f"{conn_val:.2f}")
                    conn_item.setData(Qt.ItemDataRole.UserRole + 1, conn_val)
                    conn_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.device_tablewidget.setItem(row, 5, conn_item)

    @Slot()
    def remove_manual_connection(self):
        selected_items = self.manual_connection_list.selectedItems()
        if not selected_items: QMessageBox.warning(self, "提示", "请在下方列表中选择要移除的连接。"); return
        items_to_remove_from_widget, connections_removed_count, connections_to_remove_data = [], 0, []
        for item in selected_items: conn_data = item.data(Qt.ItemDataRole.UserRole); connections_to_remove_data.append(conn_data); items_to_remove_from_widget.append(item)
        for conn_data in connections_to_remove_data:
            dev1_orig, port1, dev2_orig, port2, conn_type = conn_data
            try:
                found_index = -1
                for i, existing_conn in enumerate(self.connections_result):
                    if (existing_conn[0].id == dev1_orig.id and existing_conn[1] == port1 and existing_conn[2].id == dev2_orig.id and existing_conn[3] == port2) or \
                       (existing_conn[0].id == dev2_orig.id and existing_conn[1] == port2 and existing_conn[2].id == dev1_orig.id and existing_conn[3] == port1):
                        found_index = i; break
                if found_index != -1:
                    self.connections_result.pop(found_index)
                    dev1_orig.return_port(port1); dev2_orig.return_port(port2)
                    connections_removed_count += 1
                else: print(f"警告: 移除时未找到匹配项: {conn_data}")
            except Exception as e: print(f"警告: 移除连接时出错: {e} - {conn_data}")
        if connections_removed_count > 0:
             self.node_positions = None; self.selected_node_id = None
             for item in items_to_remove_from_widget:
                 row = self.manual_connection_list.row(item)
                 if row != -1: self.manual_connection_list.takeItem(row)
             self._update_connection_views(); self._update_device_table_connections()
             self._update_device_combos()
             print(f"成功移除了 {connections_removed_count} 条连接。")
             enable_fill = bool(self.connections_result) or any(dev.get_all_available_ports() for dev in self.devices)
             self.fill_mesh_button.setEnabled(enable_fill); self.fill_ring_button.setEnabled(enable_fill)
        else: print("没有连接被移除。")

    @Slot()
    def add_manual_connection(self):
        dev1_id = self.edit_dev1_combo.currentData(); dev2_id = self.edit_dev2_combo.currentData()
        port1_text = self.edit_port1_combo.currentText(); port2_text = self.edit_port2_combo.currentText()
        if dev1_id is None or dev2_id is None or port1_text == "选择端口..." or port2_text == "选择端口...": QMessageBox.warning(self, "选择不完整", "请选择两个设备和它们各自的端口。"); return
        if dev1_id == dev2_id: QMessageBox.warning(self, "选择错误", "不能将设备连接到自身。"); return
        dev1 = next((d for d in self.devices if d.id == dev1_id), None); dev2 = next((d for d in self.devices if d.id == dev2_id), None)
        if not dev1 or not dev2: QMessageBox.critical(self, "内部错误", "找不到选定的设备对象。"); return
        port1_type = "LC" if port1_text.startswith("LC") else ("SFP" if port1_text.startswith("SFP") else ("MPO" if port1_text.startswith("MPO") else "Unknown"))
        port2_type = "LC" if port2_text.startswith("LC") else ("SFP" if port2_text.startswith("SFP") else ("MPO" if port2_text.startswith("MPO") else "Unknown"))
        conn_type_str = ""; valid_connection = False
        dev1_ref, dev2_ref = dev1, dev2; port1_ref, port2_ref = port1_text, port2_text
        if port1_type == "LC" and port2_type == "LC" and dev1_ref.type in ['MicroN UHD', 'HorizoN'] and dev2_ref.type in ['MicroN UHD', 'HorizoN']: valid_connection = True; conn_type_str = "LC-LC (100G)"
        elif port1_type == "MPO" and port2_type == "MPO" and dev1_ref.type in ['MicroN UHD', 'HorizoN'] and dev2_ref.type in ['MicroN UHD', 'HorizoN']: valid_connection = True; conn_type_str = "MPO-MPO (25G)"
        elif port1_type == "SFP" and port2_type == "SFP" and dev1_ref.type == 'MicroN' and dev2_ref.type == 'MicroN': valid_connection = True; conn_type_str = "SFP-SFP (10G)"
        elif port1_type == "MPO" and port2_type == "SFP" and dev1_ref.type in ['MicroN UHD', 'HorizoN'] and dev2_ref.type == 'MicroN': valid_connection = True; conn_type_str = "MPO-SFP (10G)"
        elif port1_type == "SFP" and port2_type == "MPO" and dev1_ref.type == 'MicroN' and dev2_ref.type in ['MicroN UHD', 'HorizoN']:
             dev1, dev2 = dev2_ref, dev1_ref; port1_text, port2_text = port2_ref, port1_ref
             valid_connection = True; conn_type_str = "MPO-SFP (10G)"
        if not valid_connection: QMessageBox.warning(self, "连接无效", f"设备类型 '{dev1_ref.type}' 的端口 '{port1_type}' 与 设备类型 '{dev2_ref.type}' 的端口 '{port2_type}' 之间不允许连接。"); return
        if dev1.use_specific_port(port1_text, dev2.name):
            if dev2.use_specific_port(port2_text, dev1.name):
                self.node_positions = None; self.selected_node_id = None
                self.connections_result.append((dev1_ref, port1_ref, dev2_ref, port2_ref, conn_type_str))
                self._update_connection_views(); self._update_device_table_connections()
                self._populate_edit_port_combos(self.edit_dev1_combo, self.edit_port1_combo)
                self._populate_edit_port_combos(self.edit_dev2_combo, self.edit_port2_combo)
                print(f"成功添加手动连接: {dev1_ref.name}[{port1_ref}] <-> {dev2_ref.name}[{port2_ref}]")
                enable_fill = bool(self.connections_result) or any(dev.get_all_available_ports() for dev in self.devices)
                self.fill_mesh_button.setEnabled(enable_fill); self.fill_ring_button.setEnabled(enable_fill)
            else:
                dev1.return_port(port1_text)
                QMessageBox.critical(self, "错误", f"端口 '{port2_ref}' 在设备 '{dev2_ref.name}' 上不可用或已被占用。")
                self._populate_edit_port_combos(self.edit_dev1_combo if dev1.id == dev1_id else self.edit_dev2_combo, self.edit_port1_combo if dev1.id == dev1_id else self.edit_port2_combo)
        else: QMessageBox.critical(self, "错误", f"端口 '{port1_ref}' 在设备 '{dev1_ref.name}' 上不可用或已被占用。")

    # --- 修改：填充按钮调用新的内部辅助函数 ---
    @Slot()
    def fill_remaining_mesh(self):
        """使用 Mesh 算法填充剩余的可用端口 (V35 修正)"""
        if not self.devices: QMessageBox.information(self, "提示", "请先添加设备。"); return

        print("开始填充剩余连接 (Mesh)...")
        # 调用新的辅助函数，传递当前设备状态
        new_connections = _fill_connections_mesh_style(self.devices) # <--- 使用新辅助函数

        if new_connections:
            print(f"找到 {len(new_connections)} 条新 Mesh 连接可以添加。")
            self.connections_result.extend(new_connections)
            print("更新设备状态...")
            # 更新实际设备对象的状态
            for conn in new_connections:
                dev1, port1, dev2, port2, _ = conn
                # 在原始设备对象上标记端口已用
                if not dev1.use_specific_port(port1, dev2.name):
                     print(f"严重警告: Mesh 填充同步状态时未能使用端口 {dev1.name}[{port1}]")
                if not dev2.use_specific_port(port2, dev1.name):
                     print(f"严重警告: Mesh 填充同步状态时未能使用端口 {dev2.name}[{port2}]")

            print("更新视图...")
            self.selected_node_id = None # 清除选中状态
            # 填充后不重置布局，保持现有位置
            self._update_connection_views()
            self._update_device_table_connections()
            self._update_device_combos()
            QMessageBox.information(self, "填充完成", f"成功添加了 {len(new_connections)} 条新 Mesh 连接。")
        else:
            QMessageBox.information(self, "填充完成", "没有找到更多可以建立的 Mesh 连接。")
            print("未找到可填充的新 Mesh 连接。")

        # 填充后禁用填充按钮
        self.fill_mesh_button.setEnabled(False)
        self.fill_ring_button.setEnabled(False)

    @Slot()
    def fill_remaining_ring(self):
        """使用 Ring 算法填充剩余的可用端口 (V35 修正)"""
        if not self.devices: QMessageBox.information(self, "提示", "请先添加设备。"); return

        print("开始填充剩余连接 (环形)...")
        # 调用新的辅助函数
        new_connections = _fill_connections_ring_style(self.devices) # <--- 使用新辅助函数

        if new_connections:
            print(f"找到 {len(new_connections)} 条新环形连接段可以添加。")
            self.connections_result.extend(new_connections)
            print("更新设备状态...")
            # 更新实际设备对象的状态
            for conn in new_connections:
                dev1, port1, dev2, port2, _ = conn
                if not dev1.use_specific_port(port1, dev2.name):
                     print(f"严重警告: 环形填充同步状态时未能使用端口 {dev1.name}[{port1}]")
                if not dev2.use_specific_port(port2, dev1.name):
                     print(f"严重警告: 环形填充同步状态时未能使用端口 {dev2.name}[{port2}]")

            print("更新视图...")
            self.selected_node_id = None # 清除选中状态
            # 填充后不重置布局
            self._update_connection_views()
            self._update_device_table_connections()
            self._update_device_combos()
            QMessageBox.information(self, "填充完成", f"成功添加了 {len(new_connections)} 条新环形连接段。")
        else:
            QMessageBox.information(self, "填充完成", "没有找到更多可以建立的环形连接段。")
            print("未找到可填充的新环形连接段。")

        # 填充后禁用填充按钮
        self.fill_mesh_button.setEnabled(False)
        self.fill_ring_button.setEnabled(False)
    # --- 结束修改 ---

    @Slot()
    def on_layout_change(self):
        if self.connections_result:
            print("DIAG: Layout selection changed, resetting stored positions.")
            self.node_positions = None; self.selected_node_id = None
            self._update_connection_views()

    @Slot()
    def save_config(self):
        if not self.devices: QMessageBox.warning(self, "提示", "设备列表为空。"); return
        filepath, _ = QFileDialog.getSaveFileName(self, "保存设备配置", "", "JSON 文件 (*.json);;所有文件 (*)")
        if not filepath: return
        config_data = [dev.to_dict() for dev in self.devices]
        try:
            with open(filepath, 'w', encoding='utf-8') as f: json.dump(config_data, f, indent=4, ensure_ascii=False)
            QMessageBox.information(self, "成功", f"配置已保存到:\n{filepath}")
        except Exception as e: QMessageBox.critical(self, "保存失败", f"无法保存配置文件:\n{e}")

    @Slot()
    def load_config(self):
        if self.devices and QMessageBox.question(self, "确认", "加载配置将覆盖当前设备列表，确定吗？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No) == QMessageBox.StandardButton.No: return
        filepath, _ = QFileDialog.getOpenFileName(self, "加载设备配置", "", "JSON 文件 (*.json);;所有文件 (*)")
        if not filepath: return
        try:
            with open(filepath, 'r', encoding='utf-8') as f: config_data = json.load(f)
            self.devices = []; self.device_tablewidget.setRowCount(0); self.device_id_counter = 0; self.clear_results(); max_id = 0
            for data in config_data:
                new_device = Device.from_dict(data); self.devices.append(new_device); self._add_device_to_table(new_device)
                if new_device.id > max_id: max_id = new_device.id
            self.device_id_counter = max_id
            self._update_device_combos()
            QMessageBox.information(self, "成功", f"配置已从以下文件加载:\n{filepath}")
            self.node_positions = None; self.selected_node_id = None
        except FileNotFoundError: QMessageBox.critical(self, "加载失败", "找不到文件。")
        except json.JSONDecodeError: QMessageBox.critical(self, "加载失败", "文件格式错误。")
        except Exception as e: QMessageBox.critical(self, "加载失败", f"发生错误:\n{e}")

    @Slot()
    def export_connections(self):
        if not self.connections_result: QMessageBox.warning(self, "提示", "没有连接结果可导出。"); return
        filepath, selected_filter = QFileDialog.getSaveFileName(self, "导出连接列表", "", "文本文件 (*.txt);;CSV 文件 (*.csv);;所有文件 (*)")
        if not filepath: return
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                is_csv = "csv" in selected_filter.lower()
                if is_csv: f.write("序号,设备1,端口1,设备2,端口2,连接类型\n")
                for i, conn in enumerate(self.connections_result):
                    dev1, port1, dev2, port2, conn_type = conn
                    if is_csv: f.write(f"{i+1},{dev1.name},{port1},{dev2.name},{port2},{conn_type}\n")
                    else: f.write(f"{i+1}. {dev1.name} [{port1}] <-> {dev2.name} [{port2}] ({conn_type})\n")
            QMessageBox.information(self, "成功", f"连接列表已导出到:\n{filepath}")
        except Exception as e: QMessageBox.critical(self, "导出失败", f"无法导出连接列表:\n{e}")

    @Slot()
    def export_topology(self):
        if not self.fig: QMessageBox.warning(self, "提示", "没有拓扑图可导出。"); return
        filepath, selected_filter = QFileDialog.getSaveFileName(self, "导出拓扑图", "", "PNG 图像 (*.png);;PDF 文件 (*.pdf);;SVG 文件 (*.svg);;所有文件 (*)")
        if not filepath: return
        try:
            self.fig.savefig(filepath, dpi=300, bbox_inches='tight')
            QMessageBox.information(self, "成功", f"拓扑图已导出到:\n{filepath}")
        except Exception as e: QMessageBox.critical(self, "导出失败", f"无法导出拓扑图:\n{e}")

    @Slot(str)
    def filter_device_table(self, text):
        filter_text = text.lower()
        for row in range(self.device_tablewidget.rowCount()):
            match = False
            name_item = self.device_tablewidget.item(row, 0)
            type_item = self.device_tablewidget.item(row, 1)
            if name_item and filter_text in name_item.text().lower(): match = True
            if not match and type_item and filter_text in type_item.text().lower(): match = True
            self.device_tablewidget.setRowHidden(row, not match)

    @Slot()
    def filter_connection_list(self):
        selected_type = self.conn_filter_type_combo.currentText()
        filter_device_text = self.conn_filter_device_entry.text().strip().lower()
        type_filter_active = selected_type != "所有类型"
        for i in range(self.manual_connection_list.count()):
            item = self.manual_connection_list.item(i)
            conn_data = item.data(Qt.ItemDataRole.UserRole)
            if not conn_data or not isinstance(conn_data, tuple) or len(conn_data) != 5:
                item.setHidden(False); continue
            dev1, _, dev2, _, conn_type = conn_data
            item_conn_type = conn_type
            dev1_name_lower = dev1.name.lower(); dev2_name_lower = dev2.name.lower()
            type_match = True
            if type_filter_active: type_match = (item_conn_type == selected_type)
            device_match = True
            if filter_device_text: device_match = (filter_device_text in dev1_name_lower or filter_device_text in dev2_name_lower)
            item.setHidden(not (type_match and device_match))

    @Slot(object)
    def on_canvas_click(self, event):
        if event.inaxes != self.mpl_canvas.axes or not self.node_positions:
            if self.selected_node_id is not None:
                self.selected_node_id = None; print("清除选中 (点击背景)")
                self._update_connection_views()
            return
        clicked_node_id = None; min_dist_sq = float('inf')
        x, y = event.xdata, event.ydata
        if x is None or y is None:
             if self.selected_node_id is not None:
                 self.selected_node_id = None; print("清除选中 (点击轴外)")
                 self._update_connection_views()
             return
        xlim = self.mpl_canvas.axes.get_xlim(); ylim = self.mpl_canvas.axes.get_ylim()
        diagonal_len_sq = (xlim[1]-xlim[0])**2 + (ylim[1]-ylim[0])**2
        threshold_dist_sq = diagonal_len_sq * (0.03**2)
        for node_id, (nx, ny) in self.node_positions.items():
            dist_sq = (x - nx)**2 + (y - ny)**2
            if dist_sq < min_dist_sq and dist_sq < threshold_dist_sq:
                 min_dist_sq = dist_sq; clicked_node_id = node_id
        if event.dblclick:
            if clicked_node_id is not None:
                print(f"双击节点: {clicked_node_id}")
                device = next((d for d in self.devices if d.id == clicked_node_id), None)
                if device: self._display_device_details_popup(device)
        else:
            if clicked_node_id is not None:
                if self.selected_node_id == clicked_node_id:
                    self.selected_node_id = None; print(f"取消选中节点: {clicked_node_id}")
                else:
                    self.selected_node_id = clicked_node_id; print(f"选中节点: {self.selected_node_id}")
                self._update_connection_views()
            else:
                if self.selected_node_id is not None:
                    self.selected_node_id = None; print("清除选中 (点击背景)")
                    self._update_connection_views()

# --- 程序入口 (与 V35 相同) ---
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(APP_STYLE)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
    