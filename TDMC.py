# -*- coding: utf-8 -*-
# MediorNet TDM 连接计算器 V48
# 主要变更:
# - 代码重构:
#   - 使用常量替代表格列索引和类型字符串 (可读性/维护性)。
#   - 提取辅助函数简化画布事件处理逻辑 (可读性/维护性)。
# - 保留 V47 的功能。
# - 提供完整代码，无省略。

import sys
import tkinter as tk
from tkinter import messagebox
import networkx as nx
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.lines import Line2D
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import itertools
import random
import platform
import copy
import json
from collections import defaultdict
import io
import os # 需要 os 用于 resource_path
import base64
import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QTextEdit,
    QTabWidget, QFrame, QFileDialog, QMessageBox, QSpacerItem, QSizePolicy,
    QGridLayout, QListWidgetItem, QAbstractItemView, QTableWidget, QTableWidgetItem,
    QHeaderView, QListWidget, QSplitter, QCheckBox
)
from PySide6.QtCore import Slot, Qt
from PySide6.QtGui import QFont, QGuiApplication

# --- V48: Constants ---
# Table Columns
COL_NAME = 0
COL_TYPE = 1
COL_MPO = 2
COL_LC = 3
COL_SFP = 4
COL_CONN = 5

# Device Types
DEV_UHD = 'MicroN UHD'
DEV_HORIZON = 'HorizoN'
DEV_MN = 'MicroN'
UHD_TYPES = [DEV_UHD, DEV_HORIZON] # Helper list

# Port Types
PORT_MPO = 'MPO'
PORT_LC = 'LC'
PORT_SFP = 'SFP'
PORT_UNKNOWN = 'Unknown'
# --- End Constants ---

# --- 辅助函数 resource_path (用于打包后查找资源) ---
def resource_path(relative_path):
    """ 获取资源的绝对路径，适用于开发环境和 PyInstaller 打包后 """
    try:
        # PyInstaller 创建一个临时文件夹并将路径存储在 _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # 未打包状态下，获取当前脚本所在目录
        try:
             base_path = os.path.abspath(os.path.dirname(__file__))
        except NameError: # Fallback if __file__ is not defined (e.g., interactive)
             base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
# --- 结束辅助函数 ---


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

# --- 用于数字排序的 QTableWidgetItem 子类 ---
class NumericTableWidgetItem(QTableWidgetItem):
    """自定义 QTableWidgetItem 以支持数字排序。"""
    def __lt__(self, other):
        # 尝试将文本转换为浮点数进行比较
        # 使用 UserRole+1 存储原始数值以提高排序精度和鲁棒性
        data_self = self.data(Qt.ItemDataRole.UserRole + 1)
        data_other = other.data(Qt.ItemDataRole.UserRole + 1)

        try:
            # 优先使用存储的数值数据
            num_self = float(data_self)
            num_other = float(data_other)
            return num_self < num_other
        except (TypeError, ValueError):
            # 如果存储的数据无效或不存在，则回退到比较文本
            try:
                num_self = float(self.text())
                num_other = float(other.text())
                return num_self < num_other
            except ValueError:
                 # 如果文本也无法转换为数字，按字符串比较
                 return super().__lt__(other)

# --- 数据结构 ---
class Device:
    """代表一个 MediorNet 设备"""
    def __init__(self, id, name, type, mpo_ports=0, lc_ports=0, sfp_ports=0):
        self.id = id
        self.name = name
        self.type = type
        # 确保端口数是整数
        self.mpo_total = int(mpo_ports)
        self.lc_total = int(lc_ports)
        self.sfp_total = int(sfp_ports)
        self.reset_ports()

    def reset_ports(self):
        """重置端口连接状态和计数"""
        self.connections = 0.0 # 使用浮点数以精确表示 MPO 连接 (0.25)
        self.port_connections = {} # 字典：{ "端口名": "连接对端设备名" }

    def get_all_possible_ports(self):
        """获取设备所有可能的端口名称列表"""
        ports = []
        ports.extend([f"{PORT_LC}{i+1}" for i in range(self.lc_total)])
        ports.extend([f"{PORT_SFP}{i+1}" for i in range(self.sfp_total)])
        for i in range(self.mpo_total):
            base = f"{PORT_MPO}{i+1}"
            # MPO 端口有 4 个子通道 (Breakout)
            ports.extend([f"{base}-Ch{j+1}" for j in range(4)])
        return ports

    def get_all_available_ports(self):
        """获取设备当前所有可用的端口名称列表"""
        all_ports = self.get_all_possible_ports()
        used_ports = set(self.port_connections.keys())
        available = [p for p in all_ports if p not in used_ports]
        return available

    def use_specific_port(self, port_name, target_device_name):
        """标记指定端口为已使用，并连接到目标设备"""
        if port_name in self.get_all_possible_ports() and port_name not in self.port_connections:
            self.port_connections[port_name] = target_device_name
            # 更新连接计数，MPO Breakout 算 0.25 个连接，其他算 1 个
            if port_name.startswith(PORT_MPO):
                self.connections += 0.25
            else:
                self.connections += 1
            return True
        return False

    def return_port(self, port_name):
        """释放指定端口，使其变为可用状态"""
        port_in_use_record = port_name in self.port_connections
        port_already_available = False
        port_type_valid = True

        # 检查端口名格式是否有效
        port_type = get_port_type_from_name(port_name)
        if port_type != PORT_UNKNOWN:
             port_already_available = port_name in self.get_all_available_ports()
        else:
            print(f"警告: 尝试归还未知类型的端口 {port_name}")
            port_type_valid = False

        if not port_type_valid:
            return

        if port_in_use_record:
            target = self.port_connections.pop(port_name) # 从已用端口中移除
            # 只有当端口确实被移除（之前不在可用列表）时才减少连接计数
            if not port_already_available:
                if port_name.startswith(PORT_MPO):
                    self.connections -= 0.25
                elif port_name.startswith(PORT_LC) or port_name.startswith(PORT_SFP):
                    self.connections -= 1
                self.connections = max(0.0, self.connections) # 确保不为负
        else:
             # 如果端口本来就没被使用，则无需操作
             pass

    def get_available_port(self, port_type, target_device_name):
        """获取指定类型的第一个可用端口并标记为已用（旧逻辑，可能不再需要）"""
        possible_ports = []
        if port_type == PORT_LC: possible_ports = [f"{PORT_LC}{i+1}" for i in range(self.lc_total)]
        elif port_type == PORT_MPO:
            for i in range(self.mpo_total): base = f"{PORT_MPO}{i+1}"; possible_ports.extend([f"{base}-Ch{j+1}" for j in range(4)])
            random.shuffle(possible_ports) # 旧逻辑中包含随机性
        elif port_type == PORT_SFP: possible_ports = [f"{PORT_SFP}{i+1}" for i in range(self.sfp_total)]

        used_ports = set(self.port_connections.keys())
        for port in possible_ports:
            if port not in used_ports:
                self.port_connections[port] = target_device_name
                if port.startswith(PORT_MPO): self.connections += 0.25
                else: self.connections += 1
                return port
        return None

    def get_specific_available_port(self, port_type_prefix):
        """按顺序查找指定前缀的第一个可用端口（不标记为已用）"""
        possible_ports = []
        if port_type_prefix == PORT_LC: possible_ports = sorted([f"{PORT_LC}{i+1}" for i in range(self.lc_total)], key=lambda x: int(x[len(PORT_LC):]))
        elif port_type_prefix == PORT_MPO:
            for i in range(self.mpo_total): base = f"{PORT_MPO}{i+1}"; possible_ports.extend(sorted([f"{base}-Ch{j+1}" for j in range(4)], key=lambda x: int(x.split('-Ch')[-1])))
        elif port_type_prefix == PORT_SFP: possible_ports = sorted([f"{PORT_SFP}{i+1}" for i in range(self.sfp_total)], key=lambda x: int(x[len(PORT_SFP):]))

        used_ports = set(self.port_connections.keys())
        for port in possible_ports:
            if port not in used_ports:
                return port
        return None

    def to_dict(self):
        """将设备对象转换为字典，用于保存配置"""
        return {'id': self.id, 'name': self.name, 'type': self.type,
                'mpo_ports': self.mpo_total, 'lc_ports': self.lc_total, 'sfp_ports': self.sfp_total}

    @classmethod
    def from_dict(cls, data):
        """从字典创建设备对象，用于加载配置"""
        # 提供默认 ID 以兼容旧格式
        data.setdefault('id', random.randint(10000, 99999))
        return cls(id=data['id'], name=data['name'], type=data['type'],
                   mpo_ports=data.get('mpo_ports', 0), lc_ports=data.get('lc_ports', 0), sfp_ports=data.get('sfp_ports', 0))

    def __repr__(self):
        """返回对象的字符串表示形式"""
        return f"{self.name} ({self.type})"

# --- 连接计算逻辑 ---

def get_port_type_from_name(port_name):
    """Helper to get port type string from name using constants"""
    if not isinstance(port_name, str): return PORT_UNKNOWN
    if port_name.startswith(PORT_LC): return PORT_LC
    if port_name.startswith(PORT_SFP): return PORT_SFP
    if port_name.startswith(PORT_MPO): return PORT_MPO
    return PORT_UNKNOWN

def check_port_compatibility(dev1_type, port1_name, dev2_type, port2_name):
    """检查两个端口之间是否兼容 (使用常量)"""
    port1_type = get_port_type_from_name(port1_name)
    port2_type = get_port_type_from_name(port2_name)
    is_uhd1 = dev1_type in UHD_TYPES; is_uhd2 = dev2_type in UHD_TYPES
    is_mn1 = dev1_type == DEV_MN; is_mn2 = dev2_type == DEV_MN

    if port1_type == PORT_LC and port2_type == PORT_LC and is_uhd1 and is_uhd2: return True, "LC-LC (100G)"
    if port1_type == PORT_MPO and port2_type == PORT_MPO and is_uhd1 and is_uhd2: return True, "MPO-MPO (25G)"
    if port1_type == PORT_SFP and port2_type == PORT_SFP and is_mn1 and is_mn2: return True, "SFP-SFP (10G)"
    if port1_type == PORT_MPO and port2_type == PORT_SFP and is_uhd1 and is_mn2: return True, "MPO-SFP (10G)"
    if port1_type == PORT_SFP and port2_type == PORT_MPO and is_mn1 and is_uhd2: return True, "MPO-SFP (10G)"

    return False, None

def _get_compatible_port_types(other_dev_type, other_port_name):
    """根据另一侧设备类型和端口，获取本侧设备兼容的端口类型列表 (使用常量)"""
    other_port_type = get_port_type_from_name(other_port_name)
    compatible_here = []
    is_other_uhd = other_dev_type in UHD_TYPES
    is_other_mn = other_dev_type == DEV_MN

    if is_other_uhd:
        if other_port_type == PORT_LC: compatible_here = [PORT_LC]
        elif other_port_type == PORT_MPO: compatible_here = [PORT_MPO, PORT_SFP]
    elif is_other_mn:
        if other_port_type == PORT_SFP: compatible_here = [PORT_MPO, PORT_SFP]

    return compatible_here

def _find_best_single_link(dev1_copy, dev2_copy):
    """辅助函数：查找两个设备副本之间最高优先级的单个可用连接 (使用常量)"""
    is_uhd1 = dev1_copy.type in UHD_TYPES; is_uhd2 = dev2_copy.type in UHD_TYPES
    is_mn1 = dev1_copy.type == DEV_MN; is_mn2 = dev2_copy.type == DEV_MN
    # Try LC-LC
    if is_uhd1 and is_uhd2:
        port1 = dev1_copy.get_specific_available_port(PORT_LC)
        if port1:
            port2 = dev2_copy.get_specific_available_port(PORT_LC)
            if port2: dev1_copy.use_specific_port(port1, dev2_copy.name); dev2_copy.use_specific_port(port2, dev1_copy.name); return port1, port2, 'LC-LC (100G)'
    # Try MPO-MPO
    if is_uhd1 and is_uhd2:
        port1 = dev1_copy.get_specific_available_port(PORT_MPO)
        if port1:
            port2 = dev2_copy.get_specific_available_port(PORT_MPO)
            if port2: dev1_copy.use_specific_port(port1, dev2_copy.name); dev2_copy.use_specific_port(port2, dev1_copy.name); return port1, port2, 'MPO-MPO (25G)'
    # Try SFP-SFP
    if is_mn1 and is_mn2:
        port1 = dev1_copy.get_specific_available_port(PORT_SFP)
        if port1:
            port2 = dev2_copy.get_specific_available_port(PORT_SFP)
            if port2: dev1_copy.use_specific_port(port1, dev2_copy.name); dev2_copy.use_specific_port(port2, dev1_copy.name); return port1, port2, 'SFP-SFP (10G)'
    # Try MPO-SFP
    uhd_dev, micron_dev = (None, None); original_dev1 = dev1_copy
    if is_uhd1 and is_mn2: uhd_dev, micron_dev = dev1_copy, dev2_copy
    elif is_uhd2 and is_mn1: uhd_dev, micron_dev = dev2_copy, dev1_copy
    if uhd_dev and micron_dev:
        port_uhd = uhd_dev.get_specific_available_port(PORT_MPO)
        if port_uhd:
            port_micron = micron_dev.get_specific_available_port(PORT_SFP)
            if port_micron:
                uhd_dev.use_specific_port(port_uhd, micron_dev.name); micron_dev.use_specific_port(port_micron, uhd_dev.name)
                if original_dev1 == uhd_dev: return port_uhd, port_micron, 'MPO-SFP (10G)'
                else: return port_micron, port_uhd, 'MPO-SFP (10G)'
    return None, None, None

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
    failed_pairs_phase1 = []
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
                else:
                    if pair_key not in [fp[0] for fp in failed_pairs_phase1]: failed_pairs_phase1.append((pair_key, f"{original_dev1.name} <-> {original_dev2.name}"))
        if not made_progress_phase1: break
    if len(connected_once_pairs) < len(all_pairs_ids): print(f"警告: Phase 1 未能为所有设备对建立连接。失败 {len(all_pairs_ids) - len(connected_once_pairs)} 对。")
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

def calculate_ring_connections(devices):
    """计算环形连接，返回更详细的错误信息"""
    if len(devices) < 2: return [], {}, "设备数量少于2，无法形成环形"
    if len(devices) == 2:
        conns, state = calculate_mesh_connections(devices)
        return conns, state, None
    connections = []; temp_devices = [copy.deepcopy(dev) for dev in devices]
    for d in temp_devices: d.reset_ports()
    device_map = {dev.id: dev for dev in temp_devices}
    sorted_dev_ids = sorted([d.id for d in devices]); num_devices = len(sorted_dev_ids)
    link_established = [True] * num_devices
    failed_segments = []
    for i in range(num_devices):
        dev1_id = sorted_dev_ids[i]; dev2_id = sorted_dev_ids[(i + 1) % num_devices]
        dev1_copy = device_map[dev1_id]; dev2_copy = device_map[dev2_id]
        original_dev1 = next(d for d in devices if d.id == dev1_id); original_dev2 = next(d for d in devices if d.id == dev2_id)
        port1, port2, conn_type = _find_best_single_link(dev1_copy, dev2_copy)
        if port1 and port2:
            if dev1_copy.id == dev1_id: connections.append((original_dev1, port1, original_dev2, port2, conn_type))
            else: connections.append((original_dev2, port2, original_dev1, port1, conn_type))
        else:
            link_established[i] = False
            failed_segments.append(f"{original_dev1.name} 和 {original_dev2.name}")
            print(f"警告: 无法在 {original_dev1.name} 和 {original_dev2.name} 之间建立环形连接段。")
    error_message = None
    if not all(link_established):
        error_message = f"未能完成完整的环形连接。无法连接的段落：{', '.join(failed_segments)}。"
        print(f"警告: {error_message}")
    return connections, device_map, error_message

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

    def plot_topology(self, devices, connections, layout_algorithm='spring', fixed_pos=None, selected_node_id=None, port_totals_dict=None):
        """绘制拓扑图，支持节点高亮和显示端口总数"""
        self.axes.cla()
        if not devices: self.axes.text(0.5, 0.5, '无设备数据', ha='center', va='center'); self.draw(); return None, None

        # --- Font setup (using resource_path) ---
        font_prop = None; current_font_family = 'sans-serif'
        try:
            font_relative_path = os.path.join('assets', 'NotoSansCJKsc-Regular.otf')
            font_path = resource_path(font_relative_path)
            if os.path.exists(font_path):
                font_manager.fontManager.addfont(font_path)
                plt.rcParams['font.sans-serif'] = ['Noto Sans CJK SC'] + plt.rcParams.get('font.sans-serif', [])
                font_prop = font_manager.FontProperties(family='Noto Sans CJK SC')
                current_font_family = 'Noto Sans CJK SC'
            else: # Fallback
                print(f"警告: 字体文件未找到: {font_path}")
                plt.rcParams['font.sans-serif'] = ['sans-serif']
                font_prop = font_manager.FontProperties()

            plt.rcParams['axes.unicode_minus'] = False
        except Exception as e:
            print(f"设置 Matplotlib 字体时出错: {e}")
            plt.rcParams['font.sans-serif'] = ['sans-serif']
            plt.rcParams['axes.unicode_minus'] = False
            font_prop = font_manager.FontProperties()
            current_font_family = 'sans-serif'
        # --- End Font setup ---

        G = nx.Graph(); node_ids = [dev.id for dev in devices]; node_colors = []; node_labels = {}; node_alphas = []
        highlight_color = 'yellow'; default_alpha = 0.9; dimmed_alpha = 0.3
        for dev in devices:
            G.add_node(dev.id); node_labels[dev.id] = f"{dev.name}\n({dev.type})"
            base_color = 'grey';
            if dev.type == DEV_UHD: base_color = 'skyblue'
            elif dev.type == DEV_HORIZON: base_color = 'lightcoral'
            elif dev.type == DEV_MN: base_color = 'lightgreen'
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

        # 在图形上显示端口总数
        if port_totals_dict is not None:
            totals_text = f"端口总计: {PORT_MPO}: {port_totals_dict['mpo']}, {PORT_LC}: {port_totals_dict['lc']}, {PORT_SFP}+: {port_totals_dict['sfp']}"
            self.fig.text(0.01, 0.01, totals_text, ha='left', va='bottom', fontsize=7, color='grey', transform=self.fig.transFigure)

        self.axes.set_title("网络连接拓扑图", fontproperties=font_prop); self.axes.axis('off')
        legend_elements = [
            plt.Line2D([0], [0], marker='o', color='w', label=DEV_UHD, markerfacecolor='skyblue', markersize=10),
            plt.Line2D([0], [0], marker='o', color='w', label=DEV_HORIZON, markerfacecolor='lightcoral', markersize=10),
            plt.Line2D([0], [0], marker='o', color='w', label=DEV_MN, markerfacecolor='lightgreen', markersize=10),
            plt.Line2D([0], [0], color='blue', lw=2, label='LC-LC (100G)'), plt.Line2D([0], [0], color='red', lw=2, label='MPO-MPO (25G)'),
            plt.Line2D([0], [0], color='orange', lw=2, label='MPO-SFP (10G)'), plt.Line2D([0], [0], color='purple', lw=2, label='SFP-SFP (10G)')
        ]
        legend_prop_small = copy.copy(font_prop); legend_prop_small.set_size('small')
        self.axes.legend(handles=legend_elements, loc='best', prop=legend_prop_small)
        self.fig.tight_layout(rect=[0, 0.03, 1, 1])
        self.draw_idle()
        return self.fig, pos
# --- 结束 Matplotlib Canvas ---

# --- 主窗口 (MainWindow Class - Refactored) ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MediorNet TDM 连接计算器 V1.1")
        self.setGeometry(100, 100, 1100, 800)
        # ... (成员变量初始化，与 V47 相同) ...
        self.devices = []
        self.connections_result = []
        self.fig = None
        self.node_positions = None
        self.device_id_counter = 0
        self.selected_node_id = None
        self.dragged_node_id = None
        self.drag_offset = (0, 0)
        self.connecting_node_id = None
        self.connection_line = None
        self.suppress_confirmations = False

        # 字体加载 (同 V47)
        try:
            font_relative_path = os.path.join('assets', 'NotoSansCJKsc-Regular.otf')
            font_path = resource_path(font_relative_path)
            if os.path.exists(font_path):
                font_manager.fontManager.addfont(font_path)
                plt.rcParams['font.sans-serif'] = ['Noto Sans CJK SC'] + plt.rcParams.get('font.sans-serif', [])
                print(f"成功加载并设置打包字体: Noto Sans CJK SC (路径: {font_path})")
                self.chinese_font = QFont("Noto Sans CJK SC", 10)
            else:
                print(f"警告: 未在路径 {font_path} 找到打包的字体文件。将使用默认字体。")
                plt.rcParams['font.sans-serif'] = ['sans-serif']
                self.chinese_font = QFont(); self.chinese_font.setPointSize(10) # Use default Qt font
        except Exception as e:
            print(f"加载或设置打包字体时出错: {e}")
            plt.rcParams['font.sans-serif'] = ['sans-serif']
            self.chinese_font = QFont(); self.chinese_font.setPointSize(10)
        plt.rcParams['axes.unicode_minus'] = False

        # --- UI 布局与控件初始化 ---
        main_widget = QWidget(); self.setCentralWidget(main_widget); main_layout = QHBoxLayout(main_widget)
        main_splitter = QSplitter(Qt.Orientation.Horizontal); main_layout.addWidget(main_splitter)
        left_panel = QFrame(); left_panel.setFrameShape(QFrame.Shape.StyledPanel);
        left_layout = QVBoxLayout(left_panel); main_splitter.addWidget(left_panel)

        # 添加设备组
        add_group = QFrame(); add_group.setObjectName("addDeviceGroup"); add_group_layout = QGridLayout(add_group); add_group_layout.setContentsMargins(10, 15, 10, 10); add_group_layout.setVerticalSpacing(8); add_title = QLabel("<b>添加新设备</b>"); add_title.setFont(QFont(self.chinese_font.family(), 11)); add_group_layout.addWidget(add_title, 0, 0, 1, 2, alignment=Qt.AlignmentFlag.AlignCenter); add_group_layout.addWidget(QLabel("类型:"), 1, 0); self.device_type_combo = QComboBox(); self.device_type_combo.addItems([DEV_UHD, DEV_HORIZON, DEV_MN]); self.device_type_combo.setFont(self.chinese_font); self.device_type_combo.currentIndexChanged.connect(self.update_port_entries); add_group_layout.addWidget(self.device_type_combo, 1, 1); add_group_layout.addWidget(QLabel("名称:"), 2, 0); self.device_name_entry = QLineEdit(); self.device_name_entry.setFont(self.chinese_font); add_group_layout.addWidget(self.device_name_entry, 2, 1); self.mpo_label = QLabel(f"{PORT_MPO} 端口:"); add_group_layout.addWidget(self.mpo_label, 3, 0); self.mpo_entry = QLineEdit("4"); self.mpo_entry.setFont(self.chinese_font); add_group_layout.addWidget(self.mpo_entry, 3, 1); self.lc_label = QLabel(f"{PORT_LC} 端口:"); add_group_layout.addWidget(self.lc_label, 4, 0); self.lc_entry = QLineEdit("2"); self.lc_entry.setFont(self.chinese_font); add_group_layout.addWidget(self.lc_entry, 4, 1); self.sfp_label = QLabel(f"{PORT_SFP}+ 端口:"); self.sfp_entry = QLineEdit("8"); self.sfp_entry.setFont(self.chinese_font); add_group_layout.addWidget(self.sfp_label, 5, 0); add_group_layout.addWidget(self.sfp_entry, 5, 1); self.sfp_label.hide(); self.sfp_entry.hide(); self.add_button = QPushButton("添加设备"); self.add_button.setFont(self.chinese_font); self.add_button.clicked.connect(self.add_device); add_group_layout.addWidget(self.add_button, 6, 0, 1, 2, alignment=Qt.AlignmentFlag.AlignCenter); left_layout.addWidget(add_group)

        # 设备列表组
        list_group = QFrame(); list_group.setObjectName("listGroup"); list_group_layout = QVBoxLayout(list_group); list_group_layout.setContentsMargins(10, 15, 10, 10); filter_layout = QHBoxLayout(); filter_layout.addWidget(QLabel("过滤:", font=self.chinese_font)); self.device_filter_entry = QLineEdit(); self.device_filter_entry.setFont(self.chinese_font); self.device_filter_entry.setPlaceholderText("按名称或类型过滤..."); self.device_filter_entry.textChanged.connect(self.filter_device_table); filter_layout.addWidget(self.device_filter_entry); list_group_layout.addLayout(filter_layout); self.device_tablewidget = QTableWidget(); self.device_tablewidget.setFont(self.chinese_font); self.device_tablewidget.setColumnCount(6); self.device_tablewidget.setHorizontalHeaderLabels(["名称", "类型", PORT_MPO, PORT_LC, f"{PORT_SFP}+", "连接数(估)"]); self.device_tablewidget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); self.device_tablewidget.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection); self.device_tablewidget.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.SelectedClicked | QAbstractItemView.EditTrigger.EditKeyPressed); self.device_tablewidget.setSortingEnabled(True); header = self.device_tablewidget.horizontalHeader(); header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive); header.setSectionResizeMode(COL_NAME, QHeaderView.ResizeMode.Stretch); self.device_tablewidget.setColumnWidth(COL_TYPE, 90); self.device_tablewidget.setColumnWidth(COL_MPO, 50); self.device_tablewidget.setColumnWidth(COL_LC, 50); self.device_tablewidget.setColumnWidth(COL_SFP, 50); self.device_tablewidget.setColumnWidth(COL_CONN, 80); self.device_tablewidget.itemDoubleClicked.connect(self.show_device_details_from_table); self.device_tablewidget.itemChanged.connect(self.on_device_item_changed); list_group_layout.addWidget(self.device_tablewidget); device_op_layout = QHBoxLayout(); self.remove_button = QPushButton("移除选中"); self.remove_button.setFont(self.chinese_font); self.remove_button.clicked.connect(self.remove_device); self.clear_button = QPushButton("清空所有"); self.clear_button.setFont(self.chinese_font); self.clear_button.clicked.connect(self.clear_all_devices); device_op_layout.addWidget(self.remove_button); device_op_layout.addWidget(self.clear_button); list_group_layout.addLayout(device_op_layout); self.port_totals_label = QLabel("总计: MPO: 0, LC: 0, SFP+: 0"); font = self.port_totals_label.font(); font.setBold(True); self.port_totals_label.setFont(font); self.port_totals_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter); self.port_totals_label.setStyleSheet("padding-top: 5px; padding-right: 5px;"); list_group_layout.addWidget(self.port_totals_label); left_layout.addWidget(list_group)

        # 文件操作组
        file_group = QFrame(); file_group.setObjectName("fileGroup"); file_group_layout = QGridLayout(file_group); file_group_layout.setContentsMargins(10, 15, 10, 10); file_title = QLabel("<b>文件操作</b>"); file_title.setFont(QFont(self.chinese_font.family(), 11)); file_group_layout.addWidget(file_title, 0, 0, 1, 2, alignment=Qt.AlignmentFlag.AlignCenter); self.save_button = QPushButton("保存配置"); self.save_button.setFont(self.chinese_font); self.save_button.clicked.connect(self.save_config); self.load_button = QPushButton("加载配置"); self.load_button.setFont(self.chinese_font); self.load_button.clicked.connect(self.load_config); file_group_layout.addWidget(self.save_button, 1, 0); file_group_layout.addWidget(self.load_button, 1, 1); self.export_list_button = QPushButton("导出列表"); self.export_list_button.setFont(self.chinese_font); self.export_list_button.clicked.connect(self.export_connections); self.export_list_button.setEnabled(False); file_group_layout.addWidget(self.export_list_button, 2, 0); self.export_topo_button = QPushButton("导出拓扑图"); self.export_topo_button.setFont(self.chinese_font); self.export_topo_button.clicked.connect(self.export_topology); self.export_topo_button.setEnabled(False); file_group_layout.addWidget(self.export_topo_button, 2, 1); self.export_report_button = QPushButton("导出报告 (HTML)"); self.export_report_button.setFont(self.chinese_font); self.export_report_button.clicked.connect(self.export_html_report); self.export_report_button.setEnabled(False); file_group_layout.addWidget(self.export_report_button, 3, 0, 1, 2); left_layout.addWidget(file_group)

        # 跳过确认弹窗设置
        suppress_frame = QFrame(); suppress_frame.setFrameShape(QFrame.Shape.NoFrame); suppress_layout = QHBoxLayout(suppress_frame); suppress_layout.setContentsMargins(10, 0, 10, 5); self.suppress_confirm_checkbox = QCheckBox("跳过确认弹窗"); self.suppress_confirm_checkbox.setFont(self.chinese_font); self.suppress_confirm_checkbox.stateChanged.connect(self._toggle_suppress_confirmations); suppress_layout.addWidget(self.suppress_confirm_checkbox); suppress_layout.addStretch(); left_layout.addWidget(suppress_frame)
        left_layout.addStretch()

        # 右侧面板与 Tab 页
        right_panel = QFrame(); right_panel.setFrameShape(QFrame.Shape.StyledPanel); right_layout = QVBoxLayout(right_panel); main_splitter.addWidget(right_panel); calculate_control_frame = QFrame(); calculate_control_frame.setObjectName("calculateControlFrame"); calculate_control_layout = QHBoxLayout(calculate_control_frame); calculate_control_layout.setContentsMargins(10, 5, 10, 5); calculate_control_layout.addWidget(QLabel("计算模式:", font=self.chinese_font)); self.topology_mode_combo = QComboBox(); self.topology_mode_combo.setFont(self.chinese_font); self.topology_mode_combo.addItems(["Mesh", "环形"]); calculate_control_layout.addWidget(self.topology_mode_combo); calculate_control_layout.addWidget(QLabel("布局:", font=self.chinese_font)); self.layout_combo = QComboBox(); self.layout_combo.setFont(self.chinese_font); self.layout_combo.addItems(["Spring", "Circular", "Kamada-Kawai", "Random", "Shell"]); self.layout_combo.currentIndexChanged.connect(self.on_layout_change); calculate_control_layout.addWidget(self.layout_combo); self.calculate_button = QPushButton("计算连接"); self.calculate_button.setFont(self.chinese_font); self.calculate_button.clicked.connect(self.calculate_and_display); calculate_control_layout.addWidget(self.calculate_button); self.fill_mesh_button = QPushButton("填充 (Mesh)"); self.fill_mesh_button.setFont(self.chinese_font); self.fill_mesh_button.setEnabled(False); self.fill_mesh_button.clicked.connect(self.fill_remaining_mesh); calculate_control_layout.addWidget(self.fill_mesh_button); self.fill_ring_button = QPushButton("填充 (环形)"); self.fill_ring_button.setFont(self.chinese_font); self.fill_ring_button.setEnabled(False); self.fill_ring_button.clicked.connect(self.fill_remaining_ring); calculate_control_layout.addWidget(self.fill_ring_button); calculate_control_layout.addStretch(); right_layout.addWidget(calculate_control_frame); self.tab_widget = QTabWidget(); self.tab_widget.setFont(self.chinese_font); right_layout.addWidget(self.tab_widget); self.connections_tab = QWidget(); connections_layout = QVBoxLayout(self.connections_tab); self.connections_textedit = QTextEdit(); self.connections_textedit.setFont(self.chinese_font); self.connections_textedit.setReadOnly(True); connections_layout.addWidget(self.connections_textedit); self.tab_widget.addTab(self.connections_tab, "连接列表"); self.topology_tab = QWidget(); topology_layout = QVBoxLayout(self.topology_tab); self.mpl_canvas = MplCanvas(self.topology_tab, width=8, height=6, dpi=100); topology_layout.addWidget(self.mpl_canvas); self.tab_widget.addTab(self.topology_tab, "拓扑图"); self.mpl_canvas.mpl_connect('button_press_event', self.on_canvas_press); self.mpl_canvas.mpl_connect('motion_notify_event', self.on_canvas_motion); self.mpl_canvas.mpl_connect('button_release_event', self.on_canvas_release); self.edit_tab = QWidget(); edit_main_layout = QVBoxLayout(self.edit_tab); add_manual_group = QFrame(); add_manual_group.setObjectName("addManualGroup"); add_manual_group.setFrameShape(QFrame.Shape.StyledPanel); add_manual_layout = QGridLayout(add_manual_group); add_manual_layout.setContentsMargins(10,10,10,10); add_manual_layout.addWidget(QLabel("<b>添加手动连接</b>", font=self.chinese_font), 0, 0, 1, 4, alignment=Qt.AlignmentFlag.AlignCenter); add_manual_layout.addWidget(QLabel("设备 1:", font=self.chinese_font), 1, 0); self.edit_dev1_combo = QComboBox(); self.edit_dev1_combo.setFont(self.chinese_font); add_manual_layout.addWidget(self.edit_dev1_combo, 1, 1); add_manual_layout.addWidget(QLabel("端口 1:", font=self.chinese_font), 1, 2); self.edit_port1_combo = QComboBox(); self.edit_port1_combo.setFont(self.chinese_font); add_manual_layout.addWidget(self.edit_port1_combo, 1, 3); add_manual_layout.addWidget(QLabel("设备 2:", font=self.chinese_font), 2, 0); self.edit_dev2_combo = QComboBox(); self.edit_dev2_combo.setFont(self.chinese_font); add_manual_layout.addWidget(self.edit_dev2_combo, 2, 1); add_manual_layout.addWidget(QLabel("端口 2:", font=self.chinese_font), 2, 2); self.edit_port2_combo = QComboBox(); self.edit_port2_combo.setFont(self.chinese_font); add_manual_layout.addWidget(self.edit_port2_combo, 2, 3); self.add_manual_button = QPushButton("添加连接"); self.add_manual_button.setFont(self.chinese_font); self.add_manual_button.clicked.connect(self.add_manual_connection); add_manual_layout.addWidget(self.add_manual_button, 3, 0, 1, 4, alignment=Qt.AlignmentFlag.AlignCenter); self.edit_dev1_combo.currentIndexChanged.connect(self._update_manual_port_options); self.edit_port1_combo.currentIndexChanged.connect(self._update_manual_port_options); self.edit_dev2_combo.currentIndexChanged.connect(self._update_manual_port_options); self.edit_port2_combo.currentIndexChanged.connect(self._update_manual_port_options); edit_main_layout.addWidget(add_manual_group); remove_manual_group = QFrame(); remove_manual_group.setObjectName("removeManualGroup"); remove_manual_group.setFrameShape(QFrame.Shape.StyledPanel); remove_manual_layout = QVBoxLayout(remove_manual_group); remove_manual_layout.setContentsMargins(10,10,10,10); remove_title = QLabel("<b>移除现有连接</b> (选中下方列表中的连接进行移除)"); remove_title.setFont(self.chinese_font); remove_manual_layout.addWidget(remove_title); filter_conn_layout = QHBoxLayout(); filter_conn_layout.addWidget(QLabel("类型过滤:", font=self.chinese_font)); self.conn_filter_type_combo = QComboBox(); self.conn_filter_type_combo.setFont(self.chinese_font); self.conn_filter_type_combo.addItems(["所有类型", "LC-LC (100G)", "MPO-MPO (25G)", "SFP-SFP (10G)", "MPO-SFP (10G)"]); self.conn_filter_type_combo.currentIndexChanged.connect(self.filter_connection_list); filter_conn_layout.addWidget(self.conn_filter_type_combo); filter_conn_layout.addSpacing(15); filter_conn_layout.addWidget(QLabel("设备过滤:", font=self.chinese_font)); self.conn_filter_device_entry = QLineEdit(); self.conn_filter_device_entry.setFont(self.chinese_font); self.conn_filter_device_entry.setPlaceholderText("按设备名称过滤..."); self.conn_filter_device_entry.textChanged.connect(self.filter_connection_list); filter_conn_layout.addWidget(self.conn_filter_device_entry); remove_manual_layout.insertLayout(1, filter_conn_layout); self.manual_connection_list = QListWidget(); self.manual_connection_list.setFont(self.chinese_font); self.manual_connection_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection); remove_manual_layout.addWidget(self.manual_connection_list); self.remove_manual_button = QPushButton("移除选中连接"); self.remove_manual_button.setFont(self.chinese_font); self.remove_manual_button.clicked.connect(self.remove_manual_connection); self.remove_manual_button.setEnabled(False); remove_manual_layout.addWidget(self.remove_manual_button, alignment=Qt.AlignmentFlag.AlignCenter); edit_main_layout.addWidget(remove_manual_group); self.tab_widget.addTab(self.edit_tab, "手动编辑")

        main_splitter.setSizes([400, 700])
        main_splitter.setStretchFactor(1, 1)
        self.update_port_entries()
        self._update_port_totals_display()
        self._update_connection_views()

    # --- 方法 ---
    @Slot()
    def update_port_entries(self):
        selected_type = self.device_type_combo.currentText()
        is_micron = selected_type == DEV_MN
        is_uhd_horizon = selected_type in UHD_TYPES
        self.mpo_label.setVisible(is_uhd_horizon); self.mpo_entry.setVisible(is_uhd_horizon)
        self.lc_label.setVisible(is_uhd_horizon); self.lc_entry.setVisible(is_uhd_horizon)
        self.sfp_label.setVisible(is_micron); self.sfp_entry.setVisible(is_micron)

    def _add_device_to_table(self, device):
        """将设备对象添加到表格中，并设置可编辑标志和使用常量"""
        row_position = self.device_tablewidget.rowCount()
        self.device_tablewidget.insertRow(row_position)
        editable_flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable
        non_editable_flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled
        name_item = QTableWidgetItem(device.name); name_item.setData(Qt.ItemDataRole.UserRole, device.id); name_item.setFlags(editable_flags)
        type_item = QTableWidgetItem(device.type); type_item.setFlags(non_editable_flags)
        mpo_item = NumericTableWidgetItem(str(device.mpo_total)); mpo_item.setData(Qt.ItemDataRole.UserRole + 1, device.mpo_total); mpo_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter); mpo_item.setFlags(editable_flags if device.type in UHD_TYPES else non_editable_flags)
        lc_item = NumericTableWidgetItem(str(device.lc_total)); lc_item.setData(Qt.ItemDataRole.UserRole + 1, device.lc_total); lc_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter); lc_item.setFlags(editable_flags if device.type in UHD_TYPES else non_editable_flags)
        sfp_item = NumericTableWidgetItem(str(device.sfp_total)); sfp_item.setData(Qt.ItemDataRole.UserRole + 1, device.sfp_total); sfp_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter); sfp_item.setFlags(editable_flags if device.type == DEV_MN else non_editable_flags)
        conn_val = float(f"{device.connections:.2f}"); conn_item = NumericTableWidgetItem(f"{conn_val:.2f}"); conn_item.setData(Qt.ItemDataRole.UserRole + 1, conn_val); conn_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter); conn_item.setFlags(non_editable_flags)
        self.device_tablewidget.setItem(row_position, COL_NAME, name_item)
        self.device_tablewidget.setItem(row_position, COL_TYPE, type_item)
        self.device_tablewidget.setItem(row_position, COL_MPO, mpo_item)
        self.device_tablewidget.setItem(row_position, COL_LC, lc_item)
        self.device_tablewidget.setItem(row_position, COL_SFP, sfp_item)
        self.device_tablewidget.setItem(row_position, COL_CONN, conn_item)

    @Slot(int)
    def _toggle_suppress_confirmations(self, state):
        """更新是否跳过确认弹窗的状态"""
        self.suppress_confirmations = (state == Qt.CheckState.Checked.value)
        print(f"跳过确认弹窗: {'已启用' if self.suppress_confirmations else '已禁用'}")

    @Slot(QTableWidgetItem)
    def on_device_item_changed(self, item):
        """处理设备表格中项目被编辑后的事件 (使用常量)"""
        if not item: return
        row = item.row(); col = item.column()
        name_item = self.device_tablewidget.item(row, COL_NAME) # Use constant
        if not name_item: return
        dev_id = name_item.data(Qt.ItemDataRole.UserRole)
        device = next((d for d in self.devices if d.id == dev_id), None)
        if not device: return

        new_value = item.text().strip()
        self.device_tablewidget.blockSignals(True)

        try:
            if col == COL_NAME: # 名称列
                old_name = device.name
                if not new_value:
                    QMessageBox.warning(self, "输入错误", "设备名称不能为空。")
                    item.setText(old_name)
                elif new_value != old_name and any(d.name == new_value for d in self.devices if d.id != dev_id):
                    QMessageBox.warning(self, "输入错误", f"设备名称 '{new_value}' 已存在。")
                    item.setText(old_name)
                elif new_value != old_name:
                    print(f"设备 '{old_name}' 重命名为 '{new_value}'")
                    device.name = new_value
                    # 更新其他设备 port_connections 中的引用
                    for other_dev in self.devices:
                        if other_dev.id != device.id:
                            ports_to_update = {p: new_value for p, target in other_dev.port_connections.items() if target == old_name}
                            other_dev.port_connections.update(ports_to_update)
                    self._update_device_combos()
                    self._update_connection_views() # Trigger redraw for graph labels and potentially list
            elif col in [COL_MPO, COL_LC, COL_SFP]: # 端口数列
                port_attr_map = {COL_MPO: 'mpo_total', COL_LC: 'lc_total', COL_SFP: 'sfp_total'}
                port_name_map = {COL_MPO: 'MPO', COL_LC: 'LC', COL_SFP: 'SFP+'}
                attr_name = port_attr_map.get(col)
                port_type_name = port_name_map.get(col)
                is_uhd_horizon = device.type in UHD_TYPES
                is_micron = device.type == DEV_MN
                can_edit_this_port = (attr_name in ['mpo_total', 'lc_total'] and is_uhd_horizon) or \
                                     (attr_name == 'sfp_total' and is_micron)

                if not can_edit_this_port:
                     old_count = getattr(device, attr_name, 0)
                     print(f"不允许修改设备类型 '{device.type}' 的 '{port_type_name}' 端口数量。")
                     item.setText(str(old_count)); self.device_tablewidget.blockSignals(False); return

                old_count = getattr(device, attr_name, 0)
                try:
                    new_count = int(new_value)
                    if new_count < 0: raise ValueError("端口数不能为负")
                    if new_count != old_count:
                        user_confirmed_port_change = True
                        if not self.suppress_confirmations:
                            reply = QMessageBox.question(self, "确认修改端口数量",
                                                         f"修改设备 '{device.name}' 的 {port_type_name} 端口数量将清除所有现有连接并可能需要重新计算布局。\n是否继续？",
                                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
                            user_confirmed_port_change = (reply == QMessageBox.StandardButton.Yes)
                        if user_confirmed_port_change:
                            print(f"设备 '{device.name}' 的 {attr_name} 从 {old_count} 修改为 {new_count}")
                            setattr(device, attr_name, new_count)
                            self.clear_results() # Will update views
                            self._update_port_totals_display()
                            self._update_manual_port_options()
                        else:
                            print("用户取消修改端口数量。"); item.setText(str(old_count))
                except ValueError:
                    QMessageBox.warning(self, "输入错误", f"{port_type_name} 端口数量必须是非负整数。")
                    item.setText(str(old_count))
        finally:
            self.device_tablewidget.blockSignals(False)

    def _update_device_combos(self):
        """更新手动编辑区域的设备下拉框 (并触发端口更新)"""
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
        self._update_manual_port_options()

    def _populate_edit_port_combos(self, device_combo_to_populate, port_combo_to_populate, other_device_combo, other_port_combo):
        """动态填充指定的端口下拉列表，并根据另一侧的选择进行过滤 (使用常量)。"""
        port_combo_to_populate.blockSignals(True)
        current_port_selection = port_combo_to_populate.currentText()
        port_combo_to_populate.clear(); port_combo_to_populate.addItem("选择端口...")
        dev_id = device_combo_to_populate.currentData(); port_combo_to_populate.setEnabled(False)
        if dev_id is not None:
            device = next((d for d in self.devices if d.id == dev_id), None)
            if device:
                available_ports = device.get_all_available_ports(); ports_to_add = []
                other_dev_id = other_device_combo.currentData()
                other_port_name = other_port_combo.currentText()
                other_device = next((d for d in self.devices if d.id == other_dev_id), None) if other_dev_id else None
                if other_device and other_port_name != "选择端口...":
                    compatible_types_here = _get_compatible_port_types(other_device.type, other_port_name)
                    for port in available_ports:
                        port_type_here = get_port_type_from_name(port) # Use helper
                        if port_type_here in compatible_types_here: ports_to_add.append(port)
                else: ports_to_add = available_ports
                if ports_to_add:
                    port_combo_to_populate.addItems(ports_to_add)
                    index_to_select = port_combo_to_populate.findText(current_port_selection)
                    port_combo_to_populate.setCurrentIndex(index_to_select if index_to_select != -1 else 0)
                    port_combo_to_populate.setEnabled(True)
                else: port_combo_to_populate.addItem("无兼容/可用端口"); port_combo_to_populate.setCurrentIndex(1)
        port_combo_to_populate.blockSignals(False)

    @Slot()
    def _update_manual_port_options(self):
        """统一更新手动添加连接中的两个端口下拉列表的选项"""
        self._populate_edit_port_combos(self.edit_dev1_combo, self.edit_port1_combo, self.edit_dev2_combo, self.edit_port2_combo)
        self._populate_edit_port_combos(self.edit_dev2_combo, self.edit_port2_combo, self.edit_dev1_combo, self.edit_port1_combo)

    def _calculate_port_totals(self):
        """计算当前设备列表中各类端口的总数"""
        total_mpo = sum(dev.mpo_total for dev in self.devices); total_lc = sum(dev.lc_total for dev in self.devices); total_sfp = sum(dev.sfp_total for dev in self.devices)
        return {'mpo': total_mpo, 'lc': total_lc, 'sfp': total_sfp}

    def _update_port_totals_display(self):
        """更新显示端口总数的标签"""
        totals = self._calculate_port_totals()
        self.port_totals_label.setText(f"总计: {PORT_MPO}: {totals['mpo']}, {PORT_LC}: {totals['lc']}, {PORT_SFP}+: {totals['sfp']}")

    @Slot()
    def add_device(self):
        # Uses constants via update_port_entries call
        dtype = self.device_type_combo.currentText(); name = self.device_name_entry.text().strip()
        if not dtype: QMessageBox.critical(self, "错误", "请选择设备类型。"); return
        if not name: QMessageBox.critical(self, "错误", "请输入设备名称。"); return
        if any(dev.name == name for dev in self.devices): QMessageBox.critical(self, "错误", f"设备名称 '{name}' 已存在。"); return
        mpo_ports_str, lc_ports_str, sfp_ports_str = self.mpo_entry.text() or "0", self.lc_entry.text() or "0", self.sfp_entry.text() or "0"
        try:
            if dtype in UHD_TYPES: mpo_ports, lc_ports, sfp_ports = int(mpo_ports_str), int(lc_ports_str), 0; assert mpo_ports >= 0 and lc_ports >= 0
            elif dtype == DEV_MN: sfp_ports, mpo_ports, lc_ports = int(sfp_ports_str), 0, 0; assert sfp_ports >= 0
            else: raise ValueError("无效类型")
        except (ValueError, AssertionError): QMessageBox.critical(self, "输入错误", "端口数量必须是非负整数。"); return
        self.device_id_counter += 1
        new_device = Device(self.device_id_counter, name, dtype, mpo_ports, lc_ports, sfp_ports)
        self.devices.append(new_device); self._add_device_to_table(new_device); self.device_name_entry.clear();
        self._update_device_combos(); self.clear_results()
        self.node_positions = None; self.selected_node_id = None
        self._update_port_totals_display()

    @Slot()
    def remove_device(self):
        selected_rows = sorted(list(set(index.row() for index in self.device_tablewidget.selectedIndexes())), reverse=True)
        if not selected_rows: QMessageBox.warning(self, "提示", "请先在表格中选择要移除的设备行。"); return
        ids_to_remove = {self.device_tablewidget.item(row_index, COL_NAME).data(Qt.ItemDataRole.UserRole) for row_index in selected_rows if self.device_tablewidget.item(row_index, COL_NAME)}
        self.devices = [dev for dev in self.devices if dev.id not in ids_to_remove]
        connections_to_remove = [conn for conn in self.connections_result if conn[0].id in ids_to_remove or conn[2].id in ids_to_remove]
        if connections_to_remove:
            print(f"移除设备，同时移除 {len(connections_to_remove)} 条相关连接。")
            self.connections_result = [c for c in self.connections_result if c not in connections_to_remove]
        for row_index in selected_rows: self.device_tablewidget.removeRow(row_index)
        self._update_device_combos(); self.clear_results()
        self.node_positions = None; self.selected_node_id = None
        self._update_port_totals_display()

    @Slot()
    def clear_all_devices(self):
        user_confirmed = True
        if not self.suppress_confirmations:
            reply = QMessageBox.question(self, "确认", "确定要清空所有设备吗？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            user_confirmed = (reply == QMessageBox.StandardButton.Yes)
        if user_confirmed:
            self.devices = []; self.device_tablewidget.setRowCount(0); self.device_id_counter = 0;
            self._update_device_combos(); self.clear_results()
            self.node_positions = None; self.selected_node_id = None
            self._update_port_totals_display()
        else: print("用户取消清空所有设备。")

    @Slot()
    def clear_results(self):
        self.connections_result = []; self.fig = None; self.node_positions = None; self.selected_node_id = None
        self.connections_textedit.clear(); self.manual_connection_list.clear()
        self.mpl_canvas.axes.cla(); self.mpl_canvas.axes.text(0.5, 0.5, '点击“计算”生成图形', ha='center', va='center'); self.mpl_canvas.draw()
        self.export_list_button.setEnabled(False); self.export_topo_button.setEnabled(False); self.export_report_button.setEnabled(False)
        self.remove_manual_button.setEnabled(False); self.fill_mesh_button.setEnabled(False); self.fill_ring_button.setEnabled(False)
        for dev in self.devices: dev.reset_ports()
        self._update_device_table_connections()
        self._update_port_totals_display()
        self._update_connection_views()

    @Slot(QTableWidgetItem)
    def show_device_details_from_table(self, item):
         if not item: return
         row = item.row(); name_item = self.device_tablewidget.item(row, COL_NAME) # Use constant
         if not name_item: return
         dev_id = name_item.data(Qt.ItemDataRole.UserRole)
         dev = next((d for d in self.devices if d.id == dev_id), None)
         if not dev: return QMessageBox.critical(self, "错误", "无法找到所选设备的详细信息。")
         self._display_device_details_popup(dev)

    def _display_device_details_popup(self, dev):
        details = f"ID: {dev.id}\n名称: {dev.name}\n类型: {dev.type}\n"
        avail_ports = dev.get_all_available_ports()
        avail_lc_count = sum(1 for p in avail_ports if p.startswith(PORT_LC))
        avail_sfp_count = sum(1 for p in avail_ports if p.startswith(PORT_SFP))
        avail_mpo_ch_count = sum(1 for p in avail_ports if p.startswith(PORT_MPO))
        if dev.type in UHD_TYPES: details += f"{PORT_MPO} 端口总数: {dev.mpo_total}\n{PORT_LC} 端口总数: {dev.lc_total}\n可用 {PORT_MPO} 子通道: {avail_mpo_ch_count}\n可用 {PORT_LC} 端口: {avail_lc_count}\n"
        elif dev.type == DEV_MN: details += f"{PORT_SFP}+ 端口总数: {dev.sfp_total}\n可用 {PORT_SFP}+ 端口: {avail_sfp_count}\n"
        details += f"当前连接数 (估算): {dev.connections:.2f}\n"
        if dev.port_connections:
            details += "\n端口连接详情:\n"
            lc_conns = {p: t for p, t in dev.port_connections.items() if p.startswith(PORT_LC)}
            sfp_conns = {p: t for p, t in dev.port_connections.items() if p.startswith(PORT_SFP)}
            mpo_conns_grouped = defaultdict(dict)
            for p, t in dev.port_connections.items():
                if p.startswith(PORT_MPO): mpo_conns_grouped[p.split('-')[0]][p] = t
            if lc_conns:
                details += f"  {PORT_LC} 连接:\n"
                for port in sorted(lc_conns.keys()): details += f"    {port} -> {lc_conns[port]}\n"
            if sfp_conns:
                details += f"  {PORT_SFP}+ 连接:\n"
                for port in sorted(sfp_conns.keys()): details += f"    {port} -> {sfp_conns[port]}\n"
            if mpo_conns_grouped:
                details += f"  {PORT_MPO} 连接 (Breakout):\n"
                for base_port in sorted(mpo_conns_grouped.keys()):
                    details += f"    {base_port}:\n"
                    for port in sorted(mpo_conns_grouped[base_port].keys(), key=lambda x: int(x.split('-Ch')[-1])):
                         details += f"      {port} -> {mpo_conns_grouped[base_port][port]}\n"
        QMessageBox.information(self, f"设备详情 - {dev.name}", details)

    @Slot()
    def calculate_and_display(self):
        """计算连接并显示结果"""
        if not self.devices: QMessageBox.information(self, "提示", "请先添加设备。"); return
        self.node_positions = None; self.selected_node_id = None
        mode = self.topology_mode_combo.currentText()
        calculated_connections, final_device_state_map, error_message = [], {}, None
        for dev in self.devices: dev.reset_ports()
        if mode == "Mesh": print("使用改进的 Mesh 算法进行计算..."); calculated_connections, final_device_state_map = calculate_mesh_connections(self.devices)
        elif mode == "环形": print("使用环形算法进行计算..."); calculated_connections, final_device_state_map, error_message = calculate_ring_connections(self.devices)
        else: QMessageBox.critical(self, "错误", f"未知的计算模式: {mode}"); return
        if error_message: QMessageBox.warning(self, f"{mode}计算警告", error_message)
        self.connections_result = calculated_connections
        for dev in self.devices:
            if dev.id in final_device_state_map:
                final_state_dev = final_device_state_map[dev.id]
                dev.connections = final_state_dev.connections; dev.port_connections = final_state_dev.port_connections.copy()
            else: dev.reset_ports()
        self._update_connection_views(); self._update_device_table_connections(); self._update_device_combos()
        enable_fill = bool(self.connections_result) or any(dev.get_all_available_ports() for dev in self.devices)
        self.fill_mesh_button.setEnabled(enable_fill); self.fill_ring_button.setEnabled(enable_fill)

    def _update_connection_views(self):
        """更新连接列表、手动编辑列表和拓扑图 (包含高亮和拖动位置)"""
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
        port_totals = self._calculate_port_totals() # Calculate totals
        self.fig, calculated_pos = self.mpl_canvas.plot_topology(
            self.devices, self.connections_result,
            layout_algorithm=selected_layout, fixed_pos=self.node_positions,
            selected_node_id=self.selected_node_id, port_totals_dict=port_totals # Pass totals
        )
        if calculated_pos is not None and self.dragged_node_id is None:
             if self.node_positions is None: self.node_positions = calculated_pos
        has_results = bool(self.connections_result)
        has_figure = self.fig is not None and bool(self.devices)
        self.export_list_button.setEnabled(has_results)
        self.export_topo_button.setEnabled(has_figure)
        self.export_report_button.setEnabled(has_results and has_figure)
        self.remove_manual_button.setEnabled(has_results)

    def _update_device_table_connections(self):
        """更新设备表格中的连接数估算列 (使用常量)"""
        for row in range(self.device_tablewidget.rowCount()):
            name_item = self.device_tablewidget.item(row, COL_NAME)
            if name_item:
                dev_id = name_item.data(Qt.ItemDataRole.UserRole)
                device = next((d for d in self.devices if d.id == dev_id), None)
                if device:
                    conn_val = float(f"{device.connections:.2f}")
                    conn_item = NumericTableWidgetItem(f"{conn_val:.2f}"); conn_item.setData(Qt.ItemDataRole.UserRole + 1, conn_val); conn_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    conn_item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                    self.device_tablewidget.setItem(row, COL_CONN, conn_item)

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
             self._update_manual_port_options()
             self._update_port_totals_display()
             print(f"成功移除了 {connections_removed_count} 条连接。")
             enable_fill = bool(self.connections_result) or any(dev.get_all_available_ports() for dev in self.devices)
             self.fill_mesh_button.setEnabled(enable_fill); self.fill_ring_button.setEnabled(enable_fill)
        else: print("没有连接被移除。")

    @Slot()
    def add_manual_connection(self):
        # Uses check_port_compatibility which uses constants
        dev1_id = self.edit_dev1_combo.currentData(); dev2_id = self.edit_dev2_combo.currentData()
        port1_text = self.edit_port1_combo.currentText(); port2_text = self.edit_port2_combo.currentText()
        if dev1_id is None or dev2_id is None or port1_text == "选择端口..." or port2_text == "选择端口...": QMessageBox.warning(self, "选择不完整", "请选择两个设备和它们各自的端口。"); return
        if dev1_id == dev2_id: QMessageBox.warning(self, "选择错误", "不能将设备连接到自身。"); return
        dev1 = next((d for d in self.devices if d.id == dev1_id), None); dev2 = next((d for d in self.devices if d.id == dev2_id), None)
        if not dev1 or not dev2: QMessageBox.critical(self, "内部错误", "找不到选定的设备对象。"); return
        compatible, conn_type_str = check_port_compatibility(dev1.type, port1_text, dev2.type, port2_text)
        if not compatible:
            QMessageBox.warning(self, "连接无效", f"设备类型 '{dev1.type}' 的端口 '{port1_text}' 与 设备类型 '{dev2.type}' 的端口 '{port2_text}' 之间不允许连接。"); return
        if port1_text not in dev1.get_all_available_ports():
             QMessageBox.warning(self, "端口错误", f"端口 '{port1_text}' 在设备 '{dev1.name}' 上已被占用或无效。"); self._update_manual_port_options(); return
        if port2_text not in dev2.get_all_available_ports():
             QMessageBox.warning(self, "端口错误", f"端口 '{port2_text}' 在设备 '{dev2.name}' 上已被占用或无效。"); self._update_manual_port_options(); return
        if dev1.use_specific_port(port1_text, dev2.name):
            if dev2.use_specific_port(port2_text, dev1.name):
                self.node_positions = None; self.selected_node_id = None
                self.connections_result.append((dev1, port1_text, dev2, port2_text, conn_type_str))
                self._update_connection_views(); self._update_device_table_connections(); self._update_manual_port_options(); self._update_port_totals_display()
                print(f"成功添加手动连接: {dev1.name}[{port1_text}] <-> {dev2.name}[{port2_text}]")
                enable_fill = bool(self.connections_result) or any(dev.get_all_available_ports() for dev in self.devices); self.fill_mesh_button.setEnabled(enable_fill); self.fill_ring_button.setEnabled(enable_fill)
            else: dev1.return_port(port1_text); QMessageBox.critical(self, "错误", f"尝试占用端口 '{port2_text}' 失败 (设备 '{dev2.name}')，可能已被占用。"); self._update_manual_port_options()
        else: QMessageBox.critical(self, "错误", f"尝试占用端口 '{port1_text}' 失败 (设备 '{dev1.name}')，可能已被占用。"); self._update_manual_port_options()

    @Slot()
    def fill_remaining_mesh(self):
        """使用 Mesh 算法填充剩余的可用端口"""
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
            self._update_connection_views(); self._update_device_table_connections(); self._update_manual_port_options(); self._update_port_totals_display()
            QMessageBox.information(self, "填充完成", f"成功添加了 {len(new_connections)} 条新 Mesh 连接。")
        else: QMessageBox.information(self, "填充完成", "没有找到更多可以建立的 Mesh 连接。"); print("未找到可填充的新 Mesh 连接。")
        self.fill_mesh_button.setEnabled(False); self.fill_ring_button.setEnabled(False)

    @Slot()
    def fill_remaining_ring(self):
        """使用 Ring 算法填充剩余的可用端口"""
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
            self._update_connection_views(); self._update_device_table_connections(); self._update_manual_port_options(); self._update_port_totals_display()
            QMessageBox.information(self, "填充完成", f"成功添加了 {len(new_connections)} 条新环形连接段。")
        else: QMessageBox.information(self, "填充完成", "没有找到更多可以建立的环形连接段。"); print("未找到可填充的新环形连接段。")
        self.fill_mesh_button.setEnabled(False); self.fill_ring_button.setEnabled(False)

    @Slot()
    def on_layout_change(self):
        if self.connections_result or self.devices:
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
        user_confirmed_load = True
        if self.devices:
            if not self.suppress_confirmations:
                reply = QMessageBox.question(self, "确认", "加载配置将覆盖当前设备列表，确定吗？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
                user_confirmed_load = (reply == QMessageBox.StandardButton.Yes)
        if not user_confirmed_load: print("用户取消加载配置。"); return
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
            self._update_port_totals_display()
            self._update_connection_views()
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

    @Slot()
    def export_html_report(self):
        """导出包含拓扑图和连接列表的 HTML 报告"""
        if not self.devices or not self.mpl_canvas.fig:
            QMessageBox.warning(self, "无法导出", "请先添加设备并生成拓扑图（可能需要计算连接）。")
            return
        filepath, _ = QFileDialog.getSaveFileName(self, "导出 HTML 报告", "", "HTML 文件 (*.html);;所有文件 (*)")
        if not filepath: return
        try:
            buffer = io.BytesIO()
            self.mpl_canvas.fig.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
            buffer.seek(0)
            image_base64 = base64.b64encode(buffer.getvalue()).decode('utf-8')
            img_data_uri = f"data:image/png;base64,{image_base64}"
            connections_html = """
            <div class="mt-8"> <h2 class="text-lg font-semibold mb-3">连接列表</h2> <div class="overflow-x-auto bg-white rounded-lg shadow">
            <table class="min-w-full leading-normal"> <thead> <tr>
            <th class="px-5 py-3 border-b-2 border-gray-200 bg-gray-100 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">序号</th>
            <th class="px-5 py-3 border-b-2 border-gray-200 bg-gray-100 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">设备 1</th>
            <th class="px-5 py-3 border-b-2 border-gray-200 bg-gray-100 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">端口 1</th>
            <th class="px-5 py-3 border-b-2 border-gray-200 bg-gray-100 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">设备 2</th>
            <th class="px-5 py-3 border-b-2 border-gray-200 bg-gray-100 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">端口 2</th>
            <th class="px-5 py-3 border-b-2 border-gray-200 bg-gray-100 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">类型</th>
            </tr> </thead> <tbody>
            """
            if self.connections_result:
                for i, conn in enumerate(self.connections_result):
                    dev1, port1, dev2, port2, conn_type = conn
                    bg_class = "bg-white" if i % 2 == 0 else "bg-gray-50"
                    connections_html += f"""
                            <tr class="{bg_class}"> <td class="px-5 py-4 border-b border-gray-200 text-sm">{i+1}</td> <td class="px-5 py-4 border-b border-gray-200 text-sm">{dev1.name} ({dev1.type})</td> <td class="px-5 py-4 border-b border-gray-200 text-sm">{port1}</td> <td class="px-5 py-4 border-b border-gray-200 text-sm">{dev2.name} ({dev2.type})</td> <td class="px-5 py-4 border-b border-gray-200 text-sm">{port2}</td> <td class="px-5 py-4 border-b border-gray-200 text-sm">{conn_type}</td> </tr>
                    """
            else: connections_html += """ <tr> <td colspan="6" class="px-5 py-5 border-b border-gray-200 bg-white text-center text-sm text-gray-500">无连接</td> </tr> """
            connections_html += """ </tbody> </table> </div> </div> """
            html_content = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head> <meta charset="UTF-8"> <meta name="viewport" content="width=device-width, initial-scale=1.0"> <title>MediorNet 连接报告</title> <script src="https://cdn.tailwindcss.com"></script> <style> body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; }} @media print {{ body {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }} .bg-gray-100 {{ background-color: #f7fafc !important; }} .bg-gray-50 {{ background-color: #f9fafb !important; }} .border-b-2 {{ border-bottom-width: 2px !important; }} .border-gray-200 {{ border-color: #edf2f7 !important; }} .shadow {{ box-shadow: none !important; }} }} </style> </head>
<body class="bg-gray-100"> <div class="container mx-auto p-6 md:p-10 bg-white rounded-lg shadow-xl my-10 max-w-4xl"> <h1 class="text-2xl font-bold text-center mb-8 text-gray-700">MediorNet 连接报告</h1> <div class="mb-8"> <h2 class="text-lg font-semibold mb-3 text-gray-600">网络连接拓扑图</h2> <div class="flex justify-center p-4 border border-gray-200 rounded-lg bg-gray-50 shadow-inner"> <img src="{img_data_uri}" alt="网络拓扑图" style="max-width: 100%; height: auto;" class="rounded"> </div> </div> {connections_html} <div class="text-center text-xs text-gray-400 mt-10"> 报告生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} </div> </div> </body> </html>
            """
            with open(filepath, 'w', encoding='utf-8') as f: f.write(html_content)
            QMessageBox.information(self, "导出成功", f"HTML 报告已成功导出到:\n{filepath}")
        except Exception as e: QMessageBox.critical(self, "导出失败", f"导出 HTML 报告时发生错误:\n{e}"); print(f"导出 HTML 报告错误: {e}")

    @Slot(str)
    def filter_device_table(self, text):
        filter_text = text.lower()
        for row in range(self.device_tablewidget.rowCount()):
            match = False
            name_item = self.device_tablewidget.item(row, COL_NAME) # Use constant
            type_item = self.device_tablewidget.item(row, COL_TYPE) # Use constant
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

    # --- V48: 画布事件处理辅助函数 ---
    def _get_node_at_event(self, event):
        """辅助函数：查找鼠标事件位置下的节点 ID"""
        if event.inaxes != self.mpl_canvas.axes or not self.node_positions: return None
        x, y = event.xdata, event.ydata
        if x is None or y is None: return None
        clicked_node_id = None; min_dist_sq = float('inf')
        xlim = self.mpl_canvas.axes.get_xlim(); ylim = self.mpl_canvas.axes.get_ylim()
        # 简单的固定阈值或基于视图范围的动态阈值
        threshold_dist_sq = ((xlim[1]-xlim[0])**2 + (ylim[1]-ylim[0])**2) * (0.03**2)
        for node_id, (nx, ny) in self.node_positions.items():
            dist_sq = (x - nx)**2 + (y - ny)**2
            if dist_sq < min_dist_sq and dist_sq < threshold_dist_sq:
                 min_dist_sq = dist_sq; clicked_node_id = node_id
        return clicked_node_id

    def _start_node_drag(self, node_id, event_xdata, event_ydata):
        """辅助函数：开始节点拖动"""
        self.dragged_node_id = node_id
        self.connecting_node_id = None
        nx, ny = self.node_positions[node_id]
        self.drag_offset = (event_xdata - nx, event_ydata - ny)
        if self.selected_node_id != node_id:
            self.selected_node_id = node_id
            print(f"选中节点 (准备拖动): {self.selected_node_id}")
            self._update_connection_views() # 更新高亮
        else:
            print(f"开始拖动节点: {self.selected_node_id}")

    def _start_connection_drag(self, node_id):
        """辅助函数：开始连接拖动"""
        print(f"开始连接拖动: 从 {node_id}")
        self.connecting_node_id = node_id
        self.dragged_node_id = None
        if self.selected_node_id != node_id:
            self.selected_node_id = node_id
            self._update_connection_views() # 更新高亮

    def _handle_background_press(self):
        """辅助函数：处理画布背景点击"""
        needs_redraw = self.selected_node_id is not None or self.dragged_node_id is not None or self.connecting_node_id is not None
        self.selected_node_id = None; self.dragged_node_id = None; self.connecting_node_id = None
        if needs_redraw:
            print("清除选中/状态 (点击背景)")
            self._update_connection_views()

    def _end_node_drag(self):
        """辅助函数：结束节点拖动"""
        if self.dragged_node_id is not None: # 检查以防万一
            print(f"结束拖动节点: {self.dragged_node_id}")
            self.dragged_node_id = None
            self.mpl_canvas.draw_idle() # 确保最终状态

    def _end_connection_drag(self, event):
        """辅助函数：结束连接拖动并尝试创建连接"""
        start_node_id = self.connecting_node_id
        self.connecting_node_id = None # 重置状态

        # 清理临时线
        if self.connection_line:
            try: self.connection_line.remove(); self.connection_line = None
            except ValueError: pass
            self.mpl_canvas.draw_idle()

        target_node_id = self._get_node_at_event(event) # 使用辅助函数查找目标

        if target_node_id is not None and target_node_id != start_node_id:
            start_device = next((d for d in self.devices if d.id == start_node_id), None)
            target_device = next((d for d in self.devices if d.id == target_node_id), None)
            if start_device and target_device:
                print(f"尝试连接: {start_device.name} -> {target_device.name}")
                dev1_copy = copy.deepcopy(start_device); dev2_copy = copy.deepcopy(target_device)
                port1, port2, conn_type = _find_best_single_link(dev1_copy, dev2_copy)
                if port1 and port2:
                    actual_start_dev, actual_target_dev = start_device, target_device
                    start_port, target_port = port1, port2
                    if dev1_copy.id != start_node_id: start_port, target_port = port2, port1
                    compatible, final_conn_type = check_port_compatibility(actual_start_dev.type, start_port, actual_target_dev.type, target_port)
                    if not compatible:
                        QMessageBox.warning(self, "连接错误", f"自动选择的端口 {start_port} ({actual_start_dev.type}) 和 {target_port} ({actual_target_dev.type}) 之间存在兼容性问题。"); self._update_manual_port_options(); return
                    if start_device.use_specific_port(start_port, target_device.name):
                         if target_device.use_specific_port(target_port, start_device.name):
                             self.node_positions = None; self.selected_node_id = None
                             self.connections_result.append((start_device, start_port, target_device, target_port, final_conn_type))
                             self._update_connection_views(); self._update_device_table_connections(); self._update_manual_port_options(); self._update_port_totals_display()
                             print(f"成功通过拖拽添加连接: {start_device.name}[{start_port}] <-> {target_device.name}[{target_port}]")
                             enable_fill = bool(self.connections_result) or any(dev.get_all_available_ports() for dev in self.devices); self.fill_mesh_button.setEnabled(enable_fill); self.fill_ring_button.setEnabled(enable_fill)
                         else: start_device.return_port(start_port); QMessageBox.warning(self, "连接失败", f"无法在设备 '{target_device.name}' 上使用端口 '{target_port}' (可能已被占用)。"); self._update_manual_port_options()
                    else: QMessageBox.warning(self, "连接失败", f"无法在设备 '{start_device.name}' 上使用端口 '{start_port}' (可能已被占用)。"); self._update_manual_port_options()
                else: QMessageBox.information(self, "连接失败", f"在 {start_device.name} 和 {target_device.name} 之间未找到兼容的可用端口。")
            else: print("错误：找不到拖放连接的设备对象。")
        else: print("连接拖动取消或目标无效。")


    # --- V48: Refactored Event Handlers ---
    @Slot(object)
    def on_canvas_press(self, event):
        """处理画布上的鼠标按下事件 (Refactored)"""
        if self.connection_line: # Reset line first
            try: self.connection_line.remove(); self.connection_line = None
            except ValueError: pass

        clicked_node_id = self._get_node_at_event(event)

        # Handle clicks outside axes or data range first
        if event.inaxes != self.mpl_canvas.axes or event.xdata is None or event.ydata is None:
            self._handle_background_press(); return

        modifiers = QGuiApplication.keyboardModifiers()
        is_shift_pressed = modifiers == Qt.KeyboardModifier.ShiftModifier

        if event.dblclick:
            self.dragged_node_id = None; self.connecting_node_id = None # Cancel any drag state
            if clicked_node_id is not None:
                device = next((d for d in self.devices if d.id == clicked_node_id), None)
                if device: self._display_device_details_popup(device)
        elif event.button == 1: # Left button press
            if clicked_node_id is not None: # Clicked on a node
                if is_shift_pressed: self._start_connection_drag(clicked_node_id)
                else: self._start_node_drag(clicked_node_id, event.xdata, event.ydata)
            else: # Clicked on background
                self._handle_background_press()

    @Slot(object)
    def on_canvas_motion(self, event):
        """处理画布上的鼠标移动事件 (Refactored)"""
        if event.inaxes != self.mpl_canvas.axes: return # Ignore motion outside axes
        x, y = event.xdata, event.ydata
        if x is None or y is None: return # Ignore motion outside data range

        # Handle Node Dragging
        if self.dragged_node_id is not None and event.button == 1:
            new_x = x - self.drag_offset[0]; new_y = y - self.drag_offset[1]
            self.node_positions[self.dragged_node_id] = (new_x, new_y)
            self._update_connection_views() # Redraw all for node move

        # Handle Connection Line Dragging
        elif self.connecting_node_id is not None and event.button == 1:
            start_pos = self.node_positions[self.connecting_node_id]
            # Remove previous line if it exists
            if self.connection_line:
                try: self.connection_line.remove(); self.connection_line = None
                except ValueError: pass
            # Draw new line
            self.connection_line = Line2D([start_pos[0], x], [start_pos[1], y], ls='--', c='gray', lw=1.5, transform=self.mpl_canvas.axes.transData, zorder=10)
            self.mpl_canvas.axes.add_line(self.connection_line)
            self.mpl_canvas.draw_idle() # Use draw_idle for smoother line drawing

    @Slot(object)
    def on_canvas_release(self, event):
        """处理画布上的鼠标释放事件 (Refactored)"""
        if event.button == 1: # Left button release
            if self.dragged_node_id is not None: self._end_node_drag()
            elif self.connecting_node_id is not None: self._end_connection_drag(event)
    # --- 结束事件处理 ---

# --- 程序入口 ---
if __name__ == "__main__":
    # 确保在 Windows 上高 DPI 显示正常 (如果需要)
    # try:
    #     from PySide6 import QtWinExtras
    #     QtWinExtras.QtWin.setCurrentProcessExplicitAppUserModelID('myappid') # 设置 AppUserModelID
    # except ImportError:
    #     pass # 不是 Windows 或没有 QtWinExtras
    # QApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    # QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)

    import datetime # 确保导入
    app = QApplication(sys.argv)
    app.setStyleSheet(APP_STYLE)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

