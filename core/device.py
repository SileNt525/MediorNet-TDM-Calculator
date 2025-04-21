# -*- coding: utf-8 -*-
"""
core/device.py

定义 MediorNet 设备的数据结构 (Device 类) 和相关常量、辅助函数。
"""
import random
from collections import defaultdict

# --- 设备和端口类型常量 ---
DEV_UHD = 'MicroN UHD'
DEV_HORIZON = 'HorizoN'
DEV_MN = 'MicroN'
UHD_TYPES = [DEV_UHD, DEV_HORIZON] # 辅助列表，用于判断是否为 UHD/HorizoN 类型

PORT_MPO = 'MPO'
PORT_LC = 'LC'
PORT_SFP = 'SFP'
PORT_UNKNOWN = 'Unknown'
# --- 结束常量 ---

# --- 辅助函数 ---
def get_port_type_from_name(port_name):
    """
    根据端口名称字符串返回端口类型常量。

    Args:
        port_name (str): 端口名称 (例如 "LC1", "MPO1-Ch2", "SFP3").

    Returns:
        str: 端口类型常量 (PORT_LC, PORT_MPO, PORT_SFP) 或 PORT_UNKNOWN。
    """
    if not isinstance(port_name, str):
        return PORT_UNKNOWN
    if port_name.startswith(PORT_LC):
        return PORT_LC
    if port_name.startswith(PORT_SFP):
        return PORT_SFP
    if port_name.startswith(PORT_MPO):
        return PORT_MPO
    return PORT_UNKNOWN
# --- 结束辅助函数 ---

# --- 数据结构 ---
class Device:
    """代表一个 MediorNet 设备及其端口状态。"""
    def __init__(self, id, name, type, mpo_ports=0, lc_ports=0, sfp_ports=0):
        """
        初始化 Device 对象。

        Args:
            id (int): 设备的唯一标识符。
            name (str): 设备名称。
            type (str): 设备类型 (使用常量 DEV_UHD, DEV_HORIZON, DEV_MN)。
            mpo_ports (int): MPO 端口的数量 (仅适用于 UHD/HorizoN)。
            lc_ports (int): LC 端口的数量 (仅适用于 UHD/HorizoN)。
            sfp_ports (int): SFP+ 端口的数量 (仅适用于 MicroN)。
        """
        self.id = id
        self.name = name
        self.type = type
        # 确保端口数是整数
        self.mpo_total = int(mpo_ports) if self.type in UHD_TYPES else 0
        self.lc_total = int(lc_ports) if self.type in UHD_TYPES else 0
        self.sfp_total = int(sfp_ports) if self.type == DEV_MN else 0
        self.reset_ports()

    def reset_ports(self):
        """重置端口连接状态和计数。在清除连接或重新计算时调用。"""
        # 使用浮点数以精确表示 MPO 连接 (每个 Breakout 通道算 0.25)
        self.connections = 0.0
        # 字典：存储每个已连接端口及其连接的对端设备名称
        # 键: 本设备端口名 (例如 "MPO1-Ch3")
        # 值: 对端设备名 (例如 "DeviceB")
        self.port_connections = {}

    def get_all_possible_ports(self):
        """
        获取此设备所有可能的端口名称列表 (无论是否已连接)。

        Returns:
            list[str]: 所有可能的端口名称字符串列表。
        """
        ports = []
        # LC 端口 (仅 UHD/HorizoN)
        ports.extend([f"{PORT_LC}{i+1}" for i in range(self.lc_total)])
        # SFP+ 端口 (仅 MicroN)
        ports.extend([f"{PORT_SFP}{i+1}" for i in range(self.sfp_total)])
        # MPO 端口 (仅 UHD/HorizoN), 每个 MPO 有 4 个 Breakout 子通道
        for i in range(self.mpo_total):
            base = f"{PORT_MPO}{i+1}"
            ports.extend([f"{base}-Ch{j+1}" for j in range(4)])
        return ports

    def get_all_available_ports(self):
        """
        获取此设备当前所有 *可用* (未连接) 的端口名称列表。

        Returns:
            list[str]: 可用端口名称字符串列表。
        """
        all_ports = self.get_all_possible_ports()
        used_ports = set(self.port_connections.keys())
        available = [p for p in all_ports if p not in used_ports]
        return available

    def use_specific_port(self, port_name, target_device_name):
        """
        标记指定端口为已使用，并记录连接的目标设备。

        Args:
            port_name (str): 要使用的本设备端口名称。
            target_device_name (str): 连接的目标设备名称。

        Returns:
            bool: 如果端口可用且成功标记为已使用，则返回 True；否则返回 False。
        """
        # 检查端口是否属于该设备且当前是否可用
        if port_name in self.get_all_possible_ports() and port_name not in self.port_connections:
            self.port_connections[port_name] = target_device_name
            # 更新连接计数
            port_type = get_port_type_from_name(port_name)
            if port_type == PORT_MPO:
                self.connections += 0.25 # MPO Breakout 算 0.25 个连接
            elif port_type in [PORT_LC, PORT_SFP]:
                self.connections += 1.0 # LC 和 SFP 算 1 个连接
            else:
                print(f"警告: 在设备 {self.name} 上使用未知类型的端口 {port_name} 进行连接计数。")
            return True
        # 如果端口无效或已被占用，返回 False
        print(f"调试: 尝试使用端口 {self.name}[{port_name}] 失败。可能原因：无效端口或已被占用 ({port_name in self.port_connections})")
        return False

    def return_port(self, port_name):
        """
        释放指定端口，使其变为可用状态，并更新连接计数。

        Args:
            port_name (str): 要释放的端口名称。
        """
        # 检查端口是否确实在已连接列表中
        if port_name in self.port_connections:
            target = self.port_connections.pop(port_name) # 从已用端口字典中移除
            # 减少连接计数
            port_type = get_port_type_from_name(port_name)
            if port_type == PORT_MPO:
                self.connections -= 0.25
            elif port_type in [PORT_LC, PORT_SFP]:
                self.connections -= 1.0
            else:
                 print(f"警告: 在设备 {self.name} 上归还未知类型的端口 {port_name} 进行连接计数。")

            # 确保连接数不为负
            self.connections = max(0.0, self.connections)
            print(f"调试: 端口 {self.name}[{port_name}] 已释放 (之前连接到 {target})。当前连接数: {self.connections:.2f}")
        else:
            # 如果端口本来就没被记录为使用，则无需操作，可能是一个状态错误或重复释放
            print(f"调试: 尝试释放端口 {self.name}[{port_name}]，但它不在已连接列表中。")


    def get_specific_available_port(self, port_type_prefix):
        """
        按顺序查找指定前缀的第一个可用端口（*不* 标记为已用）。
        用于连接计算逻辑中探测可用端口。

        Args:
            port_type_prefix (str): 端口类型前缀 (例如 PORT_LC, PORT_MPO, PORT_SFP)。

        Returns:
            str | None: 找到的第一个可用端口的名称，如果该类型无可用端口则返回 None。
        """
        possible_ports = []
        # 根据前缀生成可能的端口列表，并排序以保证查找顺序
        if port_type_prefix == PORT_LC and self.type in UHD_TYPES:
            possible_ports = sorted(
                [f"{PORT_LC}{i+1}" for i in range(self.lc_total)],
                key=lambda x: int(x[len(PORT_LC):])
            )
        elif port_type_prefix == PORT_MPO and self.type in UHD_TYPES:
            for i in range(self.mpo_total):
                base = f"{PORT_MPO}{i+1}"
                possible_ports.extend(sorted(
                    [f"{base}-Ch{j+1}" for j in range(4)],
                    key=lambda x: (int(x[len(PORT_MPO):].split('-Ch')[0]), int(x.split('-Ch')[-1])) # 先按 MPO 序号，再按通道号排序
                ))
        elif port_type_prefix == PORT_SFP and self.type == DEV_MN:
            possible_ports = sorted(
                [f"{PORT_SFP}{i+1}" for i in range(self.sfp_total)],
                key=lambda x: int(x[len(PORT_SFP):])
            )

        # 检查哪个端口是可用的
        used_ports = set(self.port_connections.keys())
        for port in possible_ports:
            if port not in used_ports:
                return port # 返回第一个找到的可用端口
        return None # 如果没有找到该类型的可用端口

    def to_dict(self):
        """将设备对象转换为字典，用于保存配置。"""
        return {
            'id': self.id,
            'name': self.name,
            'type': self.type,
            'mpo_ports': self.mpo_total,
            'lc_ports': self.lc_total,
            'sfp_ports': self.sfp_total
            # 注意：不保存 port_connections 和 connections 状态，这些由连接列表推断
        }

    @classmethod
    def from_dict(cls, data):
        """
        从字典创建设备对象，用于加载配置。

        Args:
            data (dict): 包含设备信息的字典。

        Returns:
            Device: 创建的 Device 对象。
        """
        # 提供默认 ID 以兼容可能缺少 ID 的旧格式
        data.setdefault('id', random.randint(10000, 99999))
        return cls(
            id=data['id'],
            name=data['name'],
            type=data['type'],
            mpo_ports=data.get('mpo_ports', 0), # 使用 .get 提供默认值
            lc_ports=data.get('lc_ports', 0),
            sfp_ports=data.get('sfp_ports', 0)
        )

    def __repr__(self):
        """返回对象的字符串表示形式，方便调试。"""
        return f"Device(id={self.id}, name='{self.name}', type='{self.type}')"

    def __eq__(self, other):
        """比较两个 Device 对象是否相等（基于 ID）。"""
        if not isinstance(other, Device):
            return NotImplemented
        return self.id == other.id

    def __hash__(self):
        """允许 Device 对象用作字典键或集合元素（基于 ID）。"""
        return hash(self.id)

