# -*- coding: utf-8 -*-
"""
core/network_manager.py

定义 NetworkManager 类，负责管理设备、连接、计算和项目状态。
这是应用程序的核心模型。
"""
import copy
import itertools
import random
import json
import networkx as nx
from typing import List, Dict, Tuple, Optional, Set, Any # 导入 Any
from collections import defaultdict # <--- **修复: 添加了 defaultdict 导入**

# 从同级目录的 device 模块导入
from .device import (
    Device,
    DEV_UHD, DEV_HORIZON, DEV_MN, UHD_TYPES,
    PORT_MPO, PORT_LC, PORT_SFP, PORT_UNKNOWN,
    get_port_type_from_name
)

# 定义连接元组的类型别名，提高可读性
# (源设备, 源端口名, 目标设备, 目标端口名, 连接类型描述)
ConnectionType = Tuple[Device, str, Device, str, str]

class NetworkManager:
    """管理 MediorNet 设备网络状态和连接的核心类。"""

    def __init__(self):
        """初始化 NetworkManager。"""
        self.devices: List[Device] = []            # 当前系统中的设备列表
        self.connections: List[ConnectionType] = [] # 当前系统中的连接列表
        self.graph: nx.Graph = nx.Graph()          # NetworkX 图对象，用于拓扑可视化
        self.device_id_counter: int = 0            # 用于生成唯一的设备 ID

    # --- 设备管理 ---

    def add_device(self, name: str, type: str, mpo_ports: int, lc_ports: int, sfp_ports: int) -> Optional[Device]:
        """
        添加一个新设备到管理器中。

        Args:
            name (str): 设备名称。
            type (str): 设备类型 (来自 device 模块的常量)。
            mpo_ports (int): MPO 端口数。
            lc_ports (int): LC 端口数。
            sfp_ports (int): SFP+ 端口数。

        Returns:
            Optional[Device]: 成功添加则返回新创建的 Device 对象，否则返回 None (例如名称冲突)。
        """
        if any(dev.name == name for dev in self.devices):
            print(f"错误: 设备名称 '{name}' 已存在。")
            return None
        self.device_id_counter += 1
        new_device = Device(self.device_id_counter, name, type, mpo_ports, lc_ports, sfp_ports)
        self.devices.append(new_device)
        self._update_graph() # 添加节点到图中
        print(f"设备已添加: {new_device}")
        return new_device

    def remove_device(self, device_id: int) -> bool:
        """
        根据 ID 移除一个设备及其所有相关连接。

        Args:
            device_id (int): 要移除的设备的 ID。

        Returns:
            bool: 如果成功找到并移除设备则返回 True，否则返回 False。
        """
        device_to_remove = self.get_device_by_id(device_id)
        if not device_to_remove:
            print(f"错误: 找不到 ID 为 {device_id} 的设备进行移除。")
            return False

        # 1. 移除与该设备相关的所有连接
        connections_removed_count = self._remove_connections_for_device(device_id) # 这个方法内部会更新图
        print(f"移除设备 {device_to_remove.name} 时，移除了 {connections_removed_count} 条相关连接。")

        # 2. 从设备列表中移除设备
        self.devices.remove(device_to_remove)

        # 3. 从图中移除节点 (如果 _remove_connections_for_device 没有处理边导致节点孤立的话)
        # _update_graph 会处理节点的移除，所以这里不需要显式移除
        # if self.graph.has_node(device_id):
        #     self.graph.remove_node(device_id)
        self._update_graph() # 确保图状态正确

        print(f"设备已移除: {device_to_remove.name}")
        return True

    def get_device_by_id(self, device_id: int) -> Optional[Device]:
        """根据 ID 获取设备对象。"""
        return next((dev for dev in self.devices if dev.id == device_id), None)

    def get_device_by_name(self, name: str) -> Optional[Device]:
        """根据名称获取设备对象。"""
        return next((dev for dev in self.devices if dev.name == name), None)

    def get_all_devices(self) -> List[Device]:
        """获取所有设备的列表。"""
        return self.devices

    def update_device(self, device_id: int, new_name: Optional[str] = None,
                      new_mpo: Optional[int] = None, new_lc: Optional[int] = None, new_sfp: Optional[int] = None) -> bool:
        """
        更新指定 ID 设备的信息。端口数量的修改会导致连接被清空。

        Args:
            device_id (int): 要更新的设备 ID。
            new_name (Optional[str]): 新的设备名称。
            new_mpo (Optional[int]): 新的 MPO 端口数量。
            new_lc (Optional[int]): 新的 LC 端口数量。
            new_sfp (Optional[int]): 新的 SFP+ 端口数量。

        Returns:
            bool: 更新是否成功。
        """
        device = self.get_device_by_id(device_id)
        if not device:
            print(f"错误: 找不到 ID 为 {device_id} 的设备进行更新。")
            return False

        port_changed = False
        name_changed = False
        old_name = device.name

        # 更新名称
        if new_name is not None and new_name.strip() and new_name != device.name:
            if any(d.name == new_name and d.id != device_id for d in self.devices):
                print(f"错误: 设备名称 '{new_name}' 已存在。")
                return False
            print(f"设备 '{device.name}' 重命名为 '{new_name}'")
            device.name = new_name
            name_changed = True

        # 检查并更新端口数量 (需要设备类型匹配)
        if device.type in UHD_TYPES:
            if new_mpo is not None and new_mpo >= 0 and new_mpo != device.mpo_total:
                print(f"设备 '{device.name}' MPO 端口从 {device.mpo_total} 修改为 {new_mpo}")
                device.mpo_total = new_mpo
                port_changed = True
            if new_lc is not None and new_lc >= 0 and new_lc != device.lc_total:
                print(f"设备 '{device.name}' LC 端口从 {device.lc_total} 修改为 {new_lc}")
                device.lc_total = new_lc
                port_changed = True
        elif device.type == DEV_MN:
            if new_sfp is not None and new_sfp >= 0 and new_sfp != device.sfp_total:
                print(f"设备 '{device.name}' SFP+ 端口从 {device.sfp_total} 修改为 {new_sfp}")
                device.sfp_total = new_sfp
                port_changed = True

        # 如果端口数量改变，必须清除所有连接
        if port_changed:
            print(f"警告: 设备 '{device.name}' 的端口数量已更改，将清除所有现有连接。")
            self.clear_connections() # 清除所有连接并重置所有设备端口状态
            return True # 端口修改成功（即使清空了连接）

        # 如果只是名称改变，需要更新连接中的目标名称和图标签
        if name_changed and not port_changed:
             for other_dev in self.devices:
                 if other_dev.id != device.id:
                     ports_to_update = {p: new_name for p, target in other_dev.port_connections.items() if target == old_name}
                     other_dev.port_connections.update(ports_to_update)
             # 更新图中的标签
             self._update_graph()
             return True

        # 如果没有任何改变
        if not name_changed and not port_changed:
            return True # 没有需要更新的，也算成功

        return False # 一般不会到这里

    def clear_all_devices_and_connections(self):
        """清空所有设备和连接。"""
        self.devices = []
        self.connections = []
        self.graph.clear()
        self.device_id_counter = 0
        print("所有设备和连接已清空。")

    # --- 连接管理 ---

    def add_connection(self, dev1_id: int, port1_name: str, dev2_id: int, port2_name: str) -> Optional[ConnectionType]:
        """
        在两个设备之间添加一条手动连接。

        Args:
            dev1_id (int): 设备 1 的 ID。
            port1_name (str): 设备 1 的端口名称。
            dev2_id (int): 设备 2 的 ID。
            port2_name (str): 设备 2 的端口名称。

        Returns:
            Optional[ConnectionType]: 如果连接成功添加，返回连接元组；否则返回 None。
        """
        dev1 = self.get_device_by_id(dev1_id)
        dev2 = self.get_device_by_id(dev2_id)

        if not dev1 or not dev2:
            print(f"错误: 添加连接时找不到设备 ID {dev1_id} 或 {dev2_id}。")
            return None
        if dev1_id == dev2_id:
             print(f"错误: 不能将设备 {dev1.name} 连接到自身。")
             return None

        # 检查端口是否有效且可用
        if port1_name not in dev1.get_all_possible_ports():
             print(f"错误: 端口 '{port1_name}' 在设备 '{dev1.name}' 上无效。")
             return None
        if port2_name not in dev2.get_all_possible_ports():
             print(f"错误: 端口 '{port2_name}' 在设备 '{dev2.name}' 上无效。")
             return None
        if port1_name not in dev1.get_all_available_ports():
             print(f"错误: 端口 '{port1_name}' 在设备 '{dev1.name}' 上已被占用 ({dev1.port_connections.get(port1_name)})。")
             return None
        if port2_name not in dev2.get_all_available_ports():
             print(f"错误: 端口 '{port2_name}' 在设备 '{dev2.name}' 上已被占用 ({dev2.port_connections.get(port2_name)})。")
             return None

        # 检查兼容性
        compatible, conn_type_str = self.check_port_compatibility(dev1_id, port1_name, dev2_id, port2_name)
        if not compatible:
            print(f"错误: 端口 {dev1.name}[{port1_name}] ({dev1.type}) 与 {dev2.name}[{port2_name}] ({dev2.type}) 不兼容。")
            return None

        # 尝试占用端口并添加连接
        if dev1.use_specific_port(port1_name, dev2.name):
            if dev2.use_specific_port(port2_name, dev1.name):
                connection: ConnectionType = (dev1, port1_name, dev2, port2_name, conn_type_str)
                self.connections.append(connection)
                # 更新图
                self._update_graph()
                print(f"连接已添加: {dev1.name}[{port1_name}] <-> {dev2.name}[{port2_name}] ({conn_type_str})")
                return connection
            else:
                # 如果占用 dev2 端口失败，需要回滚 dev1 的端口占用
                dev1.return_port(port1_name)
                print(f"错误: 占用端口 {dev2.name}[{port2_name}] 失败，回滚操作。")
                return None
        else:
            print(f"错误: 占用端口 {dev1.name}[{port1_name}] 失败。")
            return None

    def remove_connection(self, dev1_id: int, port1_name: str, dev2_id: int, port2_name: str) -> bool:
        """
        移除指定的连接。

        Args:
            dev1_id (int): 设备 1 的 ID。
            port1_name (str): 设备 1 的端口名称。
            dev2_id (int): 设备 2 的 ID。
            port2_name (str): 设备 2 的端口名称。

        Returns:
            bool: 如果成功找到并移除连接则返回 True，否则返回 False。
        """
        found_index = -1
        # 查找匹配的连接 (双向检查)
        for i, conn in enumerate(self.connections):
            c_dev1, c_port1, c_dev2, c_port2, _ = conn
            if (c_dev1.id == dev1_id and c_port1 == port1_name and c_dev2.id == dev2_id and c_port2 == port2_name) or \
               (c_dev1.id == dev2_id and c_port1 == port2_name and c_dev2.id == dev1_id and c_port2 == port1_name):
                found_index = i
                break

        if found_index != -1:
            removed_conn = self.connections.pop(found_index)
            dev1, port1, dev2, port2, _ = removed_conn
            # 释放两端的端口 (需要获取最新的设备对象引用)
            actual_dev1 = self.get_device_by_id(dev1.id)
            actual_dev2 = self.get_device_by_id(dev2.id)
            if actual_dev1: actual_dev1.return_port(port1)
            else: print(f"警告: 移除连接时找不到设备 ID {dev1.id}")
            if actual_dev2: actual_dev2.return_port(port2)
            else: print(f"警告: 移除连接时找不到设备 ID {dev2.id}")

            # 更新图
            self._update_graph()
            print(f"连接已移除: {dev1.name}[{port1}] <-> {dev2.name}[{port2}]")
            return True
        else:
            print(f"错误: 未找到要移除的连接: {dev1_id}[{port1_name}] <-> {dev2_id}[{port2_name}]")
            return False

    def _remove_connections_for_device(self, device_id: int) -> int:
        """
        移除与指定设备相关的所有连接（内部辅助方法）。

        Args:
            device_id (int): 设备的 ID。

        Returns:
            int: 被移除的连接数量。
        """
        connections_to_remove: List[ConnectionType] = []
        indices_to_remove: Set[int] = set()

        # 查找所有涉及该设备的连接
        for i, conn in enumerate(self.connections):
            if conn[0].id == device_id or conn[2].id == device_id:
                connections_to_remove.append(conn)
                indices_to_remove.add(i)

        removed_count = len(connections_to_remove)
        if removed_count > 0:
            # 释放涉及的端口
            for conn in connections_to_remove:
                dev1, port1, dev2, port2, _ = conn
                # 获取最新的设备对象引用来释放端口
                actual_dev1 = self.get_device_by_id(dev1.id)
                actual_dev2 = self.get_device_by_id(dev2.id)
                if actual_dev1: actual_dev1.return_port(port1)
                if actual_dev2: actual_dev2.return_port(port2)


            # 从主连接列表中移除
            # 从后往前删除，避免索引问题
            for index in sorted(list(indices_to_remove), reverse=True):
                self.connections.pop(index)

            # 更新图 (移除相关边)
            self._update_graph()

        return removed_count

    def clear_connections(self):
        """清除所有连接，并重置所有设备的端口状态。"""
        if not self.connections:
            return # 没有连接可清除

        print(f"正在清除 {len(self.connections)} 条连接...")
        # 重置所有设备的端口状态
        for dev in self.devices:
            dev.reset_ports()
        # 清空连接列表
        self.connections = []
        # 更新图 (移除所有边)
        self._update_graph()
        print("所有连接已清除，设备端口状态已重置。")

    def get_all_connections(self) -> List[ConnectionType]:
        """获取当前所有连接的列表。"""
        return self.connections

    # --- 计算逻辑 ---

    def _find_best_single_link(self, dev1_copy: Device, dev2_copy: Device) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        """
        辅助函数：查找两个设备副本之间最高优先级的单个可用连接。
        优先级: LC-LC > MPO-MPO > SFP-SFP > MPO-SFP
        此方法会修改传入的设备副本状态 (dev1_copy, dev2_copy)。

        Args:
            dev1_copy (Device): 设备 1 的深拷贝副本。
            dev2_copy (Device): 设备 2 的深拷贝副本。

        Returns:
            Tuple[Optional[str], Optional[str], Optional[str]]:
                (dev1_copy 的端口名, dev2_copy 的端口名, 连接类型描述) 或 (None, None, None)。
        """
        is_uhd1 = dev1_copy.type in UHD_TYPES
        is_uhd2 = dev2_copy.type in UHD_TYPES
        is_mn1 = dev1_copy.type == DEV_MN
        is_mn2 = dev2_copy.type == DEV_MN

        # 尝试 LC-LC (仅 UHD/HorizoN 之间)
        if is_uhd1 and is_uhd2:
            port1 = dev1_copy.get_specific_available_port(PORT_LC)
            if port1:
                port2 = dev2_copy.get_specific_available_port(PORT_LC)
                if port2:
                    # 找到可用 LC 对，标记为使用并返回
                    if dev1_copy.use_specific_port(port1, dev2_copy.name) and dev2_copy.use_specific_port(port2, dev1_copy.name):
                         return port1, port2, f"{PORT_LC}-{PORT_LC} (100G)"
                    else: # 理论上不应发生，因为 get_specific_available_port 检查过
                         print(f"警告: _find_best_single_link 中 LC 端口占用失败 {dev1_copy.name}[{port1}] 或 {dev2_copy.name}[{port2}]")
                         if port1 in dev1_copy.port_connections: dev1_copy.return_port(port1) # 回滚
                         if port2 in dev2_copy.port_connections: dev2_copy.return_port(port2) # 回滚

        # 尝试 MPO-MPO (仅 UHD/HorizoN 之间)
        if is_uhd1 and is_uhd2:
            port1 = dev1_copy.get_specific_available_port(PORT_MPO)
            if port1:
                port2 = dev2_copy.get_specific_available_port(PORT_MPO)
                if port2:
                    if dev1_copy.use_specific_port(port1, dev2_copy.name) and dev2_copy.use_specific_port(port2, dev1_copy.name):
                        return port1, port2, f"{PORT_MPO}-{PORT_MPO} (25G)"
                    else:
                         print(f"警告: _find_best_single_link 中 MPO 端口占用失败 {dev1_copy.name}[{port1}] 或 {dev2_copy.name}[{port2}]")
                         if port1 in dev1_copy.port_connections: dev1_copy.return_port(port1)
                         if port2 in dev2_copy.port_connections: dev2_copy.return_port(port2)

        # 尝试 SFP-SFP (仅 MicroN 之间)
        if is_mn1 and is_mn2:
            port1 = dev1_copy.get_specific_available_port(PORT_SFP)
            if port1:
                port2 = dev2_copy.get_specific_available_port(PORT_SFP)
                if port2:
                    if dev1_copy.use_specific_port(port1, dev2_copy.name) and dev2_copy.use_specific_port(port2, dev1_copy.name):
                        return port1, port2, f"{PORT_SFP}-{PORT_SFP} (10G)"
                    else:
                        print(f"警告: _find_best_single_link 中 SFP 端口占用失败 {dev1_copy.name}[{port1}] 或 {dev2_copy.name}[{port2}]")
                        if port1 in dev1_copy.port_connections: dev1_copy.return_port(port1)
                        if port2 in dev2_copy.port_connections: dev2_copy.return_port(port2)

        # 尝试 MPO-SFP (UHD/HorizoN 与 MicroN 之间)
        uhd_dev, micron_dev = (None, None)
        original_dev1_id = dev1_copy.id # 记录原始顺序
        if is_uhd1 and is_mn2:
            uhd_dev, micron_dev = dev1_copy, dev2_copy
        elif is_uhd2 and is_mn1:
            uhd_dev, micron_dev = dev2_copy, dev1_copy

        if uhd_dev and micron_dev:
            port_uhd = uhd_dev.get_specific_available_port(PORT_MPO)
            if port_uhd:
                port_micron = micron_dev.get_specific_available_port(PORT_SFP)
                if port_micron:
                    if uhd_dev.use_specific_port(port_uhd, micron_dev.name) and micron_dev.use_specific_port(port_micron, uhd_dev.name):
                        # 根据原始顺序返回端口
                        if uhd_dev.id == original_dev1_id:
                            return port_uhd, port_micron, f"{PORT_MPO}-{PORT_SFP} (10G)"
                        else:
                            return port_micron, port_uhd, f"{PORT_MPO}-{PORT_SFP} (10G)" # 注意顺序反了
                    else:
                        print(f"警告: _find_best_single_link 中 MPO-SFP 端口占用失败 {uhd_dev.name}[{port_uhd}] 或 {micron_dev.name}[{port_micron}]")
                        if port_uhd in uhd_dev.port_connections: uhd_dev.return_port(port_uhd)
                        if port_micron in micron_dev.port_connections: micron_dev.return_port(port_micron)

        # 没有找到任何兼容的可用连接
        return None, None, None

    def calculate_mesh(self) -> List[ConnectionType]:
        """
        计算 Mesh 连接方案。
        此方法不修改 NetworkManager 的状态，仅返回计算出的连接列表。

        Returns:
            List[ConnectionType]: 计算出的 Mesh 连接元组列表。
        """
        if len(self.devices) < 2:
            return []

        calculated_connections: List[ConnectionType] = []
        # 使用设备的深拷贝进行计算，避免修改原始状态
        temp_devices = [copy.deepcopy(dev) for dev in self.devices]
        for d in temp_devices: d.reset_ports() # 确保副本状态干净
        device_map = {dev.id: dev for dev in temp_devices}

        all_pairs_ids = list(itertools.combinations([d.id for d in temp_devices], 2))
        connected_once_pairs: Set[Tuple[int, int]] = set()

        print("Mesh Phase 1: 尝试为每个设备对建立第一条连接...")
        made_progress_phase1 = True
        failed_pairs_phase1 = []
        while made_progress_phase1:
            made_progress_phase1 = False
            for dev1_id, dev2_id in all_pairs_ids:
                pair_key = tuple(sorted((dev1_id, dev2_id)))
                if pair_key not in connected_once_pairs:
                    dev1_copy = device_map[dev1_id]
                    dev2_copy = device_map[dev2_id]
                    # 注意：_find_best_single_link 会修改 dev1_copy 和 dev2_copy 的状态
                    port1, port2, conn_type = self._find_best_single_link(dev1_copy, dev2_copy)
                    if port1 and port2:
                        # 获取原始设备对象用于添加到结果列表
                        original_dev1 = self.get_device_by_id(dev1_id)
                        original_dev2 = self.get_device_by_id(dev2_id)
                        if original_dev1 and original_dev2:
                             # 确保返回的元组中设备顺序与 _find_best_single_link 内部逻辑一致
                             if dev1_copy.id == dev1_id: # 检查副本 ID 是否与原始 ID 匹配
                                 calculated_connections.append((original_dev1, port1, original_dev2, port2, conn_type))
                             else: # 如果副本顺序反了，结果也要反
                                 calculated_connections.append((original_dev2, port2, original_dev1, port1, conn_type))
                             connected_once_pairs.add(pair_key)
                             made_progress_phase1 = True
                        else:
                             print(f"严重错误: Mesh Phase 1 中找不到原始设备对象 ID {dev1_id} 或 {dev2_id}")
                    else:
                        # 记录第一次失败的对
                        if pair_key not in [fp[0] for fp in failed_pairs_phase1]:
                            original_dev1 = self.get_device_by_id(dev1_id)
                            original_dev2 = self.get_device_by_id(dev2_id)
                            if original_dev1 and original_dev2:
                                failed_pairs_phase1.append((pair_key, f"{original_dev1.name} <-> {original_dev2.name}"))
                            else:
                                failed_pairs_phase1.append((pair_key, f"ID {dev1_id} <-> ID {dev2_id}"))

            if not made_progress_phase1:
                break # 如果一轮没有任何进展，则退出

        if len(connected_once_pairs) < len(all_pairs_ids):
            print(f"警告: Mesh Phase 1 未能为所有设备对建立连接。失败 {len(all_pairs_ids) - len(connected_once_pairs)} 对。")
            # 可以选择性地报告 failed_pairs_phase1

        print(f"Mesh Phase 1 完成. 建立了 {len(calculated_connections)} 条初始连接。")
        print("Mesh Phase 2: 填充剩余端口...")
        connection_made_in_full_pass_phase2 = True
        while connection_made_in_full_pass_phase2:
            connection_made_in_full_pass_phase2 = False
            random.shuffle(all_pairs_ids) # 随机化顺序以尝试不同组合
            for dev1_id, dev2_id in all_pairs_ids:
                dev1_copy = device_map[dev1_id]
                dev2_copy = device_map[dev2_id]
                port1, port2, conn_type = self._find_best_single_link(dev1_copy, dev2_copy)
                if port1 and port2:
                    original_dev1 = self.get_device_by_id(dev1_id)
                    original_dev2 = self.get_device_by_id(dev2_id)
                    if original_dev1 and original_dev2:
                        if dev1_copy.id == dev1_id:
                             calculated_connections.append((original_dev1, port1, original_dev2, port2, conn_type))
                        else:
                             calculated_connections.append((original_dev2, port2, original_dev1, port1, conn_type))
                        connection_made_in_full_pass_phase2 = True
                    else:
                         print(f"严重错误: Mesh Phase 2 中找不到原始设备对象 ID {dev1_id} 或 {dev2_id}")

        print(f"Mesh Phase 2 完成. 总计算连接数: {len(calculated_connections)}")
        return calculated_connections

    def calculate_ring(self) -> Tuple[List[ConnectionType], Optional[str]]:
        """
        计算 Ring 连接方案。
        此方法不修改 NetworkManager 的状态，仅返回计算出的连接列表和错误信息。

        Returns:
            Tuple[List[ConnectionType], Optional[str]]:
                (计算出的 Ring 连接元组列表, 如果无法形成完整环则返回错误信息字符串，否则为 None)。
        """
        if len(self.devices) < 2:
            return [], "设备数量少于 2，无法形成环形"
        if len(self.devices) == 2:
            # 两个设备的情况退化为 Mesh 计算（通常是 1 或 2 条直连）
            print("设备数量为 2，使用 Mesh 逻辑计算连接。")
            mesh_conns = self.calculate_mesh()
            return mesh_conns, None

        calculated_connections: List[ConnectionType] = []
        temp_devices = [copy.deepcopy(dev) for dev in self.devices]
        for d in temp_devices: d.reset_ports()
        device_map = {dev.id: dev for dev in temp_devices}

        # 按 ID 排序以确保环连接顺序一致
        sorted_dev_ids = sorted([d.id for d in self.devices])
        num_devices = len(sorted_dev_ids)
        link_established = [True] * num_devices
        failed_segments = []

        for i in range(num_devices):
            dev1_id = sorted_dev_ids[i]
            dev2_id = sorted_dev_ids[(i + 1) % num_devices] # 连接到下一个，最后一个连回第一个
            dev1_copy = device_map[dev1_id]
            dev2_copy = device_map[dev2_id]

            port1, port2, conn_type = self._find_best_single_link(dev1_copy, dev2_copy)

            if port1 and port2:
                original_dev1 = self.get_device_by_id(dev1_id)
                original_dev2 = self.get_device_by_id(dev2_id)
                if original_dev1 and original_dev2:
                     if dev1_copy.id == dev1_id:
                         calculated_connections.append((original_dev1, port1, original_dev2, port2, conn_type))
                     else:
                         calculated_connections.append((original_dev2, port2, original_dev1, port1, conn_type))
                else:
                     print(f"严重错误: Ring 计算中找不到原始设备对象 ID {dev1_id} 或 {dev2_id}")
                     link_established[i] = False # 标记此段失败
                     failed_segments.append(f"ID {dev1_id} <-> ID {dev2_id} (对象丢失)")
            else:
                link_established[i] = False
                original_dev1 = self.get_device_by_id(dev1_id)
                original_dev2 = self.get_device_by_id(dev2_id)
                segment_name = f"ID {dev1_id} <-> ID {dev2_id}"
                if original_dev1 and original_dev2:
                    segment_name = f"{original_dev1.name} <-> {original_dev2.name}"
                failed_segments.append(segment_name)
                print(f"警告: 无法在 {segment_name} 之间建立环形连接段。")

        error_message = None
        if not all(link_established):
            error_message = f"未能完成完整的环形连接。无法连接的段落：{', '.join(failed_segments)}。"
            print(f"警告: {error_message}")

        print(f"Ring 计算完成. 计算连接数: {len(calculated_connections)}")
        return calculated_connections, error_message

    def _fill_connections_style(self, style='mesh') -> List[ConnectionType]:
        """
        内部辅助函数：使用指定风格（Mesh 或 Ring）填充当前状态下的剩余端口。
        这个方法 *会* 修改 NetworkManager 的状态 (self.connections 和 device 状态)。

        Args:
            style (str): 填充风格，'mesh' 或 'ring'。

        Returns:
            List[ConnectionType]: 新添加的连接列表。
        """
        if len(self.devices) < 2:
            return []

        newly_added_connections: List[ConnectionType] = []
        # 注意：这里直接在 self.devices 上操作，因为是要修改当前状态
        # 但计算逻辑内部仍然需要探测，所以 _find_best_single_link 需要副本
        # 更好的方式是让 _find_best_single_link 不修改状态，这里循环查找并添加

        all_pairs_ids = list(itertools.combinations([d.id for d in self.devices], 2))
        sorted_dev_ids = sorted([d.id for d in self.devices])
        num_devices = len(sorted_dev_ids)

        connection_made_in_full_pass = True
        while connection_made_in_full_pass:
            connection_made_in_full_pass = False
            # 根据风格选择迭代顺序
            iterator: Any = all_pairs_ids if style == 'mesh' else range(num_devices)
            if style == 'mesh': random.shuffle(iterator) # Mesh 随机化

            for item in iterator:
                if style == 'mesh':
                    dev1_id, dev2_id = item
                else: # style == 'ring'
                    i = item
                    dev1_id = sorted_dev_ids[i]
                    dev2_id = sorted_dev_ids[(i + 1) % num_devices]

                dev1 = self.get_device_by_id(dev1_id)
                dev2 = self.get_device_by_id(dev2_id)

                if not dev1 or not dev2 or dev1_id == dev2_id: continue # 设备不存在或相同

                # 使用副本进行探测，避免修改真实状态
                dev1_copy = copy.deepcopy(dev1)
                dev2_copy = copy.deepcopy(dev2)
                port1, port2, conn_type = self._find_best_single_link(dev1_copy, dev2_copy)

                if port1 and port2:
                    # 找到可用连接，现在尝试在真实对象上添加
                    # 需要确保端口在真实对象上仍然可用（理论上应该可用，因为我们没修改真实状态）
                    actual_port1_name = port1 if dev1_copy.id == dev1_id else port2
                    actual_port2_name = port2 if dev1_copy.id == dev1_id else port1
                    added_connection = self.add_connection(dev1_id, actual_port1_name, dev2_id, actual_port2_name)
                    if added_connection:
                        newly_added_connections.append(added_connection)
                        connection_made_in_full_pass = True # 成功添加，可能还有更多

        print(f"填充完成 ({style} 风格). 新增连接数: {len(newly_added_connections)}")
        return newly_added_connections

    def fill_connections_mesh(self) -> List[ConnectionType]:
        """使用 Mesh 风格填充剩余端口。"""
        return self._fill_connections_style(style='mesh')

    def fill_connections_ring(self) -> List[ConnectionType]:
        """使用 Ring 风格填充剩余端口。"""
        return self._fill_connections_style(style='ring')


    # --- 验证与辅助 ---

    def check_port_compatibility(self, dev1_id: int, port1_name: str, dev2_id: int, port2_name: str) -> Tuple[bool, Optional[str]]:
        """
        检查两个指定端口之间是否兼容。

        Args:
            dev1_id (int): 设备 1 的 ID。
            port1_name (str): 设备 1 的端口名称。
            dev2_id (int): 设备 2 的 ID。
            port2_name (str): 设备 2 的端口名称。

        Returns:
            Tuple[bool, Optional[str]]: (是否兼容, 连接类型描述字符串 或 None)。
        """
        dev1 = self.get_device_by_id(dev1_id)
        dev2 = self.get_device_by_id(dev2_id)
        if not dev1 or not dev2:
            return False, None

        port1_type = get_port_type_from_name(port1_name)
        port2_type = get_port_type_from_name(port2_name)
        is_uhd1 = dev1.type in UHD_TYPES
        is_uhd2 = dev2.type in UHD_TYPES
        is_mn1 = dev1.type == DEV_MN
        is_mn2 = dev2.type == DEV_MN

        # 规则 1: LC 只能 UHD/HorizoN 之间互连
        if port1_type == PORT_LC and port2_type == PORT_LC and is_uhd1 and is_uhd2:
            return True, f"{PORT_LC}-{PORT_LC} (100G)"
        # 规则 2: MPO 只能 UHD/HorizoN 之间互连
        if port1_type == PORT_MPO and port2_type == PORT_MPO and is_uhd1 and is_uhd2:
            return True, f"{PORT_MPO}-{PORT_MPO} (25G)"
        # 规则 3: SFP 只能 MicroN 之间互连
        if port1_type == PORT_SFP and port2_type == PORT_SFP and is_mn1 and is_mn2:
            return True, f"{PORT_SFP}-{PORT_SFP} (10G)"
        # 规则 4: MPO (UHD/HorizoN) 可以连接 SFP (MicroN)
        if port1_type == PORT_MPO and port2_type == PORT_SFP and is_uhd1 and is_mn2:
            return True, f"{PORT_MPO}-{PORT_SFP} (10G)"
        if port1_type == PORT_SFP and port2_type == PORT_MPO and is_mn1 and is_uhd2:
            return True, f"{PORT_MPO}-{PORT_SFP} (10G)" # 描述统一

        # 其他组合均不兼容
        return False, None

    def get_compatible_port_types(self, target_dev_id: int, target_port_name: str) -> List[str]:
        """
        根据目标设备和端口，获取当前设备上兼容的端口类型列表。
        用于 UI 过滤。

        Args:
            target_dev_id (int): 目标设备的 ID。
            target_port_name (str): 目标设备的端口名称。

        Returns:
            List[str]: 本地设备上兼容的端口类型常量列表 (例如 [PORT_MPO, PORT_SFP])。
        """
        target_device = self.get_device_by_id(target_dev_id)
        if not target_device:
            return []

        target_port_type = get_port_type_from_name(target_port_name)
        compatible_here = []
        is_target_uhd = target_device.type in UHD_TYPES
        is_target_mn = target_device.type == DEV_MN

        if is_target_uhd:
            if target_port_type == PORT_LC: compatible_here = [PORT_LC] # LC 只能连 LC
            elif target_port_type == PORT_MPO: compatible_here = [PORT_MPO, PORT_SFP] # MPO 可连 MPO 或 SFP
        elif is_target_mn:
            if target_port_type == PORT_SFP: compatible_here = [PORT_MPO, PORT_SFP] # SFP 可连 MPO 或 SFP

        return compatible_here

    def get_available_ports(self, device_id: int) -> List[str]:
        """获取指定设备的所有可用端口列表。"""
        device = self.get_device_by_id(device_id)
        if device:
            return device.get_all_available_ports()
        return []

    def calculate_port_totals(self) -> Dict[str, int]:
        """计算当前所有设备各类端口的总数。"""
        total_mpo = sum(dev.mpo_total for dev in self.devices)
        total_lc = sum(dev.lc_total for dev in self.devices)
        total_sfp = sum(dev.sfp_total for dev in self.devices)
        return {'mpo': total_mpo, 'lc': total_lc, 'sfp': total_sfp}

    # --- 图管理 ---

    def _update_graph(self):
        """根据当前的设备和连接列表更新 NetworkX 图对象。"""
        self.graph.clear() # 清空旧图

        # 添加所有设备作为节点
        for dev in self.devices:
            self.graph.add_node(dev.id, label=f"{dev.name}\n({dev.type})", device_type=dev.type) # 添加属性

        # 添加所有连接作为边
        edge_data = defaultdict(lambda: defaultdict(int)) # {(u,v): {conn_type: count}}
        edge_ports_info: Dict[Tuple[int, int], List[Dict[str, Any]]] = defaultdict(list) # {(u,v): [{'source': p1, 'target': p2, 'type': t}, ...]}


        for conn in self.connections:
            dev1, port1, dev2, port2, conn_type = conn
            # 确保节点存在于图中 (理论上应该存在)
            if self.graph.has_node(dev1.id) and self.graph.has_node(dev2.id):
                 # 使用排序后的 ID 作为键，确保边的唯一性
                 u, v = tuple(sorted((dev1.id, dev2.id)))
                 # 聚合相同类型连接的数量
                 base_conn_type = conn_type.split(' ')[0] # 例如 "LC-LC"
                 edge_data[(u, v)][base_conn_type] += 1

                 # 记录端口信息
                 if u == dev1.id:
                     edge_ports_info[(u,v)].append({'source': port1, 'target': port2, 'type': conn_type})
                 else: # 确保 source/target 对应 u/v
                     edge_ports_info[(u,v)].append({'source': port2, 'target': port1, 'type': conn_type})
            else:
                 print(f"警告: 更新图时发现连接涉及未知节点: {dev1.name} 或 {dev2.name}")

        # 添加边和属性到图中
        for (u, v), type_counts in edge_data.items():
            label_parts = []
            ports_list = edge_ports_info.get((u,v), [])
            for base_type, count in type_counts.items():
                # 从端口信息中获取该类型的完整描述
                full_type_desc = "Unknown Type"
                first_match = next((p['type'] for p in ports_list if p['type'].startswith(base_type)), None)
                if first_match:
                    full_type_desc = first_match
                label_parts.append(f"{full_type_desc} x{count}")

            # 确定边的颜色
            first_base_type = next(iter(type_counts))
            color = 'black'
            if 'LC-LC' in first_base_type: color = 'blue'
            elif 'MPO-MPO' in first_base_type: color = 'red'
            elif 'MPO-SFP' in first_base_type: color = 'orange'
            elif 'SFP-SFP' in first_base_type: color = 'purple'

            # 添加边及所有属性
            self.graph.add_edge(
                u, v,
                label="\n".join(label_parts),
                count=sum(type_counts.values()), # 总连接数
                color=color,
                ports=ports_list # 存储详细端口信息
            )

        print(f"图已更新: {self.graph.number_of_nodes()} 个节点, {self.graph.number_of_edges()} 条边")


    def get_graph(self) -> nx.Graph:
        """获取当前的 NetworkX 图对象。"""
        # 确保图是最新的
        # 在每次获取前都更新可能效率不高，取决于调用频率
        # 更好的方式是在修改 devices 或 connections 后调用 _update_graph
        # self._update_graph() # 暂时注释掉，依赖于修改方法调用它
        return self.graph

    # --- 保存与加载 ---

    def save_project(self, filepath: str) -> bool:
        """
        将当前设备列表和连接列表保存到 JSON 文件。

        Args:
            filepath (str): 保存文件的完整路径。

        Returns:
            bool: 保存是否成功。
        """
        try:
            project_data = {
                'version': '1.1-refactored', # 添加版本标记
                'devices': [dev.to_dict() for dev in self.devices],
                'connections': [
                    {
                        'dev1_id': conn[0].id,
                        'port1': conn[1],
                        'dev2_id': conn[2].id,
                        'port2': conn[3],
                        'type': conn[4]
                    }
                    for conn in self.connections
                ]
                # TODO: 未来可以保存 node_positions 等视图状态
            }
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(project_data, f, indent=4, ensure_ascii=False)
            print(f"项目已保存到: {filepath}")
            return True
        except Exception as e:
            print(f"错误: 保存项目失败 - {e}")
            return False

    def load_project(self, filepath: str) -> bool:
        """
        从 JSON 文件加载设备列表和连接列表，覆盖当前状态。
        **增加了对旧格式（仅设备列表）的兼容性处理。**

        Args:
            filepath (str): 加载文件的完整路径。

        Returns:
            bool: 加载是否成功。
        """
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                project_data = json.load(f)

            # 1. 清空当前状态
            self.clear_all_devices_and_connections()
            print(f"尝试从 {filepath} 加载项目...") # 移到清空之后打印

            loaded_devices_data = []
            loaded_connections_data = []

            # **修复: 检查加载的数据格式**
            if isinstance(project_data, dict):
                # 新格式：包含 'devices' 和 'connections' 键的字典
                loaded_devices_data = project_data.get('devices', [])
                loaded_connections_data = project_data.get('connections', [])
                print("检测到新格式配置文件 (包含设备和连接)。")
            elif isinstance(project_data, list):
                # 旧格式：只包含设备列表
                loaded_devices_data = project_data
                loaded_connections_data = [] # 旧格式无连接信息
                print("警告: 检测到旧格式配置文件 (仅包含设备)，连接信息将不会被加载。")
                # 可以在这里通过 parent_window 显示一个更明显的警告
                # if parent_window: # 需要将 parent_window 传递进来
                #     QMessageBox.warning(parent_window, "旧格式文件", "加载的文件是旧格式，仅设备信息被加载，连接信息丢失。")
            else:
                # 无效格式
                raise TypeError("无法识别的项目文件格式 (既不是列表也不是字典)。")


            # 2. 加载设备
            max_id = 0
            temp_device_map: Dict[int, Device] = {} # 用于查找设备对象
            for data in loaded_devices_data:
                try:
                    # 确保 data 是字典
                    if not isinstance(data, dict):
                         print(f"警告: 跳过无效的设备条目 (非字典): {data}")
                         continue
                    new_device = Device.from_dict(data)
                    self.devices.append(new_device)
                    temp_device_map[new_device.id] = new_device
                    if new_device.id > max_id:
                        max_id = new_device.id
                except Exception as e:
                    print(f"警告: 加载设备数据时出错: {data} - {e}")
            self.device_id_counter = max_id # 更新 ID 计数器

            # 3. 加载并重建连接状态 (仅对新格式有效)
            rebuilt_connections: List[ConnectionType] = []
            if loaded_connections_data: # 只有新格式才有连接数据
                print(f"正在加载 {len(loaded_connections_data)} 条连接...")
                for conn_data in loaded_connections_data:
                    # 确保 conn_data 是字典
                    if not isinstance(conn_data, dict):
                        print(f"警告: 跳过无效的连接条目 (非字典): {conn_data}")
                        continue

                    dev1_id = conn_data.get('dev1_id')
                    port1 = conn_data.get('port1')
                    dev2_id = conn_data.get('dev2_id')
                    port2 = conn_data.get('port2')
                    conn_type = conn_data.get('type', 'Unknown Type') # 提供默认值

                    dev1 = temp_device_map.get(dev1_id)
                    dev2 = temp_device_map.get(dev2_id)

                    if dev1 and dev2 and port1 and port2:
                        # 尝试在设备上标记端口占用
                        # 注意：这里不进行兼容性检查，假设保存的文件是有效的
                        if dev1.use_specific_port(port1, dev2.name):
                            if dev2.use_specific_port(port2, dev1.name):
                                rebuilt_connections.append((dev1, port1, dev2, port2, conn_type))
                            else:
                                print(f"警告: 加载连接时，设备 {dev2.name} 端口 {port2} 占用失败，已回滚 {dev1.name} 端口 {port1}。")
                                dev1.return_port(port1) # 回滚
                        else:
                             print(f"警告: 加载连接时，设备 {dev1.name} 端口 {port1} 占用失败。")
                    else:
                        print(f"警告: 加载连接数据时跳过无效条目: {conn_data} (设备 ID {dev1_id} 或 {dev2_id} 未找到，或端口信息缺失)")

            self.connections = rebuilt_connections

            # 4. 更新图
            self._update_graph()

            print(f"项目已从 {filepath} 加载。 设备数: {len(self.devices)}, 连接数: {len(self.connections)}")
            return True

        except FileNotFoundError:
            print(f"错误: 找不到项目文件: {filepath}")
            # 清空状态以防部分加载
            self.clear_all_devices_and_connections()
            return False
        except json.JSONDecodeError:
            print(f"错误: 项目文件格式错误 (无法解析 JSON): {filepath}")
            self.clear_all_devices_and_connections()
            return False
        except TypeError as e: # 捕获我们自己抛出的 TypeError
             print(f"错误: 加载项目失败 - {e}")
             self.clear_all_devices_and_connections()
             return False
        except Exception as e:
            print(f"错误: 加载项目时发生未知错误 - {e}")
            # 确保状态清空
            self.clear_all_devices_and_connections()
            return False
