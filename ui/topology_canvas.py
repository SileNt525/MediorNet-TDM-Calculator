# -*- coding: utf-8 -*-
"""
ui/topology_canvas.py

定义 MplCanvas 类，一个用于显示 Matplotlib 图形的 Qt Widget。
负责绘制网络拓扑图。
"""
import sys
import os
import copy
from typing import List, Dict, Tuple, Optional, Any

# Matplotlib imports
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.lines import Line2D # 导入 Line2D 用于图例

# NetworkX import
import networkx as nx

# PySide6 imports
from PySide6.QtWidgets import QSizePolicy, QWidget # 导入 QWidget 以便类型提示
from PySide6.QtGui import QFontDatabase, QFont # 导入 QFontDatabase, QFont

# 从项目模块导入
try:
    # 假设 utils 和 core 在上一级目录
    from core.device import (
        Device,
        DEV_UHD, DEV_HORIZON, DEV_MN, UHD_TYPES,
        PORT_MPO, PORT_LC, PORT_SFP # 导入常量
    )
    from utils.misc_utils import resource_path # 导入资源路径函数
except ImportError as e:
     print(f"导入错误 (topology_canvas.py): {e} - 请确保 core 和 utils 包已正确创建。")
     # Fallbacks
     Device = object
     DEV_UHD, DEV_HORIZON, DEV_MN, UHD_TYPES = '', '', '', []
     PORT_MPO, PORT_LC, PORT_SFP = '', '', ''
     resource_path = lambda x: x


class MplCanvas(FigureCanvas):
    """
    用于嵌入 Matplotlib 图形的自定义 Qt Widget。
    专门用于绘制 MediorNet 网络拓扑图。
    """
    def __init__(self, parent: Optional[QWidget] = None, width: int = 5, height: int = 4, dpi: int = 100):
        """
        初始化 Matplotlib 画布。

        Args:
            parent (Optional[QWidget]): 父 Widget。
            width (int): 画布宽度（英寸）。
            height (int): 画布高度（英寸）。
            dpi (int): 每英寸点数。
        """
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super().__init__(self.fig)
        self.setParent(parent)

        # 设置尺寸策略，使其可以缩放
        FigureCanvas.setSizePolicy(self, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        FigureCanvas.updateGeometry(self)

        # 初始化字体属性
        self.chinese_font_prop = self._get_matplotlib_font_prop()
        self.current_font_family = self.chinese_font_prop.get_name() if self.chinese_font_prop else 'sans-serif'


    def _get_matplotlib_font_prop(self) -> Optional[font_manager.FontProperties]:
         """获取用于 Matplotlib 的 FontProperties 对象"""
         try:
             font_relative_path = os.path.join('assets', 'NotoSansCJKsc-Regular.otf')
             font_path = resource_path(font_relative_path)
             if os.path.exists(font_path):
                 font_manager.fontManager.addfont(font_path)
                 # 检查字体是否真的被 Matplotlib 识别
                 font_families = [f.name for f in font_manager.fontManager.ttflist]
                 target_family = 'Noto Sans CJK SC' # 字体文件名可能与 family name 不同
                 # 查找可能的 family name
                 prop = font_manager.FontProperties(fname=font_path)
                 actual_family = prop.get_name()

                 if actual_family in font_families:
                     print(f"Matplotlib 字体设置成功: {actual_family}")
                     plt.rcParams['font.sans-serif'] = [actual_family] + plt.rcParams.get('font.sans-serif', [])
                     plt.rcParams['axes.unicode_minus'] = False
                     return font_manager.FontProperties(family=actual_family)
                 else:
                      print(f"警告: 字体文件 {font_path} 已添加，但在 Matplotlib 中未找到 family '{actual_family}' 或 'Noto Sans CJK SC'。")

         except Exception as e:
              print(f"获取 Matplotlib 字体属性时出错: {e}")

         # 回退
         print("警告: 未能加载或设置中文字体，回退到默认 sans-serif。")
         plt.rcParams['font.sans-serif'] = ['sans-serif']
         plt.rcParams['axes.unicode_minus'] = False
         return font_manager.FontProperties(family='sans-serif')


    def plot_topology(self,
                      devices: List[Device],
                      connections: List[Tuple[Device, str, Device, str, str]],
                      layout_algorithm: str = 'spring',
                      fixed_pos: Optional[Dict[int, Tuple[float, float]]] = None,
                      selected_node_id: Optional[int] = None,
                      port_totals_dict: Optional[Dict[str, int]] = None
                      ) -> Tuple[Optional[Figure], Optional[Dict[int, Tuple[float, float]]]]:
        """
        在画布上绘制网络拓扑图。

        Args:
            devices (List[Device]): 要绘制的设备列表。
            connections (List[Tuple[Device, str, Device, str, str]]): 连接列表。
            layout_algorithm (str): 使用的布局算法 ('spring', 'circular', etc.)。
            fixed_pos (Optional[Dict[int, Tuple[float, float]]]): 预设的节点位置。如果提供且节点未变，则使用此布局。
            selected_node_id (Optional[int]): 要高亮显示的节点 ID。
            port_totals_dict (Optional[Dict[str, int]]): 包含端口总数的字典，用于显示。

        Returns:
            Tuple[Optional[Figure], Optional[Dict[int, Tuple[float, float]]]]:
                (绘制的 Figure 对象, 计算出的节点位置字典)
        """
        self.axes.cla() # 清除之前的绘图

        if not devices:
            self.axes.text(0.5, 0.5, '无设备数据', ha='center', va='center', fontproperties=self.chinese_font_prop)
            self.draw()
            return self.fig, None # 返回 Figure 但无位置信息

        # --- 构建 NetworkX 图 ---
        G = nx.Graph()
        node_ids = [dev.id for dev in devices]
        node_colors = []
        node_labels = {}
        node_alphas = []
        highlight_color = 'yellow'
        default_alpha = 0.9
        dimmed_alpha = 0.3

        # 添加节点并设置标签和基础颜色/透明度
        for dev in devices:
            G.add_node(dev.id)
            node_labels[dev.id] = f"{dev.name}\n({dev.type})"
            base_color = 'grey'
            if dev.type == DEV_UHD: base_color = 'skyblue'
            elif dev.type == DEV_HORIZON: base_color = 'lightcoral'
            elif dev.type == DEV_MN: base_color = 'lightgreen'

            # 根据是否有选中节点来设置颜色和透明度
            if selected_node_id is not None:
                if dev.id == selected_node_id:
                    node_colors.append(highlight_color)
                    node_alphas.append(default_alpha)
                else:
                    # 检查是否为邻居
                    is_neighbor = any(
                        (conn[0].id == selected_node_id and conn[2].id == dev.id) or \
                        (conn[0].id == dev.id and conn[2].id == selected_node_id)
                        for conn in connections
                    )
                    node_colors.append(base_color)
                    node_alphas.append(default_alpha if is_neighbor else dimmed_alpha)
            else:
                # 没有选中节点，所有节点都正常显示
                node_colors.append(base_color)
                node_alphas.append(default_alpha)

        # 添加边并聚合标签
        edge_labels: Dict[Tuple[int, int], str] = {}
        edge_counts: Dict[Tuple[int, int], Dict[str, Dict[str, Any]]] = {} # {(u,v): {base_type: {'count': n, 'details': full_desc}}}
        highlighted_edges: Set[Tuple[int, int]] = set()

        if connections:
            for conn in connections:
                dev1, _, dev2, _, conn_type = conn
                if dev1.id in node_ids and dev2.id in node_ids:
                    edge_key = tuple(sorted((dev1.id, dev2.id))) # 确保边的键顺序一致
                    G.add_edge(dev1.id, dev2.id)

                    # 聚合连接类型和数量
                    if edge_key not in edge_counts:
                        edge_counts[edge_key] = {}
                    base_conn_type = conn_type.split(' ')[0] # 例如 "LC-LC"
                    if base_conn_type not in edge_counts[edge_key]:
                        edge_counts[edge_key][base_conn_type] = {'count': 0, 'details': conn_type} # 存储第一个遇到的完整描述
                    edge_counts[edge_key][base_conn_type]['count'] += 1

                    # 如果有选中节点，标记高亮的边
                    if selected_node_id is not None and (dev1.id == selected_node_id or dev2.id == selected_node_id):
                        highlighted_edges.add(edge_key)

            # 生成最终的边标签
            for edge_key, type_groups in edge_counts.items():
                label_parts = [f"{data['details']} x{data['count']}" for base_type, data in type_groups.items()]
                edge_labels[edge_key] = "\n".join(label_parts)

        # --- 计算布局 ---
        pos = None
        if not G:
            print("DIAG (Plot): 图为空，不计算布局。")
            self.axes.text(0.5, 0.5, '无连接数据', ha='center', va='center', fontproperties=self.chinese_font_prop)
            self.draw()
            return self.fig, None

        # 检查是否可以使用固定布局
        if fixed_pos:
            current_node_ids = set(G.nodes())
            stored_node_ids = set(fixed_pos.keys())
            # 只有当节点完全相同时才使用固定布局
            if current_node_ids == stored_node_ids:
                pos = fixed_pos
            else:
                print("DIAG (Plot): 节点已更改，重新计算布局。")
                fixed_pos = None # 强制重新计算

        # 如果没有固定布局或节点已变，则计算新布局
        if pos is None:
            try:
                if layout_algorithm == 'circular':
                    pos = nx.circular_layout(G)
                elif layout_algorithm == 'kamada-kawai':
                    pos = nx.kamada_kawai_layout(G)
                elif layout_algorithm == 'random':
                    pos = nx.random_layout(G, seed=42) # 使用种子保证随机布局可复现
                elif layout_algorithm == 'shell':
                    # 按设备类型分层
                    shells = []
                    types_present = sorted(list(set(dev.type for dev in devices)))
                    for t in types_present:
                        shells.append([dev.id for dev in devices if dev.type == t])
                    # 如果只有一层或没有设备，shell 布局可能效果不好，回退到 spring
                    if len(shells) < 2:
                         print("DIAG (Plot): Shell 布局层数不足，使用 Spring 布局。")
                         pos = nx.spring_layout(G, seed=42, k=0.8) # k 值调整节点间距
                    else:
                         pos = nx.shell_layout(G, nlist=shells)
                else: # 默认为 spring
                    pos = nx.spring_layout(G, seed=42, k=0.8)
            except Exception as e:
                print(f"警告: 计算布局 '{layout_algorithm}' 时出错: {e}. 使用 spring 布局回退。")
                pos = nx.spring_layout(G, seed=42, k=0.8)

        # --- 绘制图形 ---
        # 绘制节点
        nx.draw_networkx_nodes(G, pos, node_color=node_colors, node_size=3500, ax=self.axes, alpha=node_alphas)
        # 绘制节点标签
        nx.draw_networkx_labels(G, pos, labels=node_labels, font_size=9, ax=self.axes, font_family=self.current_font_family)

        # 绘制边和边标签
        if connections and G.edges():
            unique_edges = list(G.edges())
            edge_colors = []
            edge_widths = []
            edge_alphas = []
            highlight_edge_width = 2.5
            default_edge_width = 1.5
            dimmed_edge_alpha = 0.15
            default_edge_alpha = 0.7

            # 设置边的颜色、宽度和透明度
            for u, v in unique_edges:
                edge_key = tuple(sorted((u, v)))
                is_highlighted = edge_key in highlighted_edges
                is_selected_node_present = selected_node_id is not None

                # 确定边颜色
                color_found = 'black' # 默认颜色
                if edge_key in edge_counts:
                    # 基于第一个连接类型确定颜色（可能需要更复杂的逻辑如果混合类型）
                    first_base_type = next(iter(edge_counts[edge_key]))
                    if 'LC-LC' in first_base_type: color_found = 'blue'
                    elif 'MPO-MPO' in first_base_type: color_found = 'red'
                    elif 'MPO-SFP' in first_base_type: color_found = 'orange'
                    elif 'SFP-SFP' in first_base_type: color_found = 'purple'
                edge_colors.append(color_found)

                # 确定边宽度
                edge_widths.append(highlight_edge_width if is_highlighted else default_edge_width)

                # 确定边透明度
                edge_alphas.append(default_edge_alpha if (not is_selected_node_present or is_highlighted) else dimmed_edge_alpha)

            # 绘制边
            nx.draw_networkx_edges(G, pos, edgelist=unique_edges, edge_color=edge_colors,
                                   width=edge_widths, alpha=edge_alphas, ax=self.axes, arrows=False)

            # 绘制边标签 (需要处理高亮时的颜色)
            edge_label_colors = {}
            dimmed_label_color = 'lightgrey'
            default_label_color = 'black'
            for edge, label in edge_labels.items():
                 edge_key = tuple(sorted(edge))
                 is_highlighted = edge_key in highlighted_edges
                 is_selected_node_present = selected_node_id is not None
                 # 只有在有节点选中且当前边未高亮时才应用暗淡颜色
                 color = dimmed_label_color if (is_selected_node_present and not is_highlighted) else default_label_color
                 edge_label_colors[edge_key] = color # 存储每个标签的颜色

            # 使用 font_color 参数一次性设置所有标签颜色可能不支持字典，需要单独绘制或修改 NetworkX 源码
            # 暂时先用默认颜色绘制所有标签
            nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, font_size=7, ax=self.axes,
                                         label_pos=0.5, rotate=False, font_family=self.current_font_family,
                                         font_color=default_label_color) # TODO: 实现标签颜色区分高亮

        # --- 显示端口总数 ---
        if port_totals_dict is not None:
            totals_text = f"端口总计: {PORT_MPO}: {port_totals_dict['mpo']}, {PORT_LC}: {port_totals_dict['lc']}, {PORT_SFP}+: {port_totals_dict['sfp']}"
            # 在 Figure 坐标系左下角添加文本
            self.fig.text(0.01, 0.01, totals_text, ha='left', va='bottom', fontsize=7, color='grey', transform=self.fig.transFigure)


        # --- 设置标题和图例 ---
        self.axes.set_title("网络连接拓扑图", fontproperties=self.chinese_font_prop)
        self.axes.axis('off') # 关闭坐标轴

        # 创建图例元素
        legend_elements = [
            Line2D([0], [0], marker='o', color='w', label=DEV_UHD, markerfacecolor='skyblue', markersize=10),
            Line2D([0], [0], marker='o', color='w', label=DEV_HORIZON, markerfacecolor='lightcoral', markersize=10),
            Line2D([0], [0], marker='o', color='w', label=DEV_MN, markerfacecolor='lightgreen', markersize=10),
            Line2D([0], [0], color='blue', lw=2, label=f'{PORT_LC}-{PORT_LC} (100G)'),
            Line2D([0], [0], color='red', lw=2, label=f'{PORT_MPO}-{PORT_MPO} (25G)'),
            Line2D([0], [0], color='orange', lw=2, label=f'{PORT_MPO}-{PORT_SFP} (10G)'),
            Line2D([0], [0], color='purple', lw=2, label=f'{PORT_SFP}-{PORT_SFP} (10G)')
        ]
        # 设置图例字体
        legend_prop_small = copy.copy(self.chinese_font_prop)
        legend_prop_small.set_size('small')
        self.axes.legend(handles=legend_elements, loc='best', prop=legend_prop_small)

        # 调整布局防止标签重叠或出界
        try:
             self.fig.tight_layout(rect=[0, 0.03, 1, 1]) # rect 留出底部空间给端口总数
        except ValueError as e:
             print(f"警告: tight_layout 失败: {e}") # 有时在特定情况下会失败

        self.draw_idle() # 异步绘制

        return self.fig, pos

