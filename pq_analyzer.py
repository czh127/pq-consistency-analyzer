# -*- coding: utf-8 -*-
from __future__ import annotations  # 类型提示延迟求值，支持重型库延迟加载
"""
PQ产品一致性检测分析程序 - v2.3 优化版
=====================================
所有功能完整保留，在原 v2.2 基础上进行以下架构优化：

【架构改善】
1. AnalysisState 类替代全局变量 — 线程安全的状态管理
2. threading.Event 替代 time.sleep 轮询 — 暂停时零 CPU 占用
3. 消除 print 全局劫持 — 改用 write_log 统一日志输出
4. ImageViewer LRU 缓存策略 — 固定容量，避免内存泄漏

【代码质量改善】
5. 工具函数添加类型提示
6. 异常处理规范化（具体异常类型替代裸 except）
7. 提取公共绘图样式函数 apply_common_axis_style
8. 资源管理使用 with 语句 / finally 确保释放

【配置改善】
9. 新增 AnalysisConfig dataclass（可选使用，兼容原 dict 配置）
"""

import sys
import os
import glob
import threading
from typing import List, Dict, Optional, Tuple, Callable
from enum import IntEnum
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import warnings
from datetime import datetime
from tkinter import Checkbutton
import time
import tempfile
import json

warnings.filterwarnings('ignore')

# 重型库（pandas / numpy / matplotlib / openpyxl / PIL）使用延迟加载
# 启动时只加载 tkinter，先弹出启动画面，再在进度条中逐步加载其他库
# 这样用户点击 .py 后能立刻看到窗口，而不是长时间无响应
# 加载逻辑见 main() 函数


# ============================================
# 0. 状态管理 - AnalysisState（线程安全，替代全局变量）
# ============================================

class AnalysisStatus(IntEnum):
    """分析状态枚举"""
    IDLE = 0
    RUNNING = 1
    PAUSED = 2
    STOPPED = 3


class AnalysisState:
    """
    分析状态管理器 - 线程安全
    替代原来的 4 个全局变量：_stop_analysis, _pause_analysis, _is_paused, _stopped_by_user
    使用 threading.Event 实现暂停等待，替代 time.sleep 轮询
    """
    
    def __init__(self):
        self._lock = threading.Lock()
        self._status = AnalysisStatus.IDLE
        self._pause_event = threading.Event()
        self._pause_event.set()  # 默认不暂停（set = 可通行）
    
    @property
    def status(self) -> AnalysisStatus:
        with self._lock:
            return self._status
    
    def is_running(self) -> bool:
        with self._lock:
            return self._status == AnalysisStatus.RUNNING
    
    def is_paused(self) -> bool:
        with self._lock:
            return self._status == AnalysisStatus.PAUSED
    
    def is_stopped(self) -> bool:
        with self._lock:
            return self._status == AnalysisStatus.STOPPED
    
    def start(self):
        """开始分析"""
        with self._lock:
            self._status = AnalysisStatus.RUNNING
            self._pause_event.set()
    
    def pause(self):
        """暂停分析"""
        with self._lock:
            if self._status == AnalysisStatus.RUNNING:
                self._status = AnalysisStatus.PAUSED
                self._pause_event.clear()
    
    def resume(self):
        """恢复分析"""
        with self._lock:
            if self._status == AnalysisStatus.PAUSED:
                self._status = AnalysisStatus.RUNNING
                self._pause_event.set()
    
    def stop(self):
        """停止分析"""
        with self._lock:
            self._status = AnalysisStatus.STOPPED
            self._pause_event.set()  # 唤醒暂停中的线程
    
    def reset(self):
        """重置状态"""
        with self._lock:
            self._status = AnalysisStatus.IDLE
            self._pause_event.set()
    
    def wait_if_paused(self, timeout: float = 0.2) -> bool:
        """
        如果处于暂停状态则等待（Event 等待，零 CPU 占用）
        返回 False 表示被停止，True 表示可以继续
        """
        # 快速路径：没暂停直接返回
        if self._pause_event.is_set() and not self.is_stopped():
            return True
        
        while not self._pause_event.is_set():
            if self.is_stopped():
                return False
            self._pause_event.wait(timeout=timeout)
        return not self.is_stopped()


# 全局状态实例（唯一的全局状态对象，替代原来 4 个全局变量）
_analysis_state = AnalysisState()


# ============================================
# 0. 启动画面
# ============================================
class SplashScreen:
    """启动画面，在程序初始化时显示"""
    def __init__(self):
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.configure(bg='#3370ff')
        
        width = 500
        height = 300
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
        main_frame = tk.Frame(self.root, bg='#3370ff')
        main_frame.pack(fill='both', expand=True)
        
        title_label = tk.Label(
            main_frame,
            text="🔬",
            font=('Segoe UI', 60),
            bg='#3370ff',
            fg='white'
        )
        title_label.pack(pady=(40, 5))
        
        name_label = tk.Label(
            main_frame,
            text="PQ产品一致性检测分析程序",
            font=('Microsoft YaHei', 18, 'bold'),
            bg='#3370ff',
            fg='white'
        )
        name_label.pack()
        
        version_label = tk.Label(
            main_frame,
            text="版本 2.3 优化版",
            font=('Microsoft YaHei', 10),
            bg='#3370ff',
            fg='#b3d4ff'
        )
        version_label.pack(pady=(5, 15))
        
        self.progress = ttk.Progressbar(
            main_frame,
            length=300,
            mode='determinate',
            maximum=100
        )
        self.progress.pack(pady=10)
        
        self.status_var = tk.StringVar(value="正在初始化...")
        status_label = tk.Label(
            main_frame,
            textvariable=self.status_var,
            font=('Microsoft YaHei', 9),
            bg='#3370ff',
            fg='#b3d4ff'
        )
        status_label.pack(pady=5)
        
        copyright_label = tk.Label(
            main_frame,
            text="© 2026 All Rights Reserved",
            font=('Microsoft YaHei', 8),
            bg='#3370ff',
            fg='#88bfff'
        )
        copyright_label.pack(side='bottom', pady=10)
        
        self.progress_value = 0
        
    def update_progress(self, value, status=""):
        if value > 100:
            value = 100
        self.progress_value = value
        self.progress['value'] = value
        if status:
            self.status_var.set(status)
        self.root.update()
    
    def close(self):
        self.root.destroy()


def resource_path(relative_path):
    """获取资源的绝对路径，适用于PyInstaller打包后的exe"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ============================================
# 1. 默认配置（新增图例简化配置）
# ============================================
DEFAULT_CONFIG = {
    # ---- 图例文件名简化（新增） ----
    "LEGEND_KEYWORD": "",      # 识别文字
    "LEGEND_MODE": 0,          # 0=删除前面不保留, 1=删除前面保留, 2=删除后面不保留, 3=删除后面保留
    "OUTPUT_EXCEL": "产品检测结果总表.xlsx",
    "ENCODING": "gbk",
    "COL_CURRENT": 1,
    "COL_FLOW": 4,
    "COL_PRESS": 10,
    "REF_FLOW": [5.0, 10.0, 20.0, 40.0, 60.0],
    "PLOT_FLOW": 20.0,
    "STD_LIMIT": 10.0,
    "STD_LINE1": {"value": 10.0, "show": True},
    "TOLERANCE": 1.00,
    "CONSISTENCY_LIMIT_PCT": 5.0,
    "CONSISTENCY_MODE": 0,
    "BASE_FILE_NAME": "产品1.csv",
    "KEY_FLOW": 20.0,
    "HIGH_FLOW": 40.0,
    "KEY_CURRENTS": [0, 300, 600, 900],
    "FILTER_WINDOW": 201,
    "FILTER_ON": True,
    "X_AXIS_EXTEND": 8,
    "HIDE_TOP_RIGHT_BORDER": True,
    "PERF_SHOW_LEGEND": False,
    "PQ_COMPARE_SHOW_LEGEND": True,
    "HYSTERESIS_SHOW_LEGEND": True,
    "CON_SHOW_LEGEND": False,
    "AVG_PRESS_SHOW_LEGEND": False,
    "GOBACK_SHOW_LEGEND": False,
    "GOBACK_SUMMARY_SHOW_LEGEND": False,
    "PERF_TITLE_SUFFIX": " 实测",
    "PERF_TITLE_SIZE": 36,
    "PERF_TITLE_BOLD": False,
    "PERF_TITLE_OFFSET_X": 0.5,
    "PERF_TITLE_OFFSET_Y": 1.1,
    "PERF_AXIS_LABEL_SIZE": 20,
    "PERF_AXIS_LABEL_BOLD": False,
    "PERF_TICK_LABEL_SIZE": 20,
    "PERF_LINE_WIDTH": 2.0,
    "PERF_TICK_WIDTH": 1.2,
    "PERF_TEXT_OFFSET_X": 1.5,
    "PERF_TEXT_OFFSET_Y": 0.0,
    "PERF_CURR_TEXT_SIZE": 20,
    "PERF_CURR_TEXT_BOLD": False,
    "CON_TITLE": "一致性",
    "CON_TITLE_SIZE": 24,
    "CON_TITLE_BOLD": False,
    "CON_TITLE_OFFSET_X": 0.5,
    "CON_TITLE_OFFSET_Y": 1.1,
    "CON_AXIS_LABEL_SIZE": 20,
    "CON_AXIS_LABEL_BOLD": False,
    "CON_TICK_LABEL_SIZE": 16,
    "CON_LINE_WIDTH": 2.0,
    "CON_TICK_WIDTH": 1.2,
    "LABEL_POS_0mA": -16,
    "LABEL_POS_OTHER": 9,
    "CON_LABEL_SIZE": 15,
    "CON_CURR_TEXT_SIZE": 20,
    "GOBACK_FIG_SIZE": [12, 6],
    "GOBACK_DPI": 100,
    "GOBACK_LINE_COLOR": "#00B0F0",
    "GOBACK_TITLE_SIZE": 36,
    "GOBACK_TITLE_BOLD": False,
    "GOBACK_TITLE_OFFSET_X": 0.5,
    "GOBACK_TITLE_OFFSET_Y": 1.1,
    "GOBACK_LINE_WIDTH": 2.0,
    "SUMMARY_FIG_SIZE": [15, 8],
    "SUMMARY_DPI": 150,
    "SUMMARY_LINE_WIDTH": 1.0,
    "SUMMARY_TITLE_SIZE": 20,
    "SUMMARY_AXIS_FONT_SIZE": 14,
    "SUMMARY_PAD": 20,
    "RUN_PERFORMANCE_PLOTS": True,
    "RUN_PQ_COMPARE_PLOT": True,
    "RUN_HYSTERESIS_PLOT": True,
    "RUN_CONSISTENCY_PLOT": True,
    "RUN_AVG_PRESSURE_PLOT": True,
    "RUN_EXCEL_CHARTS": True,
    "RUN_EXCEL_REPORT": True,
    "RUN_SINGLE_GOBACK_PQ": True,
    "RUN_ALL_GOBACK_SUMMARY": True,
    "PLOT_DPI": 150,
    "SUFFIX_AUTO_NUM": True,
    "PRINT_DEBUG_LOG": False,
}

# ============================================
# 2. GUI中文名称与分组
# ============================================
PARAM_CN_NAME = {
    # ---- 新增图例简化 ----
    "LEGEND_KEYWORD": "识别文字",
    "LEGEND_MODE": "删除模式\n 0=删除识别文字以及识别文字前面, \n 1=删除识别文字前面,\n 2=删除识别文字以及识别文字后面, \n 3=删除识别文字后面",
    "OUTPUT_EXCEL": "Excel报告文件名",
    "ENCODING": "CSV文件编码",
    "COL_CURRENT": "电流列索引",
    "COL_FLOW": "流量列索引",
    "COL_PRESS": "压差列索引",
    "REF_FLOW": "标准流量点 (逗号分隔)",
    "PLOT_FLOW": "迟滞图流量点 (L/min)",
    "STD_LIMIT": "迟滞合格上限 (bar)",
    "STD_LINE1__value": "标准线数值",
    "STD_LINE1__show": "显示标准线",
    "TOLERANCE": "流量匹配容差 (L/min)",
    "CONSISTENCY_LIMIT_PCT": "一致性参考阈值 (%)",
    "CONSISTENCY_MODE": "均值计算模式 (0/1/2):\n模式 0 → (最大值 + 最小值) / 2  对称均值；\n模式 1 → 所有文件真实平均值；\n模式 2 → 以基准文件算一致性",
    "BASE_FILE_NAME": "基准文件名 (模式2)",
    "KEY_FLOW": "加权重点流量 (L/min)",
    "HIGH_FLOW": "高流量参考点 (L/min)",
    "KEY_CURRENTS": "重点电流序列 (逗号分隔)",
    "FILTER_WINDOW": "滤波窗口 (奇数)",
    "FILTER_ON": "启用滑动滤波",
    "X_AXIS_EXTEND": "X轴右侧留白 (L/min)",
    "HIDE_TOP_RIGHT_BORDER": "隐藏上右边框",
    "PERF_SHOW_LEGEND": "性能图图例",
    "PQ_COMPARE_SHOW_LEGEND": "PQ对比图图例",
    "HYSTERESIS_SHOW_LEGEND": "迟滞图图例",
    "CON_SHOW_LEGEND": "一致性图图例",
    "AVG_PRESS_SHOW_LEGEND": "压差平均值图图例",
    "GOBACK_SHOW_LEGEND": "去回程单图图例",
    "GOBACK_SUMMARY_SHOW_LEGEND": "去回程总对比图例",
    "PERF_TITLE_SUFFIX": "性能图标题后缀",
    "PERF_TITLE_SIZE": "性能图标题字号",
    "PERF_TITLE_BOLD": "标题加粗",
    "PERF_TITLE_OFFSET_X": "标题水平偏移",
    "PERF_TITLE_OFFSET_Y": "标题垂直偏移",
    "PERF_AXIS_LABEL_SIZE": "坐标轴标签字号",
    "PERF_AXIS_LABEL_BOLD": "坐标轴标签加粗",
    "PERF_TICK_LABEL_SIZE": "刻度值字号",
    "PERF_LINE_WIDTH": "性能曲线宽度",
    "PERF_TICK_WIDTH": "刻度线宽度",
    "PERF_TEXT_OFFSET_X": "电流标注X偏移",
    "PERF_TEXT_OFFSET_Y": "电流标注Y偏移",
    "PERF_CURR_TEXT_SIZE": "电流标注字号",
    "PERF_CURR_TEXT_BOLD": "电流标注加粗",
    "CON_TITLE": "一致性图主标题",
    "CON_TITLE_SIZE": "一致性图标题字号",
    "CON_TITLE_BOLD": "一致性标题加粗",
    "CON_TITLE_OFFSET_X": "一致性标题水平偏移",
    "CON_TITLE_OFFSET_Y": "一致性标题垂直偏移",
    "CON_AXIS_LABEL_SIZE": "一致性坐标轴标签字号",
    "CON_AXIS_LABEL_BOLD": "一致性坐标轴加粗",
    "CON_TICK_LABEL_SIZE": "一致性刻度字号",
    "CON_LINE_WIDTH": "一致性曲线宽度",
    "CON_TICK_WIDTH": "一致性刻度线宽度",
    "LABEL_POS_0mA": "0mA标签垂直偏移",
    "LABEL_POS_OTHER": "其他电流标签垂直偏移",
    "CON_LABEL_SIZE": "偏差标签字号",
    "CON_CURR_TEXT_SIZE": "曲线末端电流标注字号",
    "GOBACK_FIG_SIZE": "去回程画布尺寸 (逗号分隔)",
    "GOBACK_DPI": "去回程图分辨率",
    "GOBACK_LINE_COLOR": "去回程曲线颜色",
    "GOBACK_TITLE_SIZE": "去回程标题字号",
    "GOBACK_TITLE_BOLD": "去回程标题加粗",
    "GOBACK_TITLE_OFFSET_X": "去回程标题水平偏移",
    "GOBACK_TITLE_OFFSET_Y": "去回程标题垂直偏移",
    "GOBACK_LINE_WIDTH": "去回程曲线宽度",
    "SUMMARY_FIG_SIZE": "总对比图画布尺寸",
    "SUMMARY_DPI": "总对比图分辨率",
    "SUMMARY_LINE_WIDTH": "总对比图曲线宽度",
    "SUMMARY_TITLE_SIZE": "总对比图标题字号",
    "SUMMARY_AXIS_FONT_SIZE": "总对比图坐标轴字号",
    "SUMMARY_PAD": "总对比图边距",
    "RUN_PERFORMANCE_PLOTS": "生成单文件PQ散点图",
    "RUN_PQ_COMPARE_PLOT": "生成PQ对比散点图",
    "RUN_HYSTERESIS_PLOT": "生成迟滞图",
    "RUN_CONSISTENCY_PLOT": "生成一致性图",
    "RUN_AVG_PRESSURE_PLOT": "生成压差平均值图",
    "RUN_EXCEL_CHARTS": "Excel内嵌图表",
    "RUN_EXCEL_REPORT": "生成Excel报告（必要）",
    "RUN_SINGLE_GOBACK_PQ": "生成单文件PQ图",
    "RUN_ALL_GOBACK_SUMMARY": "生成PQ总图",
    "PLOT_DPI": "基础图片DPI",
    "SUFFIX_AUTO_NUM": "自动添加序号防覆盖",
    "PRINT_DEBUG_LOG": "打印调试日志"
}

# 分组顺序调整：新增“图例文件名简化”放在最前面
GROUPS = [
    # ==== 新增：图例文件名简化（放在最前面） ====
    ("标准检测流量点", ["REF_FLOW"]),
    ("一致性计算模式 & 判定参数", ["CONSISTENCY_LIMIT_PCT", "CONSISTENCY_MODE", "BASE_FILE_NAME"]),
    ("图例文件名简化", ["LEGEND_KEYWORD", "LEGEND_MODE"]),
    ("图例独立开关", ["PERF_SHOW_LEGEND", "PQ_COMPARE_SHOW_LEGEND", "HYSTERESIS_SHOW_LEGEND",
                      "CON_SHOW_LEGEND", "AVG_PRESS_SHOW_LEGEND", "GOBACK_SHOW_LEGEND", "GOBACK_SUMMARY_SHOW_LEGEND"]),
    ("功能运行总开关", ["RUN_PERFORMANCE_PLOTS", "RUN_PQ_COMPARE_PLOT", "RUN_HYSTERESIS_PLOT", "RUN_CONSISTENCY_PLOT",
                      "RUN_AVG_PRESSURE_PLOT", "RUN_EXCEL_CHARTS", "RUN_EXCEL_REPORT", "RUN_SINGLE_GOBACK_PQ", "RUN_ALL_GOBACK_SUMMARY"]),
    ("输出文件基础配置", ["OUTPUT_EXCEL", "ENCODING"]),
    ("CSV数据列索引配置", ["COL_CURRENT", "COL_FLOW", "COL_PRESS"]),
    ("迟滞 & 流量匹配容错配置", ["PLOT_FLOW", "STD_LIMIT", "TOLERANCE", "STD_LINE1__value", "STD_LINE1__show"]),
    ("加权一致性分析参数", ["KEY_FLOW", "HIGH_FLOW", "KEY_CURRENTS"]),
    ("滑动平均滤波配置", ["FILTER_WINDOW", "FILTER_ON"]),
    ("图表通用全局样式", ["X_AXIS_EXTEND", "HIDE_TOP_RIGHT_BORDER"]),
    ("性能图表样式配置", ["PERF_TITLE_SUFFIX", "PERF_TITLE_SIZE", "PERF_TITLE_BOLD", "PERF_TITLE_OFFSET_X",
                         "PERF_TITLE_OFFSET_Y", "PERF_AXIS_LABEL_SIZE", "PERF_AXIS_LABEL_BOLD", "PERF_TICK_LABEL_SIZE",
                         "PERF_LINE_WIDTH", "PERF_TICK_WIDTH"]),
    ("性能图表电流标签样式", ["PERF_TEXT_OFFSET_X", "PERF_TEXT_OFFSET_Y", "PERF_CURR_TEXT_SIZE", "PERF_CURR_TEXT_BOLD"]),
    ("一致性图表样式配置", ["CON_TITLE", "CON_TITLE_SIZE", "CON_TITLE_BOLD", "CON_TITLE_OFFSET_X", "CON_TITLE_OFFSET_Y",
                         "CON_AXIS_LABEL_SIZE", "CON_AXIS_LABEL_BOLD", "CON_TICK_LABEL_SIZE", "CON_LINE_WIDTH", "CON_TICK_WIDTH"]),
    ("一致性/均值图数据标签位置样式", ["LABEL_POS_0mA", "LABEL_POS_OTHER", "CON_LABEL_SIZE", "CON_CURR_TEXT_SIZE"]),
    ("去程回程PQ图表独立配置", ["GOBACK_FIG_SIZE", "GOBACK_DPI", "GOBACK_LINE_COLOR", "GOBACK_TITLE_SIZE",
                             "GOBACK_TITLE_BOLD", "GOBACK_TITLE_OFFSET_X", "GOBACK_TITLE_OFFSET_Y", "GOBACK_LINE_WIDTH"]),
    ("全部产品总对比PQ图独立参数", ["SUMMARY_FIG_SIZE", "SUMMARY_DPI", "SUMMARY_LINE_WIDTH", "SUMMARY_TITLE_SIZE",
                               "SUMMARY_AXIS_FONT_SIZE", "SUMMARY_PAD"]),
    ("扩展通用配置", ["PLOT_DPI", "SUFFIX_AUTO_NUM", "PRINT_DEBUG_LOG"]),
]

FLAT_DEFAULTS = {}
for k, v in DEFAULT_CONFIG.items():
    if not isinstance(v, (dict, list)):
        FLAT_DEFAULTS[k] = v
FLAT_DEFAULTS["STD_LINE1__value"] = DEFAULT_CONFIG["STD_LINE1"]["value"]
FLAT_DEFAULTS["STD_LINE1__show"] = DEFAULT_CONFIG["STD_LINE1"]["show"]
FLAT_DEFAULTS["REF_FLOW"] = ",".join(map(str, DEFAULT_CONFIG["REF_FLOW"]))
FLAT_DEFAULTS["KEY_CURRENTS"] = ",".join(map(str, DEFAULT_CONFIG["KEY_CURRENTS"]))
FLAT_DEFAULTS["GOBACK_FIG_SIZE"] = ",".join(map(str, DEFAULT_CONFIG["GOBACK_FIG_SIZE"]))
FLAT_DEFAULTS["SUMMARY_FIG_SIZE"] = ",".join(map(str, DEFAULT_CONFIG["SUMMARY_FIG_SIZE"]))

# 新增图例简化配置的默认值
FLAT_DEFAULTS["LEGEND_KEYWORD"] = ""
FLAT_DEFAULTS["LEGEND_MODE"] = 0


# ============================================
# 3. 全局辅助函数（新增图例简化函数）
# ============================================
def simplify_filename_for_legend(filename: str, config: Dict) -> str:
    """
    根据配置简化文件名
    
    4种模式：
    0: 删除前面-不保留识别文字  → 删除识别文字及其之前全部
    1: 删除前面-保留识别文字    → 删除识别文字之前全部，保留识别文字本身
    2: 删除后面-不保留识别文字  → 删除识别文字及其之后全部
    3: 删除后面-保留识别文字    → 删除识别文字之后全部，保留识别文字本身
    """
    keyword = config.get("LEGEND_KEYWORD", "")
    if not keyword:
        return filename
    
    mode = config.get("LEGEND_MODE", 0)
    pos = filename.find(keyword)
    
    if pos == -1:
        return filename
    
    if mode == 0:      # 删除前面，不保留识别文字
        return filename[pos + len(keyword):]
    elif mode == 1:    # 删除前面，保留识别文字
        return filename[pos:]
    elif mode == 2:    # 删除后面，不保留识别文字
        return filename[:pos]
    elif mode == 3:    # 删除后面，保留识别文字
        return filename[:pos + len(keyword)]
    else:
        return filename


def save_safe_plot(fig: plt.Figure, filename: str, OUTPUT_FOLDER: str, dpi: int = 150, bbox_inches: str = 'tight') -> str:
    path = os.path.join(OUTPUT_FOLDER, filename)
    base, ext = os.path.splitext(path)
    counter = 1
    while os.path.exists(path):
        path = f"{base}({counter}){ext}"
        counter += 1
    fig.savefig(path, dpi=dpi, bbox_inches=bbox_inches)
    plt.close(fig)
    return path

def get_closest_index(arr: np.ndarray, target: float) -> int:
    return np.argmin(np.abs(arr - target))

def safe_percent(numerator: float, denominator: float, default: float = 0.0) -> float:
    try:
        return round((numerator / denominator) * 100, 1)
    except ZeroDivisionError:
        return default

def apply_common_axis_style(ax, style_dict, xlabel="", ylabel="", hide_top_right=True):
    """
    v2.3 新增：统一应用坐标轴样式，消除重复代码
    
    Args:
        ax: matplotlib 坐标轴对象
        style_dict: 样式字典（含 axis_label_size, axis_label_bold, tick_label_size, tick_width 等）
        xlabel: X 轴标签文本
        ylabel: Y 轴标签文本
        hide_top_right: 是否隐藏上右边框
    """
    if xlabel:
        ax.set_xlabel(
            xlabel,
            fontsize=style_dict.get("PERF_AXIS_LABEL_SIZE", 14),
            weight="bold" if style_dict.get("PERF_AXIS_LABEL_BOLD", False) else "normal"
        )
    if ylabel:
        ax.set_ylabel(
            ylabel,
            fontsize=style_dict.get("PERF_AXIS_LABEL_SIZE", 14),
            weight="bold" if style_dict.get("PERF_AXIS_LABEL_BOLD", False) else "normal"
        )
    ax.tick_params(
        axis='both',
        labelsize=style_dict.get("PERF_TICK_LABEL_SIZE", 12),
        width=style_dict.get("PERF_TICK_WIDTH", 1.0)
    )
    if hide_top_right:
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)


def filter_press_by_current(df: pd.DataFrame, window: int, filter_on: bool) -> pd.DataFrame:
    if not filter_on:
        return df
    df_out = df.copy()
    for curr, group in df.groupby("电流值"):
        filtered_press = group["压差值"].rolling(window=window, center=True, min_periods=1).mean()
        df_out.loc[group.index, "压差值滤波"] = filtered_press
    return df_out

def clean_invalid_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.replace([np.inf, -np.inf], np.nan)
    df = df.dropna()
    return df

def analyze_consistency_influence_final(df_valid: pd.DataFrame, config: Dict) -> Tuple:
    influence = []
    KEY_CURRENTS = config["KEY_CURRENTS"]
    KEY_FLOW = config["KEY_FLOW"]
    HIGH_FLOW = config["HIGH_FLOW"]

    group_mean = df_valid.groupby(["电流值", "参考流量值", "数据分区"])["实际压差值"].mean().reset_index()
    group_mean.rename(columns={"实际压差值": "工况均值"}, inplace=True)
    df_merge = df_valid.merge(group_mean, on=["电流值", "参考流量值", "数据分区"], how="left")

    for _, row in df_merge.iterrows():
        I, rf, part = row["电流值"], row["参考流量值"], row["数据分区"]
        f = row["数据源文件"]
        P = row["实际压差值"]
        mean_ir = row["工况均值"]
        dev = P - mean_ir

        if I in [0.0, 300.0] and rf in [5.0, 10.0]:
            continue

        weight = 1.0
        if I in KEY_CURRENTS:
            if I == 0:
                base_w = 10.0
            elif I == 300:
                base_w = 8.0
            elif I == 600:
                base_w = 6.0
            elif I == 900:
                base_w = 4.0
            else:
                base_w = 1.0
            weight = base_w * 2.0 if rf == KEY_FLOW else base_w
        elif I > 900:
            weight = 1.0

        influence.append({
            "file": f, "I": I, "rf": rf,
            "P": P, "mean_ir": mean_ir,
            "dev": dev, "weight": weight,
            "weighted_dev": dev * weight
        })

    if not influence:
        return None, None, None, None

    df_inf = pd.DataFrame(influence)
    file_wdev = df_inf.groupby("file")["weighted_dev"].sum()
    file_weight = df_inf.groupby("file")["weight"].sum()
    file_avg_dev = file_wdev / file_weight
    file_score = file_avg_dev.abs().sort_values(ascending=False)

    global_mean_all = df_inf["mean_ir"].mean()
    suggest = {}
    for f in file_avg_dev.index:
        f_mean = df_inf[df_inf["file"] == f]["P"].mean()
        adj_mean = f_mean + file_avg_dev[f]
        scale = round(global_mean_all / adj_mean, 3) if adj_mean != 0 else 1.0
        suggest[f] = scale

    total_abs = file_score.sum()
    impact_rate = {f: round(s / total_abs * 100, 1) for f, s in file_score.items()}
    worst_file = file_score.index[0]

    return file_score, impact_rate, worst_file, suggest

def get_real_0mA60L_press(df_valid, files):
    result = {}
    target_I = 0.0
    target_rf = 60.0
    for f in files:
        sub = df_valid[
            (df_valid["数据源文件"] == f) &
            (df_valid["电流值"] == target_I) &
            (df_valid["参考流量值"] == target_rf) &
            (df_valid["数据分区"] == "前50%")
        ]
        if not sub.empty:
            p = round(sub["实际压差值"].iloc[0], 2)
        else:
            p = 0.0
        result[f] = p
    return result

def calculate_consistency(df_valid: pd.DataFrame, config: Dict) -> pd.DataFrame:
    con_data = []
    mode = config["CONSISTENCY_MODE"]
    base_file = config["BASE_FILE_NAME"]

    for (curr, rf, part), group in df_valid.groupby(["电流值", "参考流量值", "数据分区"]):
        vals = group["实际压差值"].dropna()
        files = group["数据源文件"]

        if len(vals) < 1:
            continue

        if mode == 0:
            maxv = vals.max()
            minv = vals.min()
            mean_val = (maxv + minv) / 2
            pos_diff = maxv - mean_val
            neg_diff = mean_val - minv
            max_file = group.loc[vals.idxmax(), "数据源文件"]
            min_file = group.loc[vals.idxmin(), "数据源文件"]
        elif mode == 1:
            mean_val = vals.mean()
            maxv = vals.max()
            minv = vals.min()
            pos_diff = maxv - mean_val
            neg_diff = mean_val - minv
            max_file = group.loc[vals.idxmax(), "数据源文件"]
            min_file = group.loc[vals.idxmin(), "数据源文件"]
        elif mode == 2:
            base_row = group[group["数据源文件"] == base_file]
            if base_row.empty:
                continue
            mean_val = base_row["实际压差值"].iloc[0]
            maxv = vals.max()
            minv = vals.min()
            pos_diff = maxv - mean_val
            neg_diff = mean_val - minv
            max_file = group.loc[vals.idxmax(), "数据源文件"]
            min_file = group.loc[vals.idxmin(), "数据源文件"]
        else:
            continue

        pos_pct = safe_percent(pos_diff, mean_val)
        neg_pct = safe_percent(neg_diff, mean_val)

        sample_count = len(vals)
        std_dev = round(vals.std(), 6)
        try:
            if abs(mean_val) < 1e-6:
                cv_val = np.nan
            else:
                cv_val = round((std_dev / mean_val) * 100, 3)
        except:
            cv_val = np.nan

        con_data.append([
            round(curr, 2), round(rf, 1), part,
            round(maxv, 2), max_file,
            round(minv, 2), min_file,
            round(mean_val, 2),
            round(pos_diff, 2), round(pos_pct, 2),
            round(neg_diff, 2), round(neg_pct, 2),
            sample_count, std_dev,
            cv_val
        ])

    df_con = pd.DataFrame(con_data, columns=[
        "电流值", "参考流量值", "数据分区",
        "最大值", "最大值文件名",
        "最小值", "最小值文件名",
        "平均值",
        "正偏差bar", "正偏差%",
        "负偏差bar", "负偏差%",
        "样本数量", "压差标准差",
        "变异系数CV(%)"
    ])
    return df_con


# ============================================
# 4. 核心控制函数（基于 AnalysisState）
# ============================================

def stop_analysis():
    """停止分析（委托给全局 AnalysisState）"""
    _analysis_state.stop()

def pause_analysis():
    """暂停分析"""
    _analysis_state.pause()

def resume_analysis():
    """恢复分析"""
    _analysis_state.resume()

def reset_stop_flag():
    """重置停止/暂停标志"""
    _analysis_state.reset()

def is_paused() -> bool:
    """是否处于暂停状态"""
    return _analysis_state.is_paused()

def is_stopped() -> bool:
    """是否被用户停止"""
    return _analysis_state.is_stopped()

def get_script_directory() -> str:
    """获取脚本所在目录（兼容 PyInstaller）"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def check_stop_or_pause(log_file, write_log) -> bool:
    """
    检查停止或暂停状态
    使用 Event 等待替代 time.sleep 轮询，暂停时零 CPU 占用
    返回 True 表示应停止，False 表示可继续
    """
    if not _analysis_state.wait_if_paused(timeout=0.2):
        write_log("⚠️ 用户停止分析\n")
        if log_file:
            log_file.write("⚠️ 用户停止分析\n")
            log_file.flush()
        return True
    return False


# ============================================
# 5. 核心分析函数（修改图例部分）
# ============================================
def run_analysis(config, log_callback=None, work_dir=None, resume=False):
    # v2.3 改善：使用 AnalysisState 管理状态，无需 global 声明
    if not resume:
        _analysis_state.start()
    
    import numpy as np
    import matplotlib.pyplot as plt

    if work_dir and os.path.exists(work_dir):
        os.chdir(work_dir)
    else:
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(base_dir)

    log_file_path = os.path.join(os.getcwd(), "analysis_log.txt")
    try:
        log_file = open(log_file_path, "a", encoding="utf-8")
    except PermissionError:
        temp_dir = tempfile.gettempdir()
        log_file_path = os.path.join(temp_dir, "analysis_log.txt")
        log_file = open(log_file_path, "a", encoding="utf-8")
    
    if not resume:
        log_file.write("\n" + "=" * 70 + "\n")
        log_file.write(f"===== PQ分析日志开始 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) =====\n")
        log_file.write(f"📁 工作目录: {os.getcwd()}\n")
        log_file.flush()
    else:
        log_file.write("\n" + "-" * 70 + "\n")
        log_file.write(f"🔄 恢复分析 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')})\n")
        log_file.flush()
    
    def write_log(msg):
        # v2.3 改善：使用 Event 等待替代 sleep 轮询
        if not _analysis_state.wait_if_paused(timeout=0.2):
            log_file.write("⚠️ 用户停止分析\n")
            log_file.flush()
            return
        log_file.write(msg)
        log_file.flush()
        if log_callback:
            if hasattr(run_analysis, 'root'):
                run_analysis.root.after(0, lambda: log_callback(msg))
            else:
                log_callback(msg)

    # v2.3 改善：不再劫持全局 print，统一使用 write_log 输出
    # 保留 original_print 变量供调试使用
    original_print = print

    try:
        if not resume:
            write_log("=" * 70 + "\n")
            write_log("🚀 PQ 产品一致性检测分析程序 启动成功\n")
            write_log("=" * 70 + "\n")

            mode = config["CONSISTENCY_MODE"]
            write_log("\n📏 【平均值计算方式】\n")
            if mode == 0:
                write_log("   模式 0 → (最大值 + 最小值) / 2  对称均值 \n")
            elif mode == 1:
                write_log("   模式 1 → 所有文件真实平均值  🎯 推荐\n")
            elif mode == 2:
                write_log(f"   模式 2 → 以基准文件 [{config['BASE_FILE_NAME']}] 为均值\n")
            else:
                write_log("   模式未知 → 使用默认真实平均值\n")
            write_log("=" * 70 + "\n")

        if config.get("SUFFIX_AUTO_NUM", True):
            OUTPUT_FOLDER = "检测结果输出"
            if not os.path.exists(OUTPUT_FOLDER):
                os.makedirs(OUTPUT_FOLDER)
            else:
                count = 1
                while True:
                    new_folder = f"{OUTPUT_FOLDER}_{count}"
                    if not os.path.exists(new_folder):
                        OUTPUT_FOLDER = new_folder
                        os.makedirs(OUTPUT_FOLDER)
                        break
                    count += 1
        else:
            OUTPUT_FOLDER = "检测结果输出"
            os.makedirs(OUTPUT_FOLDER, exist_ok=True)

        if not resume:
            write_log("\n📂 【1/12】正在扫描当前目录的CSV数据文件...\n")
            write_log(f"📁 扫描路径: {os.getcwd()}\n")
            csv_files = glob.glob("*.csv")
            write_log(f"✅ 扫描完成：共找到 {len(csv_files)} 个有效数据文件\n")
            if len(csv_files) > 0:
                for f in csv_files:
                    write_log(f"   📄 {f}\n")

            today_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            if len(csv_files) > 0:
                first_csv_name = os.path.splitext(csv_files[0])[0]
                prefix = f"{today_str}-{first_csv_name}等-"
            else:
                prefix = f"{today_str}-无数据-"
            excel_base_name = prefix + config["OUTPUT_EXCEL"]

            raw_data = []
            pq_comparison_data = []
            file_cleaned_data = {}

            write_log("\n🔍 【2/12】正在读取数据并执行滤波处理...\n")
            for file in csv_files:
                if check_stop_or_pause(log_file, write_log):
                    log_file.close()
                    return False
                
                try:
                    df = pd.read_csv(
                        file, 
                        usecols=[config["COL_CURRENT"], config["COL_FLOW"], config["COL_PRESS"]], 
                        encoding=config["ENCODING"]
                    )
                    df.columns = ["电流值", "流量值", "压差值"]
                except Exception as e:
                    write_log(f"⚠️ 读取文件 {file} 失败: {e}\n")
                    continue

                df = clean_invalid_data(df)
                if df.empty:
                    write_log(f"⚠️ 文件 {file} 无有效数据\n")
                    continue

                df["流量值原始"] = df["流量值"].copy()
                df["压差值原始"] = df["压差值"].copy()
                df["压差值滤波"] = df["压差值"].copy()
                df = filter_press_by_current(df, config["FILTER_WINDOW"], config["FILTER_ON"])

                df = clean_invalid_data(df)
                if df.empty:
                    continue

                df["压差值"] = df["压差值滤波"]
                file_short = os.path.splitext(file)[0]
                file_cleaned_data[file] = df.copy()

                for I, g in df.groupby("电流值"):
                    I = round(I, 2)
                    mid = len(g) // 2
                    first_half = g.iloc[:mid]
                    first_half = clean_invalid_data(first_half)
                    if first_half.empty:
                        continue

                    f_vals = first_half["流量值"].abs().values
                    p_vals = first_half["压差值"].abs().values
                    flow_points = []
                    press_points = []
                    for ref in config["REF_FLOW"]:
                        idx = get_closest_index(f_vals, ref)
                        fp = round(abs(f_vals[idx]), 4)
                        pp = round(abs(p_vals[idx]), 4)
                        flow_points.append(fp)
                        press_points.append(pp)

                    pq_comparison_data.append({
                        "file": file_short, "current": I,
                        "flow": flow_points, "press": press_points
                    })

                for I, g in df.groupby("电流值"):
                    I = round(I, 2)
                    mid = len(g) // 2
                    for part, part_data in [("前50%", g.iloc[:mid]), ("后50%", g.iloc[mid:])]:
                        part_data = clean_invalid_data(part_data)
                        if part_data.empty:
                            continue
                        f_vals = part_data["流量值"].abs().values
                        p_vals = part_data["压差值"].abs().values
                        for ref in config["REF_FLOW"]:
                            idx = get_closest_index(f_vals, ref)
                            real_flow = round(abs(f_vals[idx]), 4)
                            real_press = round(abs(p_vals[idx]), 4)
                            diff = round(abs(real_flow - ref), 4)
                            check = "TRUE" if diff <= config["TOLERANCE"] else "FALSE"
                            raw_data.append([file, I, part, ref, real_press, real_flow, diff, check])

            df_raw = pd.DataFrame(raw_data, columns=[
                "数据源文件", "电流值", "数据分区", "参考流量值", "实际压差值", "实际流量值", "差值数据", "校验结果"
            ])

            write_log("\n✅ 【3/12】正在筛选所有文件均合格的有效检测点位...\n")
            all_valid_groups = df_raw.groupby(["电流值", "参考流量值", "数据分区"])["校验结果"].apply(
                lambda x: (x == "TRUE").all()
            )
            valid_groups = all_valid_groups[all_valid_groups].index.tolist()
            df_valid = df_raw[df_raw.set_index(["电流值", "参考流量值", "数据分区"]).index.isin(valid_groups)].copy()
            write_log("✅ 数据预处理全部完成\n")

            if check_stop_or_pause(log_file, write_log):
                log_file.close()
                return False

            write_log("\n📊 【4/12】正在执行加权一致性影响分析...\n")
            file_score, impact_rate, worst_file, scale_suggest = analyze_consistency_influence_final(df_valid, config)
            write_log("✅ 一致性分析完成，已定位异常产品\n")

            write_log("\n📌 【5/12】正在读取0mA 60L/min 基准压差值...\n")
            all_files = df_valid["数据源文件"].unique()
            orig_0mA60L = get_real_0mA60L_press(df_valid, all_files)
            scale_0mA60L = {}
            for f in all_files:
                orig = orig_0mA60L.get(f, 0.0)
                scale = scale_suggest.get(f, 1.0)
                new_p = round(orig * scale, 2)
                scale_0mA60L[f] = (orig, new_p)
            write_log("✅ 关键点位数据提取完成\n")

            plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
            plt.rcParams['axes.unicode_minus'] = False

        if not resume:
            write_log("\n📈 【6/12】正在生成单产品性能曲线图...\n")
            if config["RUN_PERFORMANCE_PLOTS"]:
                file_list = df_valid["数据源文件"].unique()
                perf_folder = "性能图表"
                perf_full = os.path.join(OUTPUT_FOLDER, perf_folder)
                if not os.path.exists(perf_full):
                    os.makedirs(perf_full)

                for fig_idx, file in enumerate(file_list, 1):
                    if check_stop_or_pause(log_file, write_log):
                        log_file.close()
                        return False
                    
                    df_file = df_valid[(df_valid["数据源文件"] == file) & (df_valid["数据分区"] == "前50%")].copy()
                    if df_file.empty:
                        continue
                    fig, ax = plt.subplots(figsize=(12, 6), dpi=100)
                    curr_list = sorted(df_file["电流值"].unique())
                    for curr in curr_list:
                        sub = df_file[df_file["电流值"] == curr].sort_values("参考流量值")
                        ax.plot(sub["参考流量值"], sub["实际压差值"], marker='o', ms=4,
                                color='#00B0F0', lw=config["PERF_LINE_WIDTH"], label=f"{curr}mA")
                        x_last = sub["参考流量值"].iloc[-1]
                        y_last = sub["实际压差值"].iloc[-1]
                        ax.text(
                            x_last + config["PERF_TEXT_OFFSET_X"],
                            y_last + config["PERF_TEXT_OFFSET_Y"],
                            f"{curr}mA",
                            fontsize=config["PERF_CURR_TEXT_SIZE"],
                            weight="bold" if config["PERF_CURR_TEXT_BOLD"] else "normal",
                            va='center'
                        )
                    title_text = f"{os.path.splitext(file)[0]}{config['PERF_TITLE_SUFFIX']}"
                    ax.set_title(title_text, fontsize=config["PERF_TITLE_SIZE"],
                                 weight="bold" if config["PERF_TITLE_BOLD"] else "normal",
                                 x=config["PERF_TITLE_OFFSET_X"], y=config["PERF_TITLE_OFFSET_Y"])
                    ax.set_xlabel("流量 L/min", fontsize=config["PERF_AXIS_LABEL_SIZE"],
                                  weight="bold" if config["PERF_AXIS_LABEL_BOLD"] else "normal")
                    ax.set_ylabel("压差 bar", fontsize=config["PERF_AXIS_LABEL_SIZE"],
                                  weight="bold" if config["PERF_AXIS_LABEL_BOLD"] else "normal")
                    ax.tick_params(axis='both', labelsize=config["PERF_TICK_LABEL_SIZE"],
                                   width=config["PERF_TICK_WIDTH"])
                    ax.set_xlim(right=70)
                    if config["HIDE_TOP_RIGHT_BORDER"]:
                        ax.spines['top'].set_visible(False)
                        ax.spines['right'].set_visible(False)
                    if config["PERF_SHOW_LEGEND"]:
                        ax.legend(loc="upper right", frameon=False)
                    plt.tight_layout()
                    save_safe_plot(fig, os.path.join(perf_folder, f"性能图表-{fig_idx}.png"), OUTPUT_FOLDER, dpi=config["PLOT_DPI"])
                    write_log(f"  已生成 {file} 的性能图\n")
                write_log("✅ 单产品性能图表生成完成\n")
            else:
                write_log("ℹ️ 已跳过：单产品性能图表\n")

            if check_stop_or_pause(log_file, write_log):
                log_file.close()
                return False

        if not resume:
            write_log("\n📊 【7/12】正在生成所有产品PQ总对比散点图...\n")
            if config["RUN_PQ_COMPARE_PLOT"] and pq_comparison_data:
                fig, ax = plt.subplots(figsize=(18, 10), dpi=300)
                fig.set_facecolor('white')
                unique_files = sorted(list(set([item["file"] for item in pq_comparison_data])))
                unique_currents = sorted(list(set([item["current"] for item in pq_comparison_data])))

                import matplotlib.cm as cm
                colors = cm.nipy_spectral(np.linspace(0, 1, len(unique_files)))
                file_color_map = {f: colors[i] for i, f in enumerate(unique_files)}

                plotted_files = set()
                for f in unique_files:
                    dat = [d for d in pq_comparison_data if d["file"] == f]
                    for curr in unique_currents:
                        cdat = next((d for d in dat if d["current"] == curr), None)
                        if cdat:
                            show_label = f not in plotted_files
                            # ---- 简化图例 ----
                            simplified_name = simplify_filename_for_legend(f, config)
                            label = simplified_name if show_label else ""
                            if show_label:
                                plotted_files.add(f)
                            ax.plot(
                                cdat["flow"], cdat["press"],
                                marker='o', ms=3, linewidth=0.6,
                                color=file_color_map[f],
                                label=label,
                                alpha=0.8
                            )

                if unique_currents:
                    first_file = unique_files[0]
                    first_data = [d for d in pq_comparison_data if d["file"] == first_file]
                    for curr in unique_currents:
                        cdat = next((d for d in first_data if d["current"] == curr), None)
                        if cdat:
                            flows = cdat["flow"]
                            presses = cdat["press"]
                            if len(flows) > 0:
                                x_last = flows[-1]
                                y_last = presses[-1]
                                ax.text(
                                    x_last + config["PERF_TEXT_OFFSET_X"],
                                    y_last + config["PERF_TEXT_OFFSET_Y"],
                                    f"{curr}mA",
                                    fontsize=config["CON_CURR_TEXT_SIZE"],
                                    weight="bold" if config["CON_TITLE_BOLD"] else "normal",
                                    va='center'
                                )

                ax.set_title("所有产品滤波后前50% PQ曲线对比", fontsize=16, weight='bold')
                ax.set_xlabel("流量 L/min", fontsize=14)
                ax.set_ylabel("压差 bar", fontsize=14)
                ax.grid(alpha=0.3)
                ax.set_xlim(0, max(config["REF_FLOW"]) + 10)
                ax.tick_params(axis='both', labelsize=12)
                if config["HIDE_TOP_RIGHT_BORDER"]:
                    ax.spines['top'].set_visible(False)
                    ax.spines['right'].set_visible(False)

                if config["PQ_COMPARE_SHOW_LEGEND"]:
                    ncol = min(3, max(1, len(unique_files) // 12))
                    leg = ax.legend(loc="center left", bbox_to_anchor=(1.03, 0.5), ncol=ncol, fontsize=10, frameon=False)
                    leg.handlelength = 1.2
                plt.tight_layout(rect=[0, 0, 0.87, 1])
                save_safe_plot(fig, "所有产品PQ曲线对比图.png", OUTPUT_FOLDER, dpi=300)
                write_log("✅ PQ总对比散点图生成完成\n")
            else:
                write_log("ℹ️ 已跳过：PQ总对比散点图\n")

        if not resume:
            write_log("\n🔁 【8/12】正在计算产品迟滞数据...\n")
            hys_data = []
            for (f, c, rf), g in df_valid.groupby(["数据源文件", "电流值", "参考流量值"]):
                b = g[g["数据分区"] == "前50%"]["实际压差值"]
                a = g[g["数据分区"] == "后50%"]["实际压差值"]
                if len(b) == 1 and len(a) == 1:
                    hys_data.append([f, c, rf, round(abs(b.iloc[0] - a.iloc[0]), 2)])
            df_hys = pd.DataFrame(hys_data, columns=["数据源文件", "电流值", "参考流量值", "迟滞"])
            write_log("✅ 迟滞数据计算完成\n")

        if not resume:
            write_log("\n📐 【9/12】正在计算产品一致性数据...\n")
            df_con = calculate_consistency(df_valid, config)
            write_log("✅ 一致性数据计算完成\n")

            if check_stop_or_pause(log_file, write_log):
                log_file.close()
                return False

        if not resume:
            write_log("\n📄 【10/12】正在生成Excel完整检测报告...\n")
            if config["RUN_EXCEL_REPORT"]:
                wb = Workbook()
                if "Sheet" in wb.sheetnames:
                    del wb["Sheet"]

                ws1 = wb.create_sheet("原始检测数据")
                ws1.append(["数据源文件", "电流值", "数据分区", "参考流量值", "实际压差值", "实际流量值", "差值数据", "校验结果"])
                for _, r in df_raw.iterrows():
                    ws1.append(list(r))

                ws2 = wb.create_sheet("产品迟滞")
                ws2.append(["数据源文件", "电流值", "参考流量值", "迟滞"])
                for r in hys_data:
                    ws2.append(r)
                plot_hys = df_hys[df_hys["参考流量值"] == config["PLOT_FLOW"]].copy().sort_values(["数据源文件", "电流值"])
                ws2.cell(1, 8, "绘图文件名")
                ws2.cell(1, 9, "绘图电流")
                ws2.cell(1, 10, "绘图迟滞")
                plot_last_row = len(plot_hys) + 1
                for i, (_, r) in enumerate(plot_hys.iterrows(), 2):
                    ws2.cell(i, 8, r["数据源文件"])
                    ws2.cell(i, 9, r["电流值"])
                    ws2.cell(i, 10, r["迟滞"])
                ws2.cell(plot_last_row + 1, 8, "绘图数据结束")

                if config["RUN_EXCEL_CHARTS"]:
                    chart_hys = ScatterChart()
                    chart_hys.scatterStyle = "lineMarker"
                    chart_hys.title = "P-Q滞环@20L/min"
                    chart_hys.x_axis.title = "电流值"
                    chart_hys.y_axis.title = "迟滞/bar"

                    prod_list = []
                    for row in range(2, plot_last_row + 1):
                        prod = ws2.cell(row, 8).value
                        if prod and prod not in prod_list:
                            prod_list.append(prod)

                    COLOR_EXCEL = ["70AD47", "00B0F0", "FFC000", "FF0000", "C00000"]
                    for prod in prod_list:
                        sr = er = None
                        for row in range(2, plot_last_row + 1):
                            if ws2.cell(row, 8).value == prod:
                                sr = row if sr is None else sr
                                er = row
                        if sr is None or er is None:
                            continue

                        x = Reference(ws2, 9, sr, 9, er)
                        y = Reference(ws2, 10, sr, 10, er)

                        exceed_num = 0
                        for row in range(sr, er + 1):
                            hys_val = ws2.cell(row, 10).value
                            if hys_val and float(hys_val) > config["STD_LIMIT"]:
                                exceed_num += 1
                        color = COLOR_EXCEL[min(exceed_num, 4)]

                        ser = Series(y, x, title=prod)
                        ser.marker = Marker(size=3)
                        ser.graphicalProperties.line = LineProperties(solidFill=color, w=10000)
                        chart_hys.series.append(ser)
                    ws2.add_chart(chart_hys, "L2")

                ws3 = wb.create_sheet("一致性对比")
                ws3.append([
                    "电流值", "参考流量值", "数据分区",
                    "最大值", "最大值文件名", "最小值", "最小值文件名",
                    "平均值", "正偏差bar", "正偏差%", "负偏差bar", "负偏差%",
                    "样本数量", "压差标准差", "变异系数CV(%)"
                ])
                for _, r in df_con.iterrows():
                    ws3.append([
                        r["电流值"], r["参考流量值"], r["数据分区"],
                        r["最大值"], r["最大值文件名"], r["最小值"], r["最小值文件名"],
                        r["平均值"], r["正偏差bar"], r["正偏差%"], r["负偏差bar"], r["负偏差%"],
                        r["样本数量"], r["压差标准差"], r["变异系数CV(%)"]
                    ])

                def draw_con_chart(part, title, col, pos):
                    data = df_con[df_con["数据分区"] == part].sort_values(["电流值", "参考流量值"])
                    currs = sorted(data["电流值"].unique())
                    headers = ["电流值", "参考流量值", "平均值", "正偏差bar", "正偏差%", "负偏差bar", "负偏差%", "标签"]
                    for i, h in enumerate(headers):
                        ws3.cell(1, col + i, h)

                    data_last_row = 1
                    for i, (_, r) in enumerate(data.iterrows(), 2):
                        curr = r["电流值"]
                        flow = r["参考流量值"]
                        avg = r["平均值"]
                        pos_bar = r["正偏差bar"]
                        pos_pct = r["正偏差%"]
                        neg_bar = r["负偏差bar"]
                        neg_pct = r["负偏差%"]
                        lab = f"+{pos_pct}% / -{neg_pct}%"
                        ws3.cell(i, col + 0, curr)
                        ws3.cell(i, col + 1, flow)
                        ws3.cell(i, col + 2, avg)
                        ws3.cell(i, col + 3, pos_bar)
                        ws3.cell(i, col + 4, pos_pct)
                        ws3.cell(i, col + 5, neg_bar)
                        ws3.cell(i, col + 6, neg_pct)
                        ws3.cell(i, col + 7, lab)
                        data_last_row = i
                    ws3.cell(data_last_row + 1, col, "绘图数据结束")

                    ch = ScatterChart()
                    ch.scatterStyle = "lineMarker"
                    ch.title = f"一致性 {title}"
                    ch.x_axis.title = "流量/Lpm"
                    ch.y_axis.title = "压差/bar"
                    ch.x_axis.majorGridlines = None
                    ch.y_axis.majorGridlines = None

                    curr_list = []
                    for row in range(2, data_last_row + 1):
                        curr_val = ws3.cell(row, col + 0).value
                        if curr_val and curr_val not in curr_list:
                            curr_list.append(curr_val)
                    curr_list = sorted(curr_list)

                    for curr in curr_list:
                        sr = er = None
                        for row in range(2, data_last_row + 1):
                            if ws3.cell(row, col + 0).value == curr:
                                sr = row if sr is None else sr
                                er = row
                        if sr is None or er is None:
                            continue

                        x = Reference(ws3, col + 1, sr, col + 1, er)
                        y = Reference(ws3, col + 2, sr, col + 2, er)
                        label_ref = Reference(ws3, col + 7, sr, col + 7, er)

                        ser = Series(y, x, title=f"{curr}mA")
                        ser.marker = Marker(size=4)
                        ser.graphicalProperties.line = LineProperties(solidFill="00B0F0", w=8000)
                        ser.marker.graphicalProperties.solidFill = "00B0F0"
                        ser.dLbl = True
                        ser.labelRef = label_ref
                        ch.series.append(ser)
                    ws3.add_chart(ch, pos)

                if config["RUN_EXCEL_CHARTS"]:
                    draw_con_chart("前50%", "前50%", 18, "AQ2")
                    draw_con_chart("后50%", "后50%", 30, "AQ20")

                ws4 = wb.create_sheet("一致性影响分析(加权)")
                ws4.append([
                    "文件名",
                    "加权偏差总分",
                    "影响占比(%)",
                    "建议缩放系数P",
                    "0mA 60L/min 修正说明",
                    "调节建议"
                ])

                if file_score is not None:
                    for f, score in file_score.items():
                        rate = impact_rate.get(f, 0)
                        scale = scale_suggest.get(f, 1.0)
                        orig_p, new_p = scale_0mA60L.get(f, (0.0, 0.0))
                        correct_msg = f"{orig_p} → {new_p}"
                        note = f"全量程P值 × {scale}"
                        ws4.append([f, round(score, 2), rate, scale, correct_msg, note])

                excel_path = os.path.join(OUTPUT_FOLDER, excel_base_name)
                base, ext = os.path.splitext(excel_path)
                cnt = 1
                while os.path.exists(excel_path):
                    excel_path = f"{base}({cnt}){ext}"
                    cnt += 1
                wb.save(excel_path)
                write_log("✅ Excel检测报告生成完成\n")
            else:
                write_log("ℹ️ 已跳过：生成完整Excel检测报告\n")

            if check_stop_or_pause(log_file, write_log):
                log_file.close()
                return False

        if not resume:
            write_log("\n🎨 【11/12】正在生成迟滞、一致性、压差平均值图表...\n")
            if config["RUN_HYSTERESIS_PLOT"]:
                fig, ax = plt.subplots(figsize=(15, 6), dpi=100)
                COLOR_MAP = {0: '#70AD47', 1: '#00B0F0', 2: '#FFC000', 3: '#FF0000', 4: '#C00000'}
                MARKER_STYLES = ['o', 's', '^', 'D', 'v', 'p', '*', 'h', '8', 'P', 'X', '<', '>', 'd', 'H']
                unique_prods_list = list(plot_hys["数据源文件"].unique())
                for prod_idx, prod in enumerate(unique_prods_list):
                    sub = plot_hys[plot_hys["数据源文件"] == prod]
                    exceed_num = sum(sub["迟滞"] > config["STD_LIMIT"])
                    color_idx = min(exceed_num, 4)
                    color = COLOR_MAP[color_idx]
                    marker = MARKER_STYLES[prod_idx % len(MARKER_STYLES)]
                    # ---- 简化图例 ----
                    simplified_name = simplify_filename_for_legend(prod, config)
                    ax.plot(sub["电流值"], sub["迟滞"], marker=marker, ms=4, lw=1, color=color, label=simplified_name)
                if config["STD_LINE1"]["show"]:
                    ax.plot(plot_hys["电流值"].unique(), [10] * len(plot_hys["电流值"].unique()), 'r--', lw=1, label="标准线10")
                ax.set_title("P-Q滞环@20L/min")
                ax.set_xlabel("电流值")
                ax.set_ylabel("迟滞/bar")
                ax.grid(axis='y')
                unique_prods = plot_hys["数据源文件"].unique()
                ncol = min(3, len(unique_prods) // 15 + 1)
                if config["HYSTERESIS_SHOW_LEGEND"]:
                    ax.legend(loc="center left", bbox_to_anchor=(1.02, 0.5), ncol=ncol, fontsize=8, frameon=False)
                if config["HIDE_TOP_RIGHT_BORDER"]:
                    ax.spines['top'].set_visible(False)
                    ax.spines['right'].set_visible(False)
                plt.tight_layout()
                save_safe_plot(fig, "迟滞图.png", OUTPUT_FOLDER, dpi=config["PLOT_DPI"])
                write_log("✅ 迟滞图生成完成\n")
            else:
                write_log("ℹ️ 已跳过：迟滞图\n")

            if config["RUN_CONSISTENCY_PLOT"]:
                con_plot = df_con[df_con["数据分区"] == "前50%"].copy()
                currs = sorted(con_plot["电流值"].unique())
                fig, ax = plt.subplots(figsize=(15, 6), dpi=100)

                for idx, curr in enumerate(currs):
                    sub = con_plot[con_plot["电流值"] == curr].sort_values("参考流量值")
                    ax.plot(sub["参考流量值"], sub["平均值"], marker='o', ms=4, color='#00B0F0', label=f"{curr}mA")
                    x_last = sub["参考流量值"].iloc[-1]
                    y_last = sub["平均值"].iloc[-1]
                    ax.text(
                        x_last + config["PERF_TEXT_OFFSET_X"],
                        y_last + config["PERF_TEXT_OFFSET_Y"],
                        f"{curr}mA",
                        fontsize=config["CON_CURR_TEXT_SIZE"],
                        weight="bold" if config["CON_TITLE_BOLD"] else "normal",
                        va='center'
                    )

                    if config["CONSISTENCY_MODE"] == 0:
                        min_two_curr = currs[:2]
                        for x, y, dif, pct in zip(sub["参考流量值"], sub["平均值"], sub["正偏差bar"], sub["正偏差%"]):
                            if curr in min_two_curr and x in [5.0, 10.0]:
                                lab = f"±{dif:.2f}bar"
                            else:
                                lab = f"±{pct:.2f}%"
                            offset = config["LABEL_POS_0mA"] if curr == 0 else config["LABEL_POS_OTHER"]
                            ax.annotate(lab, (x, y), xytext=(0, offset), textcoords='offset points',
                                        ha='center', fontsize=config["CON_LABEL_SIZE"])
                    else:
                        for x, y, pb, pp, nb, np in zip(sub["参考流量值"], sub["平均值"], sub["正偏差bar"], sub["正偏差%"], sub["负偏差bar"], sub["负偏差%"]):
                            lab = f"+{pp}% / -{np}%"
                            offset = config["LABEL_POS_0mA"] if curr == 0 else config["LABEL_POS_OTHER"]
                            ax.annotate(lab, (x, y), xytext=(0, offset), textcoords='offset points',
                                        ha='center', fontsize=config["CON_LABEL_SIZE"])

                ax.set_title(config["CON_TITLE"], fontsize=config["CON_TITLE_SIZE"],
                             weight="bold" if config["CON_TITLE_BOLD"] else "normal",
                             x=config["CON_TITLE_OFFSET_X"], y=config["CON_TITLE_OFFSET_Y"])
                ax.set_xlabel("流量/Lpm", fontsize=config["CON_AXIS_LABEL_SIZE"],
                              weight="bold" if config["CON_AXIS_LABEL_BOLD"] else "normal")
                ax.set_ylabel("压差 bar", fontsize=config["CON_AXIS_LABEL_SIZE"],
                              weight="bold" if config["CON_AXIS_LABEL_BOLD"] else "normal")
                ax.tick_params(axis='both', labelsize=config["CON_TICK_LABEL_SIZE"], width=config["CON_TICK_WIDTH"])
                ax.set_xlim(right=ax.get_xlim()[1] + config["X_AXIS_EXTEND"])
                for line in ax.get_lines():
                    line.set_linewidth(config["CON_LINE_WIDTH"])
                if config["HIDE_TOP_RIGHT_BORDER"]:
                    ax.spines['top'].set_visible(False)
                    ax.spines['right'].set_visible(False)
                if config["CON_SHOW_LEGEND"]:
                    ax.legend(loc="upper right", frameon=False)
                plt.tight_layout()
                save_safe_plot(fig, "一致性图.png", OUTPUT_FOLDER, dpi=config["PLOT_DPI"])
                write_log("✅ 一致性图生成完成\n")
            else:
                write_log("ℹ️ 已跳过：一致性图\n")

            if config["RUN_AVG_PRESSURE_PLOT"]:
                con_plot_avg = df_con[df_con["数据分区"] == "前50%"].copy()
                currs_avg = sorted(con_plot_avg["电流值"].unique())

                fig_avg, ax_avg = plt.subplots(figsize=(15, 6), dpi=100)
                for idx, curr in enumerate(currs_avg):
                    sub_avg = con_plot_avg[con_plot_avg["电流值"] == curr].sort_values("参考流量值")
                    ax_avg.plot(sub_avg["参考流量值"], sub_avg["平均值"], marker='o', ms=4, color='#00B0F0', label=f"{curr}mA")
                    x_last_avg = sub_avg["参考流量值"].iloc[-1]
                    y_last_avg = sub_avg["平均值"].iloc[-1]
                    ax_avg.text(
                        x_last_avg + config["PERF_TEXT_OFFSET_X"],
                        y_last_avg + config["PERF_TEXT_OFFSET_Y"],
                        f"{curr}mA",
                        fontsize=config["CON_CURR_TEXT_SIZE"],
                        weight="bold" if config["CON_TITLE_BOLD"] else "normal",
                        va='center'
                    )

                    if config["CONSISTENCY_MODE"] == 0:
                        for x, y, avg_val, dif in zip(sub_avg["参考流量值"], sub_avg["平均值"], sub_avg["平均值"], sub_avg["正偏差bar"]):
                            lab_avg = f"{avg_val:.2f}±{dif:.2f} bar"
                            offset_avg = config["LABEL_POS_0mA"] if curr == 0 else config["LABEL_POS_OTHER"]
                            ax_avg.annotate(lab_avg, (x, y), xytext=(0, offset_avg), textcoords='offset points',
                                            ha='center', fontsize=config["CON_LABEL_SIZE"])
                    else:
                        for x, y, avg_val, pb, pp, nb, np in zip(sub_avg["参考流量值"], sub_avg["平均值"], sub_avg["平均值"], sub_avg["正偏差bar"], sub_avg["正偏差%"], sub_avg["负偏差bar"], sub_avg["负偏差%"]):
                            lab_avg = f"{avg_val:.1f} | +{pp}% / -{np}%"
                            offset_avg = config["LABEL_POS_0mA"] if curr == 0 else config["LABEL_POS_OTHER"]
                            ax_avg.annotate(lab_avg, (x, y), xytext=(0, offset_avg), textcoords='offset points',
                                            ha='center', fontsize=config["CON_LABEL_SIZE"])

                ax_avg.set_title("压差平均值", fontsize=config["CON_TITLE_SIZE"],
                                 weight="bold" if config["CON_TITLE_BOLD"] else "normal",
                                 x=config["CON_TITLE_OFFSET_X"], y=config["CON_TITLE_OFFSET_Y"])
                ax_avg.set_xlabel("流量/Lpm", fontsize=config["CON_AXIS_LABEL_SIZE"],
                                  weight="bold" if config["CON_AXIS_LABEL_BOLD"] else "normal")
                ax_avg.set_ylabel("压差 bar", fontsize=config["CON_AXIS_LABEL_SIZE"],
                                  weight="bold" if config["CON_AXIS_LABEL_BOLD"] else "normal")
                ax_avg.tick_params(axis='both', labelsize=config["CON_TICK_LABEL_SIZE"], width=config["CON_TICK_WIDTH"])
                ax_avg.set_xlim(right=ax_avg.get_xlim()[1] + config["X_AXIS_EXTEND"])
                for line in ax_avg.get_lines():
                    line.set_linewidth(config["CON_LINE_WIDTH"])
                if config["HIDE_TOP_RIGHT_BORDER"]:
                    ax_avg.spines['top'].set_visible(False)
                    ax_avg.spines['right'].set_visible(False)
                if config["AVG_PRESS_SHOW_LEGEND"]:
                    ax_avg.legend(loc="upper right", frameon=False)

                plt.tight_layout()
                save_safe_plot(fig_avg, "压差平均值图.png", OUTPUT_FOLDER, dpi=config["PLOT_DPI"])
                write_log("✅ 压差平均值图生成完成\n")
            else:
                write_log("ℹ️ 已跳过：压差平均值图\n")

            if check_stop_or_pause(log_file, write_log):
                log_file.close()
                return False

        if not resume:
            write_log("\n🔄 【12/12】正在生成去程、回程、合并PQ曲线图...\n")
            all_go_data = []
            all_back_data = []
            all_both_data = []
            LINE_COLOR = config["GOBACK_LINE_COLOR"]

            if config["RUN_SINGLE_GOBACK_PQ"]:
                folder_go = "01_去程PQ"
                folder_back = "02_回程PQ"
                folder_both = "03_去程+回程PQ"

                os.makedirs(os.path.join(OUTPUT_FOLDER, folder_go), exist_ok=True)
                os.makedirs(os.path.join(OUTPUT_FOLDER, folder_back), exist_ok=True)
                os.makedirs(os.path.join(OUTPUT_FOLDER, folder_both), exist_ok=True)

                for file_idx, file in enumerate(csv_files, 1):
                    if check_stop_or_pause(log_file, write_log):
                        log_file.close()
                        return False
                    
                    if file not in file_cleaned_data:
                        continue

                    df = file_cleaned_data[file].copy()
                    file_short = os.path.splitext(file)[0]
                    curr_list = sorted(df["电流值"].unique())
                    file_go = []
                    file_back = []

                    fig1, ax1 = plt.subplots(figsize=tuple(config["GOBACK_FIG_SIZE"]), dpi=config["GOBACK_DPI"])
                    for curr in curr_list:
                        sub = df[df["电流值"] == curr].copy()
                        sub_go = sub.iloc[:len(sub) // 2].sort_values("流量值")
                        x = sub_go["流量值"].abs()
                        y = sub_go["压差值"].abs()
                        if len(x) >= 2:
                            ax1.plot(x, y, lw=config["GOBACK_LINE_WIDTH"], color=LINE_COLOR)
                            ax1.text(x.iloc[-1] + config["PERF_TEXT_OFFSET_X"], y.iloc[-1] + config["PERF_TEXT_OFFSET_Y"],
                                     f"{round(curr)}mA", fontsize=config["PERF_CURR_TEXT_SIZE"], va="center")
                            file_go.append((x, y))
                    ax1.set_title(f"{file_short} 去程PQ", fontsize=config["GOBACK_TITLE_SIZE"],
                                  weight="bold" if config["GOBACK_TITLE_BOLD"] else "normal",
                                  x=config["GOBACK_TITLE_OFFSET_X"], y=config["GOBACK_TITLE_OFFSET_Y"])
                    ax1.set_xlabel("流量 L/min", fontsize=config["PERF_AXIS_LABEL_SIZE"])
                    ax1.set_ylabel("压差 bar", fontsize=config["PERF_AXIS_LABEL_SIZE"])
                    ax1.tick_params(labelsize=config["PERF_TICK_LABEL_SIZE"])
                    ax1.set_xlim(right=70)
                    if config["HIDE_TOP_RIGHT_BORDER"]:
                        ax1.spines['top'].set_visible(False)
                        ax1.spines['right'].set_visible(False)
                    if config["GOBACK_SHOW_LEGEND"]:
                        ax1.legend(frameon=False)
                    plt.tight_layout()
                    save_safe_plot(fig1, os.path.join(folder_go, f"去程PQ_{file_idx}.png"), OUTPUT_FOLDER, dpi=config["PLOT_DPI"])
                    all_go_data.append(file_go)

                    fig2, ax2 = plt.subplots(figsize=tuple(config["GOBACK_FIG_SIZE"]), dpi=config["GOBACK_DPI"])
                    for curr in curr_list:
                        sub = df[df["电流值"] == curr].copy()
                        sub_back = sub.iloc[len(sub) // 2:].sort_values("流量值")
                        x = sub_back["流量值"].abs()
                        y = sub_back["压差值"].abs()
                        if len(x) >= 2:
                            ax2.plot(x, y, lw=config["GOBACK_LINE_WIDTH"], color=LINE_COLOR)
                            ax2.text(x.iloc[-1] + config["PERF_TEXT_OFFSET_X"], y.iloc[-1] + config["PERF_TEXT_OFFSET_Y"],
                                     f"{round(curr)}mA", fontsize=config["PERF_CURR_TEXT_SIZE"], va="center")
                            file_back.append((x, y))
                    ax2.set_title(f"{file_short} 回程PQ", fontsize=config["GOBACK_TITLE_SIZE"],
                                  weight="bold" if config["GOBACK_TITLE_BOLD"] else "normal",
                                  x=config["GOBACK_TITLE_OFFSET_X"], y=config["GOBACK_TITLE_OFFSET_Y"])
                    ax2.set_xlabel("流量 L/min", fontsize=config["PERF_AXIS_LABEL_SIZE"])
                    ax2.set_ylabel("压差 bar", fontsize=config["PERF_AXIS_LABEL_SIZE"])
                    ax2.tick_params(labelsize=config["PERF_TICK_LABEL_SIZE"])
                    ax2.set_xlim(right=70)
                    if config["HIDE_TOP_RIGHT_BORDER"]:
                        ax2.spines['top'].set_visible(False)
                        ax2.spines['right'].set_visible(False)
                    if config["GOBACK_SHOW_LEGEND"]:
                        ax2.legend(frameon=False)
                    plt.tight_layout()
                    save_safe_plot(fig2, os.path.join(folder_back, f"回程PQ_{file_idx}.png"), OUTPUT_FOLDER, dpi=config["PLOT_DPI"])
                    all_back_data.append(file_back)

                    fig3, ax3 = plt.subplots(figsize=tuple(config["GOBACK_FIG_SIZE"]), dpi=config["GOBACK_DPI"])
                    for curr in curr_list:
                        sub = df[df["电流值"] == curr].copy()
                        sub_go = sub.iloc[:len(sub) // 2].sort_values("流量值")
                        xg = sub_go["流量值"].abs()
                        yg = sub_go["压差值"].abs()
                        if len(xg) >= 2:
                            ax3.plot(xg, yg, lw=config["GOBACK_LINE_WIDTH"], color=LINE_COLOR)

                        sub_back = sub.iloc[len(sub) // 2:].sort_values("流量值")
                        xb = sub_back["流量值"].abs()
                        yb = sub_back["压差值"].abs()
                        if len(xb) >= 2:
                            ax3.plot(xb, yb, lw=config["GOBACK_LINE_WIDTH"], color=LINE_COLOR)

                        if len(xg) >= 2:
                            ax3.text(xg.iloc[-1] + config["PERF_TEXT_OFFSET_X"], yg.iloc[-1] + config["PERF_TEXT_OFFSET_Y"],
                                     f"{round(curr)}mA", fontsize=config["PERF_CURR_TEXT_SIZE"], va="center")
                    ax3.set_title(f"{file_short} 去程+回程PQ", fontsize=config["GOBACK_TITLE_SIZE"],
                                  weight="bold" if config["GOBACK_TITLE_BOLD"] else "normal",
                                  x=config["GOBACK_TITLE_OFFSET_X"], y=config["GOBACK_TITLE_OFFSET_Y"])
                    ax3.set_xlabel("流量 L/min", fontsize=config["PERF_AXIS_LABEL_SIZE"])
                    ax3.set_ylabel("压差 bar", fontsize=config["PERF_AXIS_LABEL_SIZE"])
                    ax3.tick_params(labelsize=config["PERF_TICK_LABEL_SIZE"])
                    ax3.set_xlim(right=70)
                    if config["HIDE_TOP_RIGHT_BORDER"]:
                        ax3.spines['top'].set_visible(False)
                        ax3.spines['right'].set_visible(False)
                    if config["GOBACK_SHOW_LEGEND"]:
                        ax3.legend(frameon=False)
                    plt.tight_layout()
                    save_safe_plot(fig3, os.path.join(folder_both, f"去程回程对比_{file_idx}.png"), OUTPUT_FOLDER, dpi=config["PLOT_DPI"])
                    all_both_data.append((file_go, file_back))

                write_log("✅ 单文件去程/回程/合并PQ图生成完成\n")
            else:
                write_log("ℹ️ 已跳过：单文件去程/回程/合并PQ图\n")

            if check_stop_or_pause(log_file, write_log):
                log_file.close()
                return False

            if config["RUN_ALL_GOBACK_SUMMARY"] and config["RUN_SINGLE_GOBACK_PQ"]:
                write_log("\n📊 正在生成所有文件PQ总对比图...\n")
                folder_go = "01_去程PQ"
                folder_back = "02_回程PQ"
                folder_both = "03_去程+回程PQ"

                fig_go, ax_go = plt.subplots(
                    figsize=tuple(config["SUMMARY_FIG_SIZE"]), 
                    dpi=config["SUMMARY_DPI"]
                )
                for file_data in all_go_data:
                    for x, y in file_data:
                        ax_go.plot(x, y, color=LINE_COLOR, lw=config["SUMMARY_LINE_WIDTH"])
                ax_go.set_title("所有文件 去程PQ总对比", 
                    fontsize=config["SUMMARY_TITLE_SIZE"], 
                    pad=config["SUMMARY_PAD"])
                ax_go.set_xlabel("流量 L/min", fontsize=config["SUMMARY_AXIS_FONT_SIZE"])
                ax_go.set_ylabel("压差 bar", fontsize=config["SUMMARY_AXIS_FONT_SIZE"])
                ax_go.set_xlim(right=70)
                if config["HIDE_TOP_RIGHT_BORDER"]:
                    ax_go.spines['top'].set_visible(False)
                    ax_go.spines['right'].set_visible(False)
                if config["GOBACK_SUMMARY_SHOW_LEGEND"]:
                    ax_go.legend(frameon=False)
                plt.tight_layout()
                save_safe_plot(fig_go, os.path.join(folder_go, "所有文件_去程总对比.png"), 
                    OUTPUT_FOLDER, dpi=config["PLOT_DPI"])

                fig_back, ax_back = plt.subplots(
                    figsize=tuple(config["SUMMARY_FIG_SIZE"]), 
                    dpi=config["SUMMARY_DPI"]
                )
                for file_data in all_back_data:
                    for x, y in file_data:
                        ax_back.plot(x, y, color=LINE_COLOR, lw=config["SUMMARY_LINE_WIDTH"])
                ax_back.set_title("所有文件 回程PQ总对比", 
                    fontsize=config["SUMMARY_TITLE_SIZE"], 
                    pad=config["SUMMARY_PAD"])
                ax_back.set_xlabel("流量 L/min", fontsize=config["SUMMARY_AXIS_FONT_SIZE"])
                ax_back.set_ylabel("压差 bar", fontsize=config["SUMMARY_AXIS_FONT_SIZE"])
                ax_back.set_xlim(right=70)
                if config["HIDE_TOP_RIGHT_BORDER"]:
                    ax_back.spines['top'].set_visible(False)
                    ax_back.spines['right'].set_visible(False)
                if config["GOBACK_SUMMARY_SHOW_LEGEND"]:
                    ax_back.legend(frameon=False)
                plt.tight_layout()
                save_safe_plot(fig_back, os.path.join(folder_back, "所有文件_回程总对比.png"), 
                    OUTPUT_FOLDER, dpi=config["PLOT_DPI"])

                fig_both, ax_both = plt.subplots(
                    figsize=tuple(config["SUMMARY_FIG_SIZE"]), 
                    dpi=config["SUMMARY_DPI"]
                )
                for go_data, back_data in all_both_data:
                    for x, y in go_data:
                        ax_both.plot(x, y, color=LINE_COLOR, lw=config["SUMMARY_LINE_WIDTH"])
                    for x, y in back_data:
                        ax_both.plot(x, y, color=LINE_COLOR, lw=config["SUMMARY_LINE_WIDTH"])
                ax_both.set_title("所有文件 去程+回程PQ总对比", 
                    fontsize=config["SUMMARY_TITLE_SIZE"], 
                    pad=config["SUMMARY_PAD"])
                ax_both.set_xlabel("流量 L/min", fontsize=config["SUMMARY_AXIS_FONT_SIZE"])
                ax_both.set_ylabel("压差 bar", fontsize=config["SUMMARY_AXIS_FONT_SIZE"])
                ax_both.set_xlim(right=70)
                if config["HIDE_TOP_RIGHT_BORDER"]:
                    ax_both.spines['top'].set_visible(False)
                    ax_both.spines['right'].set_visible(False)
                if config["GOBACK_SUMMARY_SHOW_LEGEND"]:
                    ax_both.legend(frameon=False)
                plt.tight_layout()
                save_safe_plot(fig_both, os.path.join(folder_both, "所有文件_去程回程总对比.png"), OUTPUT_FOLDER, dpi=config["PLOT_DPI"])
                write_log("✅ 全体去程回程总对比图生成完成！\n")
            else:
                write_log("ℹ️ 已跳过：全体去程回程总对比图\n")

        write_log("\n" + "=" * 85 + "\n")
        write_log("📊 产品一致性检测最终结果（影响从大到小排序）\n")
        write_log("=" * 85 + "\n")

        if file_score is not None and len(file_score) > 0:
            write_log(f"{'文件名':<25} {'影响占比':<10} {'缩放系数':<10} {'0mA 60L/min 修正值（原始→新）'}\n")
            write_log("-" * 85 + "\n")
            for f, score in file_score.items():
                rate = f"{impact_rate.get(f, 0):.1f}%"
                scale = f"×{scale_suggest.get(f, 1.0):.2f}"
                orig, new = scale_0mA60L.get(f, (0.0, 0.0))
                corr = f"{orig:.2f} → {new:.2f}"
                write_log(f"{f:<25} {rate:<10} {scale:<10} {corr}\n")
        else:
            write_log("✅ 所有产品一致性均在合格范围内，无需调节\n")

        write_log(f"\n🎉 程序执行完毕！所有检测结果已保存\n")
        write_log(f"📂 结果保存路径：{OUTPUT_FOLDER}\n")

        log_file.close()
        return True
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        if config.get("PRINT_DEBUG_LOG", False):
            write_log(f"❌ 运行出错：{e}\n{error_msg}\n")
        else:
            write_log(f"❌ 运行出错：{e}\n")
        log_file.close()
        return False


# ============================================
# 6. ImageViewer 类 - 无红色标记，无累积误差
# ============================================
class ImageViewer:
    """
    v2.3 优化：支持滚轮缩放和拖拽平移的图片查看器
    主要改善：
    - LRU 缓存策略（固定容量 10 张），避免内存泄漏
    - 以鼠标位置为中心缩放，无累积误差
    - 防抖渲染，提升流畅度
    """
    CACHE_MAX_SIZE = 10
    PRELOAD_SCALES = [0.1, 0.2, 0.3, 0.5, 0.7, 1.0, 1.5, 2.0, 3.0, 5.0]
    
    def __init__(self, parent):
        self.parent = parent
        
        # 图片数据
        self.original_image = None
        self.image_path = None
        self.image_item_id = None
        self.info_text_id = None
        self.info_bg_id = None
        
        # 缩放参数
        self.scale = 1.0
        self.min_scale = 0.02
        self.max_scale = 30.0
        self.scale_step = 0.08
        
        # 拖拽参数
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.drag_start_offset_x = 0
        self.drag_start_offset_y = 0
        self.is_dragging = False
        
        # 图像偏移（图片左上角在画布上的位置）
        self.offset_x = 0.0
        self.offset_y = 0.0
        
        # 原图尺寸
        self.orig_width = 0
        self.orig_height = 0
        
        # 创建画布
        self.canvas = tk.Canvas(
            parent, 
            bg='#f5f6f8', 
            highlightthickness=0,
            cursor='hand2'
        )
        self.canvas.pack(fill='both', expand=True)
        
        # 绑定事件
        self.canvas.bind('<Configure>', self._on_configure)
        self.canvas.bind('<MouseWheel>', self._on_mousewheel)
        self.canvas.bind('<ButtonPress-1>', self._on_drag_start)
        self.canvas.bind('<B1-Motion>', self._on_drag_move)
        self.canvas.bind('<ButtonRelease-1>', self._on_drag_end)
        self.canvas.bind('<Double-Button-1>', self._on_double_click)
        
        # 渲染防抖
        self._render_timer = None
        
        # LRU 缓存：字典 + 访问顺序列表
        self._cache = {}
        self._cache_order = []
        self._is_preloaded = False
        
        # 当前显示图片引用（防止 GC）
        self._current_photo = None
        
        self.show_message("请选择图片查看\n\n💡 滚轮缩放 · 拖拽平移 · 双击复位")
    
    def _cache_get(self, scale_key):
        """从缓存获取，更新访问顺序（LRU）"""
        if scale_key in self._cache:
            self._cache_order.remove(scale_key)
            self._cache_order.append(scale_key)
            return self._cache[scale_key]
        return None
    
    def _cache_put(self, scale_key, photo):
        """放入缓存，超出容量时淘汰最久未使用的（LRU）"""
        if scale_key in self._cache:
            self._cache_order.remove(scale_key)
        elif len(self._cache) >= self.CACHE_MAX_SIZE:
            oldest = self._cache_order.pop(0)
            del self._cache[oldest]
        
        self._cache[scale_key] = photo
        self._cache_order.append(scale_key)
    
    def _cache_clear(self):
        """清空缓存"""
        self._cache.clear()
        self._cache_order.clear()
        self._is_preloaded = False
        
    def show_message(self, message):
        """显示提示信息"""
        self.canvas.delete('all')
        self.image_item_id = None
        self.info_text_id = None
        self.info_bg_id = None
        w = self.canvas.winfo_width() if self.canvas.winfo_width() > 10 else 300
        h = self.canvas.winfo_height() if self.canvas.winfo_height() > 10 else 200
        self.canvas.create_text(
            w // 2, h // 2,
            text=message,
            font=('Microsoft YaHei', 14),
            fill='#86909c',
            justify='center',
            tags='message'
        )
        
    def load_image(self, path):
        """加载图片"""
        try:
            self.image_path = path
            self.original_image = Image.open(path)
            if self.original_image.mode == 'RGBA':
                self.original_image = self.original_image.convert('RGB')
            
            self.orig_width, self.orig_height = self.original_image.size
            
            # 超大图限制尺寸，避免内存问题
            max_orig_size = 3000
            if self.orig_width > max_orig_size or self.orig_height > max_orig_size:
                ratio = min(max_orig_size / self.orig_width, max_orig_size / self.orig_height)
                new_w = int(self.orig_width * ratio)
                new_h = int(self.orig_height * ratio)
                self.original_image = self.original_image.resize(
                    (new_w, new_h), Image.Resampling.LANCZOS
                )
                self.orig_width, self.orig_height = self.original_image.size
            
            # 清空旧缓存
            self._cache_clear()
            
            self.scale = 1.0
            self.offset_x = 0.0
            self.offset_y = 0.0
            
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            
            if canvas_width < 10 or canvas_height < 10:
                canvas_width = 600
                canvas_height = 400
            
            scale_w = (canvas_width - 40) / self.orig_width
            scale_h = (canvas_height - 40) / self.orig_height
            self.scale = min(scale_w, scale_h, 1.0)
            
            display_w = int(self.orig_width * self.scale)
            display_h = int(self.orig_height * self.scale)
            self.offset_x = (canvas_width - display_w) // 2
            self.offset_y = (canvas_height - display_h) // 2
            
            self.canvas.delete('all')
            self.image_item_id = None
            self.info_text_id = None
            self.info_bg_id = None
            
            self._preload_cache_async()
            self._render_image(quality='high')
            
            return True
            
        except Exception as e:
            self.show_message(f"加载图片失败: {str(e)}")
            return False
    
    def _preload_cache_async(self):
        """异步预加载常用缩放级别"""
        if self._is_preloaded or self.original_image is None:
            return
        
        self._is_preloaded = True
        
        def preload_batch(scales):
            if not scales or self.original_image is None:
                return
            
            batch = scales[:3]
            remaining = scales[3:]
            
            for s in batch:
                if s < self.min_scale or s > self.max_scale:
                    continue
                scale_key = round(s, 2)
                if scale_key in self._cache:
                    continue
                
                try:
                    display_w = int(self.orig_width * s)
                    display_h = int(self.orig_height * s)
                    
                    if display_w < 1 or display_h < 1:
                        continue
                    
                    max_size = 2000
                    if display_w > max_size or display_h > max_size:
                        ratio = min(max_size / display_w, max_size / display_h)
                        display_w = int(display_w * ratio)
                        display_h = int(display_h * ratio)
                    
                    resized = self.original_image.resize(
                        (display_w, display_h), Image.Resampling.LANCZOS
                    )
                    photo = ImageTk.PhotoImage(resized)
                    self._cache_put(scale_key, photo)
                    
                except Exception:
                    pass
            
            if remaining:
                self.canvas.after(50, preload_batch, remaining)
        
        preload_batch(self.PRELOAD_SCALES.copy())
    
    def _get_display_size(self):
        """获取当前缩放后的显示尺寸"""
        display_w = int(self.orig_width * self.scale)
        display_h = int(self.orig_height * self.scale)
        
        if display_w < 1:
            display_w = 1
        if display_h < 1:
            display_h = 1
        
        return display_w, display_h
    
    def _render_image(self, quality='normal'):
        """渲染图片到画布"""
        if self.original_image is None:
            return
        
        display_w, display_h = self._get_display_size()
        scale_key = round(self.scale, 2)
        
        # 尝试从 LRU 缓存获取
        cached = self._cache_get(scale_key)
        if cached:
            photo = cached
        else:
            # 限制最大渲染尺寸
            max_render = 2000
            render_w, render_h = display_w, display_h
            if display_w > max_render or display_h > max_render:
                ratio = min(max_render / display_w, max_render / display_h)
                render_w = int(display_w * ratio)
                render_h = int(display_h * ratio)
            
            resample = Image.Resampling.LANCZOS if quality == 'high' else Image.Resampling.BILINEAR
            resized = self.original_image.resize((render_w, render_h), resample)
            photo = ImageTk.PhotoImage(resized)
            self._cache_put(scale_key, photo)
        
        # 更新画布
        center_x = self.offset_x + display_w // 2
        center_y = self.offset_y + display_h // 2
        
        if self.image_item_id:
            self.canvas.itemconfig(self.image_item_id, image=photo)
            self.canvas.coords(self.image_item_id, center_x, center_y)
        else:
            self.image_item_id = self.canvas.create_image(
                center_x, center_y, image=photo, anchor='center'
            )
        
        # 保持引用防止被 GC
        self._current_photo = photo
    
    def _on_configure(self, event):
        """画布尺寸变化"""
        if self.original_image:
            self._render_image()
    
    def _on_mousewheel(self, event):
        """鼠标滚轮缩放 - 以鼠标位置为中心"""
        if self.original_image is None:
            return "break"
        
        old_scale = self.scale
        if event.delta > 0:
            new_scale = min(self.scale * (1 + self.scale_step), self.max_scale)
        else:
            new_scale = max(self.scale / (1 + self.scale_step), self.min_scale)
        
        if new_scale == old_scale:
            # 已达缩放极限，仍消费事件防止冒泡到父级（如配置面板滚动）
            return "break"
        
        # 以鼠标位置为缩放中心（无累积误差）
        mouse_x = event.x - self.offset_x
        mouse_y = event.y - self.offset_y
        scale_ratio = new_scale / old_scale
        
        self.offset_x = event.x - mouse_x * scale_ratio
        self.offset_y = event.y - mouse_y * scale_ratio
        self.scale = new_scale
        
        # 防抖渲染
        if self._render_timer:
            self.canvas.after_cancel(self._render_timer)
        self._render_timer = self.canvas.after(10, self._render_image)
        
        return "break"
    
    def _on_drag_start(self, event):
        """开始拖拽"""
        if self.original_image is None:
            return
        self.is_dragging = True
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        self.drag_start_offset_x = self.offset_x
        self.drag_start_offset_y = self.offset_y
    
    def _on_drag_move(self, event):
        """拖拽中"""
        if not self.is_dragging or self.original_image is None:
            return
        
        self.offset_x = self.drag_start_offset_x + (event.x - self.drag_start_x)
        self.offset_y = self.drag_start_offset_y + (event.y - self.drag_start_y)
        
        if self.image_item_id:
            display_w, display_h = self._get_display_size()
            self.canvas.coords(
                self.image_item_id,
                self.offset_x + display_w // 2,
                self.offset_y + display_h // 2
            )
    
    def _on_drag_end(self, event):
        """结束拖拽"""
        self.is_dragging = False
    
    def _on_double_click(self, event):
        """双击复位"""
        if self.original_image is None:
            return
        
        canvas_width = max(self.canvas.winfo_width(), 10)
        canvas_height = max(self.canvas.winfo_height(), 10)
        
        scale_w = (canvas_width - 40) / self.orig_width
        scale_h = (canvas_height - 40) / self.orig_height
        self.scale = min(scale_w, scale_h, 1.0)
        
        display_w, display_h = self._get_display_size()
        self.offset_x = (canvas_width - display_w) // 2
        self.offset_y = (canvas_height - display_h) // 2
        
        self._render_image(quality='high')


class CollapsiblePanel:
    def __init__(self, parent, title, *args, **kwargs):
        self.frame = tk.Frame(parent, *args, **kwargs)
        self.frame.configure(bg='#ffffff', relief='flat', bd=0)
        
        self.canvas = tk.Canvas(self.frame, highlightthickness=0, bg='#ffffff', bd=0)
        self.canvas.pack(fill='both', expand=True)
        
        self.toggle_btn = ttk.Button(self.canvas, text=f"▶ {title}", 
                                     command=self.toggle, 
                                     style='Collapse.TButton')
        self.toggle_btn.pack(fill='x', pady=(6, 2), padx=6)
        
        self.content = tk.Frame(self.canvas, bg='#ffffff')
        self.is_expanded = False
        self.row = 0
        
        self.bg_rect_id = None
        
        self.canvas.bind('<Configure>', self._on_configure)
        self.toggle_btn.bind('<ButtonRelease-1>', lambda e: self._on_configure(None))
        
    def _on_configure(self, event):
        self.frame.after(50, self._draw_rounded_rect)
        
    def _draw_rounded_rect(self):
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w <= 10 or h <= 10:
            return
        
        if self.bg_rect_id:
            self.canvas.delete(self.bg_rect_id)
        
        radius = 10
        points = [
            radius, 2,
            w - radius, 2,
            w - 2, 2,
            w - 2, radius,
            w - 2, h - radius,
            w - 2, h - 2,
            w - radius, h - 2,
            radius, h - 2,
            2, h - 2,
            2, h - radius,
            2, radius,
            2, 2
        ]
        
        self.bg_rect_id = self.canvas.create_polygon(
            points,
            outline='#e5e6eb',
            fill='#ffffff',
            width=1,
            smooth=True
        )
        self.canvas.tag_lower(self.bg_rect_id)
        
    def toggle(self):
        if self.is_expanded:
            self.content.pack_forget()
            current_text = self.toggle_btn.cget('text')
            if current_text.startswith('▼ '):
                self.toggle_btn.config(text=f"▶ {current_text[2:]}")
            self.is_expanded = False
        else:
            self.content.pack(fill='x', padx=6, pady=6)
            current_text = self.toggle_btn.cget('text')
            if current_text.startswith('▶ '):
                self.toggle_btn.config(text=f"▼ {current_text[2:]}")
            self.is_expanded = True
        self.frame.after(100, self._draw_rounded_rect)

    def add_row(self, label_text, widget):
        if isinstance(widget, (tk.Checkbutton, ttk.Checkbutton)):
            lbl = ttk.Label(self.content, 
                           text=label_text, 
                           anchor='w', 
                           wraplength=250,
                           foreground='#1d2129',
                           background='#ffffff')
            lbl.grid(row=self.row, column=0, sticky='w', padx=(0, 10), pady=6)
            widget.grid(row=self.row, column=1, sticky='w', pady=6)
            self.content.columnconfigure(0, weight=1)
            self.content.columnconfigure(1, weight=0)
            self.row += 1
        else:
            lbl = ttk.Label(
                self.content,
                text=label_text,
                wraplength=300,
                anchor='w',
                justify='left',
                foreground='#1d2129',
                background='#ffffff'
            )
            lbl.grid(row=self.row, column=0, columnspan=2, sticky='w', padx=2, pady=(4, 0))
            self.row += 1
            widget.grid(row=self.row, column=0, columnspan=2, sticky='ew', padx=2, pady=(0, 6))
            self.content.columnconfigure(0, weight=1)
            self.content.columnconfigure(1, weight=0)
            self.row += 1


class DirectoryPickerDialog:
    """
    v2.3 新增：自定义目录选择对话框
    可以看到目录中的文件（如 csv、xlsx 等），但只能选择目录
    左侧目录树，右侧文件列表，双击目录进入
    """
    
    def __init__(self, parent, initial_dir=None):
        self.result = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("选择工作目录")
        self.dialog.geometry("700x500")
        self.dialog.minsize(600, 400)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # 居中显示
        self.dialog.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - 700) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - 500) // 2
        self.dialog.geometry(f"+{max(x, 100)}+{max(y, 100)}")
        
        # 当前路径显示
        path_frame = ttk.Frame(self.dialog, padding=(10, 10, 10, 5))
        path_frame.pack(fill='x')
        
        ttk.Label(path_frame, text="当前路径:").pack(side='left')
        self.path_var = tk.StringVar(value=initial_dir or os.getcwd())
        path_entry = ttk.Entry(path_frame, textvariable=self.path_var)
        path_entry.pack(side='left', fill='x', expand=True, padx=(5, 5))
        ttk.Button(path_frame, text="跳转", command=self._jump_to_path, width=6).pack(side='left')
        
        # 主体：PanedWindow 分隔目录树和文件列表
        main_paned = ttk.PanedWindow(self.dialog, orient='horizontal')
        main_paned.pack(fill='both', expand=True, padx=10, pady=5)
        
        # 左侧：目录树
        tree_frame = ttk.LabelFrame(main_paned, text="目录", padding=5)
        main_paned.add(tree_frame, weight=1)
        
        self.dir_tree = ttk.Treeview(tree_frame, columns=('size',), show='tree headings', selectmode='browse')
        self.dir_tree.heading('#0', text='名称')
        self.dir_tree.heading('size', text='大小')
        self.dir_tree.column('size', width=60, anchor='e', stretch=False)
        
        tree_scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=self.dir_tree.yview)
        self.dir_tree.configure(yscrollcommand=tree_scroll.set)
        
        self.dir_tree.pack(side='left', fill='both', expand=True)
        tree_scroll.pack(side='right', fill='y')
        
        self.dir_tree.bind('<<TreeviewSelect>>', self._on_dir_select)
        self.dir_tree.bind('<Double-1>', self._on_dir_double_click)
        
        # 右侧：文件列表
        file_frame = ttk.LabelFrame(main_paned, text="文件（仅浏览，不可选择）", padding=5)
        main_paned.add(file_frame, weight=2)
        
        self.file_tree = ttk.Treeview(file_frame, columns=('size', 'type'), show='headings', selectmode='none')
        self.file_tree.heading('name', text='名称')
        self.file_tree.heading('size', text='大小')
        self.file_tree.heading('type', text='类型')
        self.file_tree.column('name', width=200, anchor='w')
        self.file_tree.column('size', width=70, anchor='e')
        self.file_tree.column('type', width=80, anchor='center')
        
        # 用第一列显示文件名（show='headings' 模式下没有 #0 列）
        self.file_tree['columns'] = ('name', 'size', 'type')
        self.file_tree.heading('name', text='名称')
        self.file_tree.column('name', width=200, anchor='w')
        
        file_scroll = ttk.Scrollbar(file_frame, orient='vertical', command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=file_scroll.set)
        
        self.file_tree.pack(side='left', fill='both', expand=True)
        file_scroll.pack(side='right', fill='y')
        
        # 底部按钮
        btn_frame = ttk.Frame(self.dialog, padding=(10, 5, 10, 10))
        btn_frame.pack(fill='x')
        
        ttk.Label(btn_frame, text="💡 双击目录进入，点击选择按钮确认当前目录", 
                  font=('Microsoft YaHei', 9), foreground='#86909c').pack(side='left')
        
        ttk.Button(btn_frame, text="取消", command=self._cancel, width=10).pack(side='right', padx=(5, 0))
        ttk.Button(btn_frame, text="选择", command=self._confirm, width=10).pack(side='right')
        
        # 填充目录树
        self._populate_tree(initial_dir or os.getcwd())
        
        # 等待对话框关闭
        self.dialog.wait_window()
    
    def _populate_tree(self, start_path):
        """填充目录树"""
        self.dir_tree.delete(*self.dir_tree.get_children())
        
        # 添加根节点（我的电脑/此电脑风格）
        # 先从当前目录往上加到根
        current = os.path.abspath(start_path)
        path_parts = []
        while True:
            path_parts.insert(0, current)
            parent = os.path.dirname(current)
            if parent == current:  # 到达根目录
                break
            current = parent
        
        # 在 Windows 上，盘符是根
        # 直接从当前目录开始显示，支持向上导航
        self._add_dir_node('', start_path)
        self._load_files(start_path)
    
    def _add_dir_node(self, parent, dir_path):
        """递归添加目录节点"""
        try:
            items = sorted(os.listdir(dir_path))
        except (PermissionError, OSError):
            return
        
        dirs = []
        for item in items:
            full = os.path.join(dir_path, item)
            if os.path.isdir(full) and not item.startswith('.'):
                dirs.append(item)
        
        for d in dirs:
            full_path = os.path.join(dir_path, d)
            node = self.dir_tree.insert(parent, 'end', text=f"📁 {d}", values=('',), open=False)
            # 标记为有子节点（延迟加载）
            self.dir_tree.insert(node, 'end', text='', values=('',))  # 占位节点
            # 存储完整路径
            self.dir_tree.set(node, 'size', full_path)
        
        # 如果是根节点（parent 为空），添加"上级目录"
        if not parent:
            parent_dir = os.path.dirname(dir_path)
            if parent_dir and parent_dir != dir_path:
                self.dir_tree.insert('', 0, text='.. (上级目录)', values=(parent_dir,), open=False)
    
    def _on_dir_select(self, event):
        """选中目录时加载文件列表"""
        selection = self.dir_tree.selection()
        if not selection:
            return
        
        node = selection[0]
        item_text = self.dir_tree.item(node, 'text')
        
        # 上级目录
        if item_text.startswith('..'):
            return
        
        # 获取路径
        values = self.dir_tree.item(node, 'values')
        if values and values[0]:
            dir_path = values[0]
            self.path_var.set(dir_path)
            self._load_files(dir_path)
    
    def _on_dir_double_click(self, event):
        """双击目录进入"""
        selection = self.dir_tree.selection()
        if not selection:
            return
        
        node = selection[0]
        item_text = self.dir_tree.item(node, 'text')
        
        # 上级目录
        if item_text.startswith('..'):
            values = self.dir_tree.item(node, 'values')
            if values and values[0]:
                self._navigate_to(values[0])
            return
        
        # 检查是否已加载子节点
        children = self.dir_tree.get_children(node)
        if len(children) == 1 and self.dir_tree.item(children[0], 'text') == '':
            # 占位节点，需要真实加载
            self.dir_tree.delete(children[0])
            values = self.dir_tree.item(node, 'values')
            if values and values[0]:
                self._add_dir_node(node, values[0])
        
        # 展开节点
        self.dir_tree.item(node, open=True)
        
        # 更新路径和文件列表
        values = self.dir_tree.item(node, 'values')
        if values and values[0]:
            dir_path = values[0]
            self.path_var.set(dir_path)
            self._load_files(dir_path)
    
    def _navigate_to(self, dir_path):
        """导航到指定目录"""
        self.dir_tree.delete(*self.dir_tree.get_children())
        self._add_dir_node('', dir_path)
        self.path_var.set(dir_path)
        self._load_files(dir_path)
    
    def _jump_to_path(self):
        """跳转到输入框指定的路径"""
        path = self.path_var.get().strip()
        if os.path.isdir(path):
            self._navigate_to(path)
        else:
            messagebox.showwarning("提示", "路径不存在或不是目录", parent=self.dialog)
    
    def _load_files(self, dir_path):
        """加载右侧文件列表"""
        self.file_tree.delete(*self.file_tree.get_children())
        
        try:
            items = sorted(os.listdir(dir_path))
        except (PermissionError, OSError):
            return
        
        for item in items:
            full = os.path.join(dir_path, item)
            if os.path.isfile(full):
                try:
                    size = os.path.getsize(full)
                    size_str = self._format_size(size)
                except OSError:
                    size_str = '-'
                
                ext = os.path.splitext(item)[1].upper()
                if ext:
                    ext = ext[1:]  # 去掉点
                else:
                    ext = '文件'
                
                self.file_tree.insert('', 'end', values=(item, size_str, ext))
    
    def _format_size(self, size):
        """格式化文件大小"""
        if size < 1024:
            return f"{size} B"
        elif size < 1024 * 1024:
            return f"{size/1024:.1f} KB"
        elif size < 1024 * 1024 * 1024:
            return f"{size/(1024*1024):.1f} MB"
        else:
            return f"{size/(1024*1024*1024):.1f} GB"
    
    def _confirm(self):
        """确认选择"""
        path = self.path_var.get().strip()
        if os.path.isdir(path):
            self.result = path
            self.dialog.destroy()
        else:
            messagebox.showwarning("提示", "请选择一个有效的目录", parent=self.dialog)
    
    def _cancel(self):
        """取消"""
        self.result = None
        self.dialog.destroy()


class ConfigApp:
    def __init__(self, root):
        self.root = root
        root.title("PQ产品一致性检测程序 - 配置面板")
        # 获取屏幕尺寸
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        # 设置窗口尺寸为屏幕的85%（留出任务栏空间）
        window_width = int(screen_width * 0.75)
        window_height = int(screen_height * 0.75)

        # 计算居中位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        root.configure(bg='#f5f6f8')

        style = ttk.Style(root)
        style.theme_use('clam')

        BG_MAIN = '#f5f6f8'
        BG_CARD = '#ffffff'
        BG_HOVER = '#f0f2f5'
        TEXT_PRIMARY = '#1d2129'
        TEXT_SECONDARY = '#86909c'
        TEXT_WHITE = '#ffffff'
        BLUE_PRIMARY = '#3370ff'
        BLUE_HOVER = '#2a5fd9'
        BORDER_LIGHT = '#e5e6eb'

        FONT = ('Microsoft YaHei', 9)
        FONT_BOLD = ('Microsoft YaHei', 9, 'bold')
        FONT_TITLE = ('Microsoft YaHei', 11, 'bold')

        style.configure('.', font=FONT, background=BG_MAIN, foreground=TEXT_PRIMARY)

        style.configure('Card.TLabelframe',
                        background=BG_CARD,
                        bordercolor=BORDER_LIGHT,
                        relief='flat',
                        borderwidth=1)
        style.configure('Card.TLabelframe.Label',
                        font=FONT_TITLE,
                        foreground=TEXT_PRIMARY,
                        background=BG_CARD,
                        padding=(0, 5, 0, 0))

        style.configure('TButton',
                        font=FONT_BOLD,
                        padding=(16, 8),
                        relief='flat',
                        borderwidth=0,
                        background=BLUE_PRIMARY,
                        foreground=TEXT_WHITE,
                        focuscolor='none')
        style.map('TButton',
                  background=[('active', BLUE_HOVER), ('!active', BLUE_PRIMARY)],
                  foreground=[('active', TEXT_WHITE), ('!active', TEXT_WHITE)])

        style.configure('Secondary.TButton',
                        font=FONT_BOLD,
                        padding=(16, 8),
                        relief='flat',
                        borderwidth=0,
                        background=BG_HOVER,
                        foreground=TEXT_PRIMARY,
                        focuscolor='none')
        style.map('Secondary.TButton',
                  background=[('active', '#e5e6eb'), ('!active', BG_HOVER)],
                  foreground=[('active', TEXT_PRIMARY), ('!active', TEXT_PRIMARY)])

        style.configure('Collapse.TButton',
                        font=FONT_BOLD,
                        padding=(12, 8),
                        anchor='w',
                        relief='flat',
                        borderwidth=0,
                        background=BG_CARD,
                        foreground=TEXT_PRIMARY)
        style.map('Collapse.TButton',
                  background=[('active', BG_HOVER), ('!active', BG_CARD)],
                  foreground=[('active', BLUE_PRIMARY), ('!active', TEXT_PRIMARY)])

        style.configure('TEntry',
                        fieldbackground=BG_CARD,
                        foreground=TEXT_PRIMARY,
                        borderwidth=1,
                        relief='solid',
                        padding=6)
        style.map('TEntry',
                  bordercolor=[('focus', BLUE_PRIMARY), ('!focus', BORDER_LIGHT)],
                  fieldbackground=[('focus', BG_CARD), ('!focus', BG_CARD)])

        style.configure('TCheckbutton',
                        background=BG_MAIN,
                        font=FONT,
                        foreground=TEXT_PRIMARY,
                        focuscolor='none')
        style.map('TCheckbutton',
                  background=[('active', BG_MAIN), ('!active', BG_MAIN)],
                  foreground=[('active', TEXT_PRIMARY), ('!active', TEXT_PRIMARY)],
                  indicatorcolor=[('selected', BLUE_PRIMARY), ('!selected', BG_CARD)])

        style.configure('Vertical.TScrollbar',
                        background=BG_HOVER,
                        troughcolor=BG_MAIN,
                        arrowcolor='#86909c',
                        borderwidth=0,
                        relief='flat')
        style.map('Vertical.TScrollbar',
                  background=[('active', '#d0d2d6'), ('!active', BG_HOVER)])

        main_frame = ttk.Frame(root, padding=20)
        main_frame.pack(fill='both', expand=True)

        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=0)
        main_frame.rowconfigure(1, weight=1)

        top_frame = ttk.Frame(main_frame)
        top_frame.grid(row=0, column=0, sticky='ew', pady=(0, 15))

        dir_frame = ttk.Frame(top_frame)
        dir_frame.pack(fill='x')

        ttk.Label(dir_frame, text="📁 工作目录:", font=('Microsoft YaHei', 10)).pack(side='left', padx=(0, 10))

        script_dir = get_script_directory()
        self.dir_var = tk.StringVar(value=script_dir)
        dir_entry = ttk.Entry(dir_frame, textvariable=self.dir_var, width=60)
        dir_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))

        ttk.Button(dir_frame, text="浏览...", command=self.select_directory, width=10).pack(side='left')
        ttk.Button(dir_frame, text="切换到脚本所在目录", command=self.reset_to_script_dir, width=18).pack(side='left', padx=(5, 0))

        # v2.3 改善：使用 PanedWindow 实现可拖拽调整列宽
        main_paned = ttk.PanedWindow(main_frame, orient='horizontal')
        main_paned.grid(row=1, column=0, sticky='nsew')

        # 左列：配置参数
        config_frame = ttk.LabelFrame(main_paned, text="⚙️ 配置参数", style='Card.TLabelframe', padding=15)
        main_paned.add(config_frame, weight=1)

        canvas_frame = ttk.Frame(config_frame)
        canvas_frame.pack(fill='both', expand=True)

        self.canvas = tk.Canvas(canvas_frame, highlightthickness=0, bg=BG_MAIN, bd=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient='vertical', command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=BG_MAIN)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor='nw')
        self.canvas.configure(yscrollcommand=scrollbar.set)

        def _configure_canvas(event):
            self.canvas.itemconfig(self.canvas_window, width=event.width - 4)
        self.canvas.bind('<Configure>', _configure_canvas)

        self.canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        def on_mousewheel(event):
            # v2.3 修复：event.widget 可能是字符串（widget 路径名）而非对象
            widget = event.widget
            if isinstance(widget, str):
                try:
                    widget = self.root.nametowidget(widget)
                except (KeyError, tk.TclError):
                    return
            
            # v2.3 修复：如果鼠标在图片预览区（ImageViewer），不滚动配置面板
            # 避免滚轮缩放图片时，配置面板也跟着滚动
            try:
                w = widget
                while w is not None:
                    if hasattr(self, 'image_viewer') and w is self.image_viewer.canvas:
                        return
                    w = getattr(w, 'master', None)
            except Exception:
                pass
            
            # 向上查找父级，判断鼠标是否在滚动区域内
            try:
                w = widget
                while w is not None:
                    if w in (self.canvas, self.scrollable_frame):
                        break
                    w = getattr(w, 'master', None)
                else:
                    return
            except Exception:
                return

            if hasattr(event, 'delta') and event.delta:
                delta = -1 if event.delta > 0 else 1
            elif hasattr(event, 'num') and event.num in (4, 5):
                delta = -1 if event.num == 4 else 1 if event.num == 5 else 0
            else:
                return

            step = 0.03 * delta
            first, last = self.canvas.yview()
            new_first = max(0.0, min(1.0 - (last - first), first + step))
            self.canvas.yview_moveto(new_first)

        self.root.bind_all("<MouseWheel>", on_mousewheel)

        self.widgets = {}

        for group_name, flat_keys in GROUPS:
            panel = CollapsiblePanel(self.scrollable_frame, group_name)
            panel.frame.pack(fill='x', pady=(0, 8))
            for fk in flat_keys:
                default = FLAT_DEFAULTS.get(fk, "")
                display_name = PARAM_CN_NAME.get(fk, fk)
                if isinstance(default, bool):
                    var = tk.BooleanVar(value=default)
                    cb = Checkbutton(panel.content,
                                     variable=var,
                                     font=FONT,
                                     bg='#ffffff',
                                     fg='#1d2129',
                                     activebackground='#ffffff',
                                     activeforeground='#1d2129',
                                     selectcolor='#ffffff',
                                     relief='flat',
                                     bd=0)
                    panel.add_row(display_name, cb)
                    self.widgets[fk] = var
                else:
                    entry = ttk.Entry(panel.content, width=40)
                    entry.insert(0, str(default))
                    panel.add_row(display_name, entry)
                    self.widgets[fk] = entry
            panel.content.pack_forget()

        # ===== 日志 + 图片预览区域 =====
        # 中列：运行日志 & 图片预览
        log_frame = ttk.LabelFrame(main_paned, text="📋 运行日志 & 图片预览", style='Card.TLabelframe', padding=10)
        main_paned.add(log_frame, weight=2)

        self.notebook = ttk.Notebook(log_frame)
        self.notebook.pack(fill='both', expand=True)

        log_tab = ttk.Frame(self.notebook)
        self.notebook.add(log_tab, text="📋 日志")

        image_tab = ttk.Frame(self.notebook)
        self.notebook.add(image_tab, text="🖼️ 图片预览")

        self.log_text = scrolledtext.ScrolledText(
            log_tab,
            height=25,
            state='disabled',
            wrap='word',
            font=('Consolas', 10),
            bg=BG_CARD,
            fg=TEXT_PRIMARY,
            insertbackground=TEXT_PRIMARY,
            relief='flat',
            borderwidth=0,
            padx=10,
            pady=10
        )
        self.log_text.pack(fill='both', expand=True)

        # 图片预览区域
        image_control_frame = ttk.Frame(image_tab)
        image_control_frame.pack(fill='x', padx=10, pady=10)
        
        # 使用 grid 布局实现响应式收缩
        # 列 0: 标签, 列 1: 下拉框(可伸缩, 有最小宽度), 列 2: 复位按钮, 列 3: 提示文字
        image_control_frame.columnconfigure(1, weight=1, minsize=150)  # 下拉框优先伸缩，最小150px
        
        ttk.Label(image_control_frame, text="选择图片:").grid(row=0, column=0, sticky='w', padx=(0, 10))

        self.image_var = tk.StringVar()
        self.image_combo = ttk.Combobox(image_control_frame, textvariable=self.image_var, state='readonly', width=20)
        self.image_combo.grid(row=0, column=1, sticky='ew', padx=(0, 10))
        self.image_combo.bind('<<ComboboxSelected>>', self.on_image_select)

        # v2.3 改善：删除刷新按钮，打开文件夹移到运行控制区，只保留复位按钮
        ttk.Button(image_control_frame, text="🔍 复位", command=self.reset_image_view).grid(row=0, column=2, sticky='ew')

        tip_label = ttk.Label(
            image_control_frame, 
            text="💡 滚轮缩放 · 拖拽平移 · 双击复位",
            font=('Microsoft YaHei', 9),
            foreground='#86909c'
        )
        tip_label.grid(row=0, column=3, sticky='e', padx=(10, 0))

        # 使用 ImageViewer
        image_display_container = ttk.Frame(image_tab)
        image_display_container.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        self.image_viewer = ImageViewer(image_display_container)

        self.output_folder = None
        self.last_viewed_image = None
        self.memory = load_memory()
        self.last_viewed_image = self.memory.get('last_image', None)

        # 右列：运行控制
        control_frame = ttk.Frame(main_paned)
        main_paned.add(control_frame, weight=1)

        control_card = ttk.LabelFrame(control_frame, text="🎯 运行控制", style='Card.TLabelframe', padding=15)
        control_card.pack(fill='both', expand=True)

        btn_frame = ttk.Frame(control_card)
        btn_frame.pack(fill='x', pady=5)

        self.run_btn = ttk.Button(btn_frame, text="🚀 运行分析", command=self.run_analysis, width=16)
        self.run_btn.pack(pady=(0, 8), fill='x')

        self.pause_btn = ttk.Button(btn_frame, text="⏸ 暂停", command=self.pause_analysis, width=16, style='Secondary.TButton')
        self.pause_btn.pack(pady=(0, 8), fill='x')
        self.pause_btn.config(state='disabled')

        self.resume_btn = ttk.Button(btn_frame, text="▶ 恢复", command=self.resume_analysis, width=16, style='Secondary.TButton')
        self.resume_btn.pack(pady=(0, 8), fill='x')
        self.resume_btn.config(state='disabled')

        self.stop_btn = ttk.Button(btn_frame, text="⏹ 停止", command=self.stop_analysis, width=16, style='Secondary.TButton')
        self.stop_btn.pack(pady=(0, 8), fill='x')
        self.stop_btn.config(state='disabled')

        # v2.3 改善：打开文件夹按钮移到运行控制区
        ttk.Button(control_card, text="📂 打开输出文件夹", command=self.open_image_folder, style='Secondary.TButton').pack(fill='x', pady=(8, 8))

        ttk.Button(control_card, text="↻ 恢复默认", command=self.reset_defaults, width=16, style='Secondary.TButton').pack(fill='x')

        self.running = False
        self.paused = False
        run_analysis.root = self.root

    def select_directory(self):
        dir_path = filedialog.askdirectory(
            title="选择工作目录",
            initialdir=self.dir_var.get()
        )
        if dir_path:
            self.dir_var.set(dir_path)
            self.log(f"📁 已选择目录: {dir_path}\n")

    def reset_to_script_dir(self):
        script_dir = get_script_directory()
        self.dir_var.set(script_dir)
        self.log(f"📁 已切换到脚本所在目录: {script_dir}\n")

    def change_working_directory(self):
        target_dir = self.dir_var.get().strip()
        if not target_dir:
            target_dir = get_script_directory()
        
        if not os.path.exists(target_dir):
            self.log(f"❌ 目录不存在: {target_dir}\n")
            return False
        
        try:
            os.chdir(target_dir)
            self.log(f"✅ 已切换工作目录到: {os.getcwd()}\n")
            return True
        except Exception as e:
            self.log(f"❌ 切换目录失败: {e}\n")
            return False

    def log(self, msg):
        self.log_text.config(state='normal')
        self.log_text.insert('end', msg)
        self.log_text.see('end')
        self.log_text.config(state='disabled')
        self.log_text.update_idletasks()

    def reset_defaults(self):
        for fk, widget in self.widgets.items():
            default = FLAT_DEFAULTS.get(fk)
            if default is None:
                continue
            if isinstance(widget, tk.BooleanVar):
                widget.set(default)
            elif isinstance(widget, ttk.Spinbox):
                widget.delete(0, 'end')
                widget.insert(0, str(default))
            elif isinstance(widget, ttk.Entry):
                widget.delete(0, 'end')
                widget.insert(0, str(default))

    def get_config(self):
        flat = {}
        for fk, widget in self.widgets.items():
            if isinstance(widget, tk.BooleanVar):
                val = widget.get()
            elif isinstance(widget, ttk.Spinbox):
                try:
                    val = int(widget.get())
                except:
                    val = 0
            else:
                val_str = widget.get()
                default = FLAT_DEFAULTS.get(fk)
                if isinstance(default, bool):
                    val = val_str.lower() in ('true', '1', 'yes')
                elif isinstance(default, int):
                    try:
                        val = int(val_str)
                    except:
                        val = 0
                elif isinstance(default, float):
                    try:
                        val = float(val_str)
                    except:
                        val = 0.0
                else:
                    val = val_str
            flat[fk] = val

        config = {}
        for fk, val in flat.items():
            if "__" in fk:
                parent, child = fk.split("__", 1)
                if parent not in config:
                    config[parent] = {}
                config[parent][child] = val
            else:
                config[fk] = val

        if "REF_FLOW" in flat:
            try:
                config["REF_FLOW"] = [float(x.strip()) for x in flat["REF_FLOW"].split(",") if x.strip()]
            except:
                config["REF_FLOW"] = DEFAULT_CONFIG["REF_FLOW"]
        if "KEY_CURRENTS" in flat:
            try:
                config["KEY_CURRENTS"] = [float(x.strip()) for x in flat["KEY_CURRENTS"].split(",") if x.strip()]
            except:
                config["KEY_CURRENTS"] = DEFAULT_CONFIG["KEY_CURRENTS"]
        if "GOBACK_FIG_SIZE" in flat:
            try:
                config["GOBACK_FIG_SIZE"] = [float(x.strip()) for x in flat["GOBACK_FIG_SIZE"].split(",") if x.strip()]
            except:
                config["GOBACK_FIG_SIZE"] = DEFAULT_CONFIG["GOBACK_FIG_SIZE"]
        if "SUMMARY_FIG_SIZE" in flat:
            try:
                config["SUMMARY_FIG_SIZE"] = [float(x.strip()) for x in flat["SUMMARY_FIG_SIZE"].split(",") if x.strip()]
            except:
                config["SUMMARY_FIG_SIZE"] = DEFAULT_CONFIG["SUMMARY_FIG_SIZE"]

        if "STD_LINE1" in config:
            if "value" not in config["STD_LINE1"]:
                config["STD_LINE1"]["value"] = DEFAULT_CONFIG["STD_LINE1"]["value"]
            if "show" not in config["STD_LINE1"]:
                config["STD_LINE1"]["show"] = DEFAULT_CONFIG["STD_LINE1"]["show"]

        return config

    # ===== 图片预览相关方法 =====
    def refresh_image_list(self, auto_select=False):
        self.output_folder = None
        folders = []
        
        for item in os.listdir("."):
            if item.startswith("检测结果输出") and os.path.isdir(item):
                folders.append(item)
        
        if folders:
            folders.sort(key=lambda x: os.path.getctime(x), reverse=True)
            self.output_folder = folders[0]
        
        if not self.output_folder or not os.path.exists(self.output_folder):
            self.image_combo['values'] = []
            self.image_var.set("")
            self.image_viewer.show_message("未找到检测结果输出文件夹\n请先运行分析程序")
            return
        
        image_files = []
        for root, dirs, files in os.walk(self.output_folder):
            for file in files:
                if file.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
                    rel_path = os.path.relpath(os.path.join(root, file), self.output_folder)
                    rel_path = rel_path.replace('\\', '/')
                    image_files.append(rel_path)
        
        if image_files:
            image_files.sort()
            self.image_combo['values'] = image_files
            
            if auto_select and self.last_viewed_image and self.last_viewed_image in image_files:
                self.image_var.set(self.last_viewed_image)
            elif not self.image_var.get() or self.image_var.get() not in image_files:
                self.image_var.set(image_files[0])
            
            self.on_image_select()
        else:
            self.image_combo['values'] = []
            self.image_var.set("")
            self.image_viewer.show_message("未找到图片文件\n请先运行分析程序生成图片")
        
        self.notebook.select(1)

    def on_image_select(self, event=None):
        img_path = self.image_var.get()
        if not img_path or not self.output_folder:
            return
        
        full_path = os.path.join(self.output_folder, img_path)
        if not os.path.exists(full_path):
            self.image_viewer.show_message(f"图片不存在: {img_path}")
            return
        
        self.image_viewer.load_image(full_path)
        
        self.last_viewed_image = img_path
        self.memory['last_image'] = img_path
        save_memory(self.memory)

    def reset_image_view(self):
        self.image_viewer._on_double_click(None)

    def open_image_folder(self):
        if self.output_folder and os.path.exists(self.output_folder):
            import subprocess
            subprocess.Popen(f'explorer "{os.path.abspath(self.output_folder)}"')
        else:
            messagebox.showwarning("提示", "未找到输出文件夹")

    def pause_analysis(self):
        pause_analysis()
        self.log("\n⏸ 正在暂停分析...\n")
        self.pause_btn.config(state='disabled')
        self.resume_btn.config(state='normal')
        self.paused = True

    def resume_analysis(self):
        resume_analysis()
        self.log("\n▶ 正在恢复分析...\n")
        self.resume_btn.config(state='disabled')
        self.pause_btn.config(state='normal')
        self.paused = False

    def stop_analysis(self):
        stop_analysis()
        self.log("\n⏹ 正在停止分析...\n")
        self.stop_btn.config(state='disabled')
        self.pause_btn.config(state='disabled')
        self.resume_btn.config(state='disabled')
        self.run_btn.config(state='normal')
        self.running = False
        self.paused = False

    def run_analysis(self):
        if self.running:
            messagebox.showinfo("提示", "程序正在运行中，请等待完成或停止")
            return

        if not self.change_working_directory():
            return

        # 自动切换到运行日志标签页
        self.notebook.select(0)

        reset_stop_flag()

        self.log_text.config(state='normal')
        self.log_text.insert('end', "\n" + "=" * 60 + "\n")
        self.log_text.insert('end', f"🔄 新一次运行开始 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')})\n")
        self.log_text.insert('end', "=" * 60 + "\n")
        self.log_text.config(state='disabled')
        self.log("🚀 开始运行分析程序...\n")
        self.log("=" * 50 + "\n")

        self.stop_btn.config(state='normal')
        self.pause_btn.config(state='normal')
        self.resume_btn.config(state='disabled')
        self.run_btn.config(state='disabled')

        config = self.get_config()
        self.running = True
        self.paused = False
        thread = threading.Thread(target=self._run_thread, args=(config,), daemon=True)
        thread.start()

    def _run_thread(self, config):
        def log_callback(msg):
            self.root.after(0, self.log, msg)

        work_dir = os.getcwd()
        success = run_analysis(config, log_callback=log_callback, work_dir=work_dir)
        self.root.after(0, self._finish_run, success)

    def _finish_run(self, success):
        self.running = False
        self.paused = False
        self.stop_btn.config(state='disabled')
        self.pause_btn.config(state='disabled')
        self.resume_btn.config(state='disabled')
        self.run_btn.config(state='normal')
        
        if success:
            self.log("\n✅ 分析完成！\n")
            self.log("\n🖼️ 正在加载生成的图片...\n")
            self.refresh_image_list(auto_select=True)
        else:
            if is_stopped():
                self.log("\n⏹ 分析已被用户停止\n")
            else:
                self.log("\n⚠️ 分析出错，请查看日志。\n")


# ============================================
# 8. 记忆功能辅助函数
# ============================================
MEMORY_FILE = "image_memory.json"

def load_memory():
    memory_file = os.path.join(get_script_directory(), MEMORY_FILE)
    if os.path.exists(memory_file):
        try:
            with open(memory_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_memory(memory):
    memory_file = os.path.join(get_script_directory(), MEMORY_FILE)
    try:
        with open(memory_file, 'w', encoding='utf-8') as f:
            json.dump(memory, f, ensure_ascii=False, indent=2)
    except IOError:
        pass


# ============================================
# 9. 启动程序
# ============================================
def main():
    splash = SplashScreen()
    
    # 第一步：显示启动信息（tkinter 已加载，这一步很快）
    splash.update_progress(3, "正在启动程序...")
    
    # 获取模块全局命名空间，加载的库注入进去，供顶层函数使用
    g = globals()
    
    # 第二步：加载 numpy（pandas 依赖它，先加载）
    splash.update_progress(10, "正在加载 numpy 数值计算库...")
    import numpy as np
    g['np'] = np
    splash.update_progress(25, "✓ numpy 加载完成")
    
    # 第三步：加载 pandas（通常最慢，占大头）
    splash.update_progress(30, "正在加载 pandas 数据处理库...")
    import pandas as pd
    g['pd'] = pd
    splash.update_progress(55, "✓ pandas 加载完成")
    
    # 第四步：加载 matplotlib
    splash.update_progress(58, "正在加载 matplotlib 绘图引擎...")
    import matplotlib
    matplotlib.use('TkAgg')
    import matplotlib.pyplot as plt
    import matplotlib.cm as cm
    g['matplotlib'] = matplotlib
    g['plt'] = plt
    g['cm'] = cm
    splash.update_progress(75, "✓ matplotlib 加载完成")
    
    # 第五步：加载 openpyxl
    splash.update_progress(78, "正在加载 openpyxl Excel 引擎...")
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.chart import ScatterChart, Reference, Series
    from openpyxl.chart.marker import Marker
    from openpyxl.drawing.line import LineProperties
    g['openpyxl'] = openpyxl
    g['Workbook'] = Workbook
    g['ScatterChart'] = ScatterChart
    g['Reference'] = Reference
    g['Series'] = Series
    g['Marker'] = Marker
    g['LineProperties'] = LineProperties
    splash.update_progress(88, "✓ openpyxl 加载完成")
    
    # 第六步：加载 PIL
    splash.update_progress(90, "正在加载 PIL 图像处理库...")
    from PIL import Image, ImageTk
    g['Image'] = Image
    g['ImageTk'] = ImageTk
    splash.update_progress(96, "✓ PIL 加载完成")
    
    # 第七步：初始化界面
    splash.update_progress(98, "正在初始化界面组件...")
    splash.update_progress(100, "加载完成！")
    
    splash.close()
    
    root = tk.Tk()
    app = ConfigApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
