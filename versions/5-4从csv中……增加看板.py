# -*- coding: utf-8 -*-
"""
PQ产品一致性检测分析程序 - 最终完整版（含所有GUI类定义）
所有功能完整保留，支持实时日志和文件日志。
"""

import sys
import os
import glob
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.drawing.line import LineProperties
import warnings
from datetime import datetime

warnings.filterwarnings('ignore')
print("✅ 当前运行的是最新版代码，滚轮已修复！")
# ============================================
# 1. 默认配置
# ============================================
DEFAULT_CONFIG = {
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
    "PRINT_DEBUG_LOG": True,
}

# ============================================
# 2. GUI中文名称与分组
# ============================================
PARAM_CN_NAME = {
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
    "RUN_PERFORMANCE_PLOTS": "生成性能图",
    "RUN_PQ_COMPARE_PLOT": "生成PQ对比图",
    "RUN_HYSTERESIS_PLOT": "生成迟滞图",
    "RUN_CONSISTENCY_PLOT": "生成一致性图",
    "RUN_AVG_PRESSURE_PLOT": "生成压差平均值图",
    "RUN_EXCEL_CHARTS": "Excel内嵌图表",
    "RUN_EXCEL_REPORT": "生成Excel报告",
    "RUN_SINGLE_GOBACK_PQ": "生成单文件去回程图",
    "RUN_ALL_GOBACK_SUMMARY": "生成全体去回程总对比",
    "PLOT_DPI": "基础图片DPI",
    "SUFFIX_AUTO_NUM": "自动添加序号防覆盖",
    "PRINT_DEBUG_LOG": "打印调试日志"
}

GROUPS = [
    ("输出文件基础配置", ["OUTPUT_EXCEL", "ENCODING"]),
    ("CSV数据列索引配置", ["COL_CURRENT", "COL_FLOW", "COL_PRESS"]),
    ("标准检测流量点", ["REF_FLOW"]),
    ("迟滞 & 流量匹配容错配置", ["PLOT_FLOW", "STD_LIMIT", "TOLERANCE", "STD_LINE1__value", "STD_LINE1__show"]),
    ("一致性计算模式 & 判定参数", ["CONSISTENCY_LIMIT_PCT", "CONSISTENCY_MODE", "BASE_FILE_NAME"]),
    ("加权一致性分析参数", ["KEY_FLOW", "HIGH_FLOW", "KEY_CURRENTS"]),
    ("滑动平均滤波配置", ["FILTER_WINDOW", "FILTER_ON"]),
    ("图表通用全局样式", ["X_AXIS_EXTEND", "HIDE_TOP_RIGHT_BORDER"]),
    ("图例独立开关", ["PERF_SHOW_LEGEND", "PQ_COMPARE_SHOW_LEGEND", "HYSTERESIS_SHOW_LEGEND",
                      "CON_SHOW_LEGEND", "AVG_PRESS_SHOW_LEGEND", "GOBACK_SHOW_LEGEND", "GOBACK_SUMMARY_SHOW_LEGEND"]),
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
    ("功能运行总开关", ["RUN_PERFORMANCE_PLOTS", "RUN_PQ_COMPARE_PLOT", "RUN_HYSTERESIS_PLOT", "RUN_CONSISTENCY_PLOT",
                      "RUN_AVG_PRESSURE_PLOT", "RUN_EXCEL_CHARTS", "RUN_EXCEL_REPORT", "RUN_SINGLE_GOBACK_PQ", "RUN_ALL_GOBACK_SUMMARY"]),
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


# ============================================
# 3. 全局辅助函数
# ============================================
def save_safe_plot(fig, filename, OUTPUT_FOLDER, dpi=150, bbox_inches='tight'):
    path = os.path.join(OUTPUT_FOLDER, filename)
    base, ext = os.path.splitext(path)
    counter = 1
    while os.path.exists(path):
        path = f"{base}({counter}){ext}"
        counter += 1
    fig.savefig(path, dpi=dpi, bbox_inches=bbox_inches)
    plt.close(fig)
    return path

def get_closest_index(arr, target):
    return np.argmin(np.abs(arr - target))

def safe_percent(numerator, denominator, default=0):
    try:
        return round((numerator / denominator) * 100, 1)
    except ZeroDivisionError:
        return default

def filter_press_by_current(df, window, filter_on):
    if not filter_on:
        return df
    df_out = df.copy()
    for curr, group in df.groupby("电流值"):
        filtered_press = group["压差值"].rolling(window=window, center=True, min_periods=1).mean()
        df_out.loc[group.index, "压差值滤波"] = filtered_press
    return df_out

def clean_invalid_data(df):
    df = df.copy()
    df = df.replace([np.inf, -np.inf], np.nan)
    df = df.dropna()
    return df

def analyze_consistency_influence_final(df_valid, config):
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

def calculate_consistency(df_valid, config):
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
# 4. 核心分析函数
# ============================================
def run_analysis(config, log_callback=None):
    """执行完整PQ分析流程，同时输出到GUI和文件日志"""
    # 内部重新导入，确保作用域正确
    import numpy as np
    import matplotlib.pyplot as plt

    # ========== 关键修复：正确切换工作目录 ==========
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)
    # =================================================

    # 创建文件日志
    log_file_path = os.path.join(os.getcwd(), "analysis_log.txt")
    log_file = open(log_file_path, "w", encoding="utf-8")
    log_file.write("===== PQ分析日志开始 =====\n")
    log_file.write(f"📁 工作目录: {os.getcwd()}\n")
    log_file.flush()

    def write_log(msg):
        log_file.write(msg)
        log_file.flush()
        if log_callback:
            if hasattr(run_analysis, 'root'):
                run_analysis.root.after(0, lambda: log_callback(msg))
            else:
                log_callback(msg)

    original_print = print
    def custom_print(*args, **kwargs):
        text = " ".join(str(a) for a in args)
        write_log(text + "\n")
        original_print(text)
    globals()['print'] = custom_print

    try:
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

        # ---------- 主流程 ----------
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
            try:
                df = pd.read_csv(file, usecols=[1, 4, 10], encoding=config["ENCODING"])
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
        total_files = df_raw["数据源文件"].nunique()
        all_valid_groups = df_raw.groupby(["电流值", "参考流量值", "数据分区"])["校验结果"].apply(
            lambda x: (x == "TRUE").all()
        )
        valid_groups = all_valid_groups[all_valid_groups].index.tolist()
        df_valid = df_raw[df_raw.set_index(["电流值", "参考流量值", "数据分区"]).index.isin(valid_groups)].copy()
        write_log("✅ 数据预处理全部完成\n")

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

        # ---- 步骤6：单文件性能图 ----
        write_log("\n📈 【6/12】正在生成单产品性能曲线图...\n")
        if config["RUN_PERFORMANCE_PLOTS"]:
            file_list = df_valid["数据源文件"].unique()
            perf_folder = "性能图表"
            perf_full = os.path.join(OUTPUT_FOLDER, perf_folder)
            if not os.path.exists(perf_full):
                os.makedirs(perf_full)

            for fig_idx, file in enumerate(file_list, 1):
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
                save_safe_plot(fig, os.path.join(perf_folder, f"性能图表-{fig_idx}.png"), OUTPUT_FOLDER)
                write_log(f"  已生成 {file} 的性能图\n")
            write_log("✅ 单产品性能图表生成完成\n")
        else:
            write_log("ℹ️ 已跳过：单产品性能图表\n")

        # ---- 步骤7：PQ总对比 ----
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
                        label = f if show_label else ""
                        if show_label:
                            plotted_files.add(f)
                        ax.plot(
                            cdat["flow"], cdat["press"],
                            marker='o', ms=3, linewidth=0.6,
                            color=file_color_map[f], label=label, alpha=0.8
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

        # ---- 步骤8：迟滞计算 ----
        write_log("\n🔁 【8/12】正在计算产品迟滞数据...\n")
        hys_data = []
        for (f, c, rf), g in df_valid.groupby(["数据源文件", "电流值", "参考流量值"]):
            b = g[g["数据分区"] == "前50%"]["实际压差值"]
            a = g[g["数据分区"] == "后50%"]["实际压差值"]
            if len(b) == 1 and len(a) == 1:
                hys_data.append([f, c, rf, round(abs(b.iloc[0] - a.iloc[0]), 2)])
        df_hys = pd.DataFrame(hys_data, columns=["数据源文件", "电流值", "参考流量值", "迟滞"])
        write_log("✅ 迟滞数据计算完成\n")

        # ---- 步骤9：一致性计算 ----
        write_log("\n📐 【9/12】正在计算产品一致性数据...\n")
        df_con = calculate_consistency(df_valid, config)
        write_log("✅ 一致性数据计算完成\n")

        # ---- 步骤10：Excel报告 ----
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

        # ---- 步骤11：PNG图表 ----
        write_log("\n🎨 【11/12】正在生成迟滞、一致性、压差平均值图表...\n")
        if config["RUN_HYSTERESIS_PLOT"]:
            fig, ax = plt.subplots(figsize=(15, 6), dpi=100)
            COLOR_MAP = {0: '#70AD47', 1: '#00B0F0', 2: '#FFC000', 3: '#FF0000', 4: '#C00000'}
            for prod in plot_hys["数据源文件"].unique():
                sub = plot_hys[plot_hys["数据源文件"] == prod]
                exceed_num = sum(sub["迟滞"] > config["STD_LIMIT"])
                color_idx = min(exceed_num, 4)
                color = COLOR_MAP[color_idx]
                ax.plot(sub["电流值"], sub["迟滞"], marker='o', ms=3, lw=1, color=color, label=prod)
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
            save_safe_plot(fig, "迟滞图.png", OUTPUT_FOLDER)
            write_log("✅ 迟滞图生成完成\n")
        else:
            write_log("ℹ️ 已跳过：迟滞图\n")

        # ---- 一致性图 ----
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
            ax.set_xlabel("流量/Lpm", fontsize=config["PERF_AXIS_LABEL_SIZE"],
                          weight="bold" if config["PERF_AXIS_LABEL_BOLD"] else "normal")
            ax.set_ylabel("压差 bar", fontsize=config["PERF_AXIS_LABEL_SIZE"],
                          weight="bold" if config["PERF_AXIS_LABEL_BOLD"] else "normal")
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
            save_safe_plot(fig, "一致性图.png", OUTPUT_FOLDER)
            write_log("✅ 一致性图生成完成\n")
        else:
            write_log("ℹ️ 已跳过：一致性图\n")

        # ---- 压差平均值图 ----
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
            ax_avg.set_xlabel("流量/Lpm", fontsize=config["PERF_AXIS_LABEL_SIZE"],
                              weight="bold" if config["PERF_AXIS_LABEL_BOLD"] else "normal")
            ax_avg.set_ylabel("压差 bar", fontsize=config["PERF_AXIS_LABEL_SIZE"],
                              weight="bold" if config["PERF_AXIS_LABEL_BOLD"] else "normal")
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
            save_safe_plot(fig_avg, "压差平均值图.png", OUTPUT_FOLDER)
            write_log("✅ 压差平均值图生成完成\n")
        else:
            write_log("ℹ️ 已跳过：压差平均值图\n")

        # ---- 步骤12：去程/回程 ----
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
                if file not in file_cleaned_data:
                    continue

                df = file_cleaned_data[file].copy()
                file_short = os.path.splitext(file)[0]
                curr_list = sorted(df["电流值"].unique())
                file_go = []
                file_back = []

                # 去程
                fig1, ax1 = plt.subplots(figsize=(12, 6), dpi=100)
                for curr in curr_list:
                    sub = df[df["电流值"] == curr].copy()
                    sub_go = sub.iloc[:len(sub) // 2].sort_values("流量值")
                    x = sub_go["流量值"].abs()
                    y = sub_go["压差值"].abs()
                    if len(x) >= 2:
                        ax1.plot(x, y, lw=config["PERF_LINE_WIDTH"], color=LINE_COLOR)
                        ax1.text(x.iloc[-1] + config["PERF_TEXT_OFFSET_X"], y.iloc[-1] + config["PERF_TEXT_OFFSET_Y"],
                                 f"{round(curr)}mA", fontsize=config["PERF_CURR_TEXT_SIZE"], va="center")
                        file_go.append((x, y))
                ax1.set_title(f"{file_short} 去程PQ", fontsize=config["PERF_TITLE_SIZE"],
                              x=config["PERF_TITLE_OFFSET_X"], y=config["PERF_TITLE_OFFSET_Y"])
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
                save_safe_plot(fig1, os.path.join(folder_go, f"去程PQ_{file_idx}.png"), OUTPUT_FOLDER)
                all_go_data.append(file_go)

                # 回程
                fig2, ax2 = plt.subplots(figsize=(12, 6), dpi=100)
                for curr in curr_list:
                    sub = df[df["电流值"] == curr].copy()
                    sub_back = sub.iloc[len(sub) // 2:].sort_values("流量值")
                    x = sub_back["流量值"].abs()
                    y = sub_back["压差值"].abs()
                    if len(x) >= 2:
                        ax2.plot(x, y, lw=config["PERF_LINE_WIDTH"], color=LINE_COLOR)
                        ax2.text(x.iloc[-1] + config["PERF_TEXT_OFFSET_X"], y.iloc[-1] + config["PERF_TEXT_OFFSET_Y"],
                                 f"{round(curr)}mA", fontsize=config["PERF_CURR_TEXT_SIZE"], va="center")
                        file_back.append((x, y))
                ax2.set_title(f"{file_short} 回程PQ", fontsize=config["PERF_TITLE_SIZE"],
                              x=config["PERF_TITLE_OFFSET_X"], y=config["PERF_TITLE_OFFSET_Y"])
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
                save_safe_plot(fig2, os.path.join(folder_back, f"回程PQ_{file_idx}.png"), OUTPUT_FOLDER)
                all_back_data.append(file_back)

                # 合并
                fig3, ax3 = plt.subplots(figsize=(12, 6), dpi=100)
                for curr in curr_list:
                    sub = df[df["电流值"] == curr].copy()
                    sub_go = sub.iloc[:len(sub) // 2].sort_values("流量值")
                    xg = sub_go["流量值"].abs()
                    yg = sub_go["压差值"].abs()
                    if len(xg) >= 2:
                        ax3.plot(xg, yg, lw=config["PERF_LINE_WIDTH"], color=LINE_COLOR)

                    sub_back = sub.iloc[len(sub) // 2:].sort_values("流量值")
                    xb = sub_back["流量值"].abs()
                    yb = sub_back["压差值"].abs()
                    if len(xb) >= 2:
                        ax3.plot(xb, yb, lw=config["PERF_LINE_WIDTH"], color=LINE_COLOR)

                    if len(xg) >= 2:
                        ax3.text(xg.iloc[-1] + config["PERF_TEXT_OFFSET_X"], yg.iloc[-1] + config["PERF_TEXT_OFFSET_Y"],
                                 f"{round(curr)}mA", fontsize=config["PERF_CURR_TEXT_SIZE"], va="center")
                ax3.set_title(f"{file_short} 去程+回程PQ", fontsize=config["PERF_TITLE_SIZE"],
                              x=config["PERF_TITLE_OFFSET_X"], y=config["PERF_TITLE_OFFSET_Y"])
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
                save_safe_plot(fig3, os.path.join(folder_both, f"去程回程对比_{file_idx}.png"), OUTPUT_FOLDER)
                all_both_data.append((file_go, file_back))

            write_log("✅ 单文件去程/回程/合并PQ图生成完成\n")
        else:
            write_log("ℹ️ 已跳过：单文件去程/回程/合并PQ图\n")

        # 全体去程回程总对比
        if config["RUN_ALL_GOBACK_SUMMARY"] and config["RUN_SINGLE_GOBACK_PQ"]:
            write_log("\n📊 正在生成所有文件PQ总对比图...\n")
            folder_go = "01_去程PQ"
            folder_back = "02_回程PQ"
            folder_both = "03_去程+回程PQ"

            fig_go, ax_go = plt.subplots(figsize=(15, 8), dpi=150)
            for file_data in all_go_data:
                for x, y in file_data:
                    ax_go.plot(x, y, color=LINE_COLOR, lw=1.0)
            ax_go.set_title("所有文件 去程PQ总对比", fontsize=20, pad=20)
            ax_go.set_xlabel("流量 L/min", fontsize=14)
            ax_go.set_ylabel("压差 bar", fontsize=14)
            ax_go.set_xlim(right=70)
            if config["HIDE_TOP_RIGHT_BORDER"]:
                ax_go.spines['top'].set_visible(False)
                ax_go.spines['right'].set_visible(False)
            if config["GOBACK_SUMMARY_SHOW_LEGEND"]:
                ax_go.legend(frameon=False)
            plt.tight_layout()
            save_safe_plot(fig_go, os.path.join(folder_go, "所有文件_去程总对比.png"), OUTPUT_FOLDER)

            fig_back, ax_back = plt.subplots(figsize=(15, 8), dpi=150)
            for file_data in all_back_data:
                for x, y in file_data:
                    ax_back.plot(x, y, color=LINE_COLOR, lw=1.0)
            ax_back.set_title("所有文件 回程PQ总对比", fontsize=20, pad=20)
            ax_back.set_xlabel("流量 L/min", fontsize=14)
            ax_back.set_ylabel("压差 bar", fontsize=14)
            ax_back.set_xlim(right=70)
            if config["HIDE_TOP_RIGHT_BORDER"]:
                ax_back.spines['top'].set_visible(False)
                ax_back.spines['right'].set_visible(False)
            if config["GOBACK_SUMMARY_SHOW_LEGEND"]:
                ax_back.legend(frameon=False)
            plt.tight_layout()
            save_safe_plot(fig_back, os.path.join(folder_back, "所有文件_回程总对比.png"), OUTPUT_FOLDER)

            fig_both, ax_both = plt.subplots(figsize=(15, 8), dpi=150)
            for go_data, back_data in all_both_data:
                for x, y in go_data:
                    ax_both.plot(x, y, color=LINE_COLOR, lw=1.0)
                for x, y in back_data:
                    ax_both.plot(x, y, color=LINE_COLOR, lw=1.0)
            ax_both.set_title("所有文件 去程+回程PQ总对比", fontsize=20, pad=20)
            ax_both.set_xlabel("流量 L/min", fontsize=14)
            ax_both.set_ylabel("压差 bar", fontsize=14)
            ax_both.set_xlim(right=70)
            if config["HIDE_TOP_RIGHT_BORDER"]:
                ax_both.spines['top'].set_visible(False)
                ax_both.spines['right'].set_visible(False)
            if config["GOBACK_SUMMARY_SHOW_LEGEND"]:
                ax_both.legend(frameon=False)
            plt.tight_layout()
            save_safe_plot(fig_both, os.path.join(folder_both, "所有文件_去程回程总对比.png"), OUTPUT_FOLDER)
            write_log("✅ 全体去程回程总对比图生成完成！\n")
        else:
            write_log("ℹ️ 已跳过：全体去程回程总对比图\n")

        # ---- 最终结果 ----
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
        write_log(f"❌ 运行出错：{e}\n{error_msg}\n")
        log_file.close()
        return False


# ============================================
# 5. GUI 界面类
# ============================================
class CollapsiblePanel:
    def __init__(self, parent, title, *args, **kwargs):
        self.frame = ttk.Frame(parent, *args, **kwargs)
        self.toggle_btn = ttk.Button(self.frame, text=f"▶ {title}", command=self.toggle)
        self.toggle_btn.pack(fill='x', pady=2)
        self.content = ttk.Frame(self.frame)
        self.is_expanded = False
        self.row = 0

    def toggle(self):
        if self.is_expanded:
            self.content.pack_forget()
            self.toggle_btn.config(text=f"▶ {self.toggle_btn.cget('text')[2:]}")
            self.is_expanded = False
        else:
            self.content.pack(fill='x', padx=10, pady=5)
            self.toggle_btn.config(text=f"▼ {self.toggle_btn.cget('text')[2:]}")
            self.is_expanded = True

    def add_row(self, label_text, widget):
        # 判断是否为复选框
        if isinstance(widget, ttk.Checkbutton):
            # 复选框：标签和控件在同一行
            lbl = ttk.Label(self.content, text=label_text, anchor='w')
            lbl.grid(row=self.row, column=0, sticky='w', padx=(0, 2), pady=2)
            widget.grid(row=self.row, column=1, sticky='w', padx=(0, 2), pady=2)
            self.content.columnconfigure(0, weight=1)
            self.content.columnconfigure(1, weight=0)  # 复选框不伸缩
            self.row += 1
        else:
            # 输入框：标签占一行，输入框占一行
            # 标签行（自动换行）
            lbl = ttk.Label(
                self.content,
                text=label_text,
                wraplength=300,
                anchor='w',
                justify='left'
            )
            lbl.grid(row=self.row, column=0, columnspan=2, sticky='w', padx=2, pady=(2, 0))
            self.row += 1
            # 输入框行（填充整行）
            widget.grid(row=self.row, column=0, columnspan=2, sticky='ew', padx=2, pady=(0, 2))
            # 设置列权重，使输入框填充整行
            self.content.columnconfigure(0, weight=1)
            self.content.columnconfigure(1, weight=0)
            self.row += 1
class ConfigApp:
    def __init__(self, root):
        self.root = root
        root.title("PQ产品一致性检测程序 - 配置面板")
        root.geometry("950x750")

        main_frame = ttk.Frame(root)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        config_frame = ttk.LabelFrame(main_frame, text="配置参数", padding=5)
        config_frame.pack(side='left', fill='both', expand=True)

        canvas_frame = ttk.Frame(config_frame)
        canvas_frame.pack(fill='both', expand=True)

        self.canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient='vertical', command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor='nw')
        self.canvas.configure(yscrollcommand=scrollbar.set)

        def _configure_canvas(event):
            self.canvas.itemconfig(self.canvas_window, width=event.width)
        self.canvas.bind('<Configure>', _configure_canvas)

        self.canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # ---------- 滚轮事件（全局绑定，但仅在配置区域内生效） ----------
        def on_mousewheel(event):
            # 判断事件源是否在配置区域内
            widget = event.widget
            while widget:
                if widget in (self.canvas, self.scrollable_frame):
                    break
                widget = widget.master
            else:
                return  # 不在配置区域，忽略

            # 滚动方向判断（兼容 Windows / Linux）
            if hasattr(event, 'delta') and event.delta:
                delta = -1 if event.delta > 0 else 1
            elif hasattr(event, 'num') and event.num in (4, 5):
                delta = -1 if event.num == 4 else 1 if event.num == 5 else 0
            else:
                return

            # 滚动步长（每次移动视口高度的 3%，可调整）
            step = 0.03 * delta
            first, last = self.canvas.yview()
            new_first = max(0.0, min(1.0 - (last - first), first + step))
            self.canvas.yview_moveto(new_first)

        # 绑定到根窗口，全局捕获滚轮事件
        self.root.bind_all("<MouseWheel>", on_mousewheel)

        self.widgets = {}

        for group_name, flat_keys in GROUPS:
            panel = CollapsiblePanel(self.scrollable_frame, group_name)
            panel.frame.pack(fill='x', pady=2)
            for fk in flat_keys:
                default = FLAT_DEFAULTS.get(fk, "")
                display_name = PARAM_CN_NAME.get(fk, fk)
                if isinstance(default, bool):
                    var = tk.BooleanVar(value=default)
                    cb = ttk.Checkbutton(panel.content, variable=var)
                    panel.add_row(display_name, cb)
                    self.widgets[fk] = var
                elif isinstance(default, int) and fk not in ["COL_CURRENT", "COL_FLOW", "COL_PRESS", "CONSISTENCY_MODE",
                                                              "FILTER_WINDOW", "STD_LINE1__show", "LABEL_POS_0mA",
                                                              "LABEL_POS_OTHER"]:
                    spin = ttk.Spinbox(panel.content, from_=0, to=100000, width=20)
                    spin.insert(0, str(default))
                    panel.add_row(display_name, spin)
                    self.widgets[fk] = spin
                else:
                    entry = ttk.Entry(panel.content, width=40)
                    entry.insert(0, str(default))
                    panel.add_row(display_name, entry)
                    self.widgets[fk] = entry
            panel.content.pack_forget()

        control_frame = ttk.LabelFrame(main_frame, text="运行控制", padding=10)
        control_frame.pack(side='right', fill='y', padx=10)

        ttk.Button(control_frame, text="🚀 运行分析", command=self.run_analysis, width=15).pack(pady=5)
        ttk.Button(control_frame, text="恢复默认", command=self.reset_defaults, width=15).pack(pady=5)

        log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding=5)
        log_frame.pack(side='bottom', fill='both', expand=True, pady=10)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, state='disabled', wrap='word')
        self.log_text.pack(fill='both', expand=True)

        self.running = False
        run_analysis.root = self.root

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

    def run_analysis(self):
        if self.running:
            messagebox.showinfo("提示", "程序正在运行中，请等待完成")
            return

        self.log_text.config(state='normal')
        self.log_text.delete(1.0, 'end')
        self.log_text.config(state='disabled')
        self.log("🚀 开始运行分析程序...\n")
        self.log("=" * 50 + "\n")

        config = self.get_config()
        self.running = True
        thread = threading.Thread(target=self._run_thread, args=(config,), daemon=True)
        thread.start()

    def _run_thread(self, config):
        def log_callback(msg):
            self.root.after(0, self.log, msg)

        success = run_analysis(config, log_callback=log_callback)
        self.root.after(0, self._finish_run, success)

    def _finish_run(self, success):
        self.running = False
        if success:
            self.log("\n✅ 分析完成！\n")
        else:
            self.log("\n❌ 分析过程中出现错误，请查看日志。\n")


# ============================================
# 6. 启动程序
# ============================================
def main():
    root = tk.Tk()
    app = ConfigApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()