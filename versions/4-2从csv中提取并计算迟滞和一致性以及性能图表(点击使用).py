import pandas as pd
import numpy as np
import os
import glob
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.drawing.line import LineProperties
import warnings
import sys

# 强制切换工作目录为脚本所在路径，确保文件读取正常
os.chdir(os.path.dirname(os.path.abspath(sys.argv[0])))

# ======================================
# ====================== 自定义配置区 ======================
# 说明：所有参数可直接修改，程序运行时自动生效，无需改动核心代码
# ======================================
warnings.filterwarnings('ignore')  # 忽略无关警告，保持控制台整洁
CONFIG = {
    # ---------------------- 输出文件基础配置 ----------------------
    "OUTPUT_EXCEL": "产品检测结果总表.xlsx",  # 最终生成的Excel检测报告名称
    "ENCODING": "gbk",                      # CSV文件编码（Windows中文系统默认编码）

    # ---------------------- CSV数据列索引配置 ----------------------
    # 索引从0开始，根据实际CSV文件调整
    "COL_CURRENT": 1,   # 电流值所在列索引
    "COL_FLOW": 4,      # 流量值所在列索引
    "COL_PRESS": 10,    # 压差值所在列索引

    # ---------------------- 标准检测流量点 ----------------------
    "REF_FLOW": [5.0, 10.0, 20.0, 40.0, 60.0],  # 程序自动匹配的标准流量点位(L/min)

    # ---------------------- 迟滞分析专用配置 ----------------------
    "PLOT_FLOW": 20.0,  # 迟滞图、一致性图固定使用的流量点
    "STD_LIMIT": 10.0,          # 迟滞合格上限(bar)
    "STD_LINE1": {"value": 10.0, "show": 1},  # 迟滞图标准线数值+显示开关
    "TOLERANCE": 1.00,          # 流量匹配允许的最大误差范围(L/min)

    # ---------------------- 一致性判定标准 ----------------------
    "CONSISTENCY_LIMIT_PCT": 5.0,  # 产品一致性合格阈值(±5%)

    # ---------------------- 加权分析关键参数 ----------------------
    "KEY_FLOW": 20.0,        # 核心关注流量点
    "HIGH_FLOW": 40.0,       # 高流量参考点
    "KEY_CURRENTS": [0, 300, 600, 900],  # 重点分析的电流值

    # ---------------------- 数据滤波配置 ----------------------
    "FILTER_WINDOW": 201,    # 滑动平均滤波窗口大小（必须为奇数）
    "FILTER_ON": True,       # 数据滤波总开关

    # ---------------------- 图表通用样式配置 ----------------------
    "X_AXIS_EXTEND": 8,      # X轴右侧延长长度
    "HIDE_TOP_RIGHT_BORDER": True,  # 是否隐藏图表顶部/右侧边框

    # ====================== 【性能图表样式配置】 ======================
    "PERF_TITLE_SUFFIX": " 实测",          # 性能图标题后缀
    "PERF_TITLE_SIZE": 36,                # 标题字体大小
    "PERF_TITLE_BOLD": False,             # 标题是否加粗
    "PERF_TITLE_OFFSET_X": 0.5,    # 标题水平位置（0.5=居中，1=最右侧）
    "PERF_TITLE_OFFSET_Y": 1.1,    # 标题垂直位置（数值越大越靠上）
    "PERF_AXIS_LABEL_SIZE": 20,           # 坐标轴标签字体大小
    "PERF_AXIS_LABEL_BOLD": False,        # 坐标轴标签是否加粗
    "PERF_TICK_LABEL_SIZE": 20,           # 刻度值字体大小
    "PERF_LINE_WIDTH": 2.0,               # 曲线宽度
    "PERF_TICK_WIDTH": 1.2,               # 刻度线宽度

    # ====================== 性能图表电流标签样式 ======================
    "PERF_TEXT_OFFSET_X": 1.5,            # 电流标注X轴偏移量
    "PERF_TEXT_OFFSET_Y": 0.0,            # 电流标注Y轴偏移量
    "PERF_CURR_TEXT_SIZE": 20,            # 电流标注字体大小
    "PERF_CURR_TEXT_BOLD": False,         # 电流标注是否加粗

    # ====================== 【一致性图表样式配置】 ======================
    "CON_TITLE": "一致性",                      # 一致性图表主标题
    "CON_TITLE_SIZE": 24,                       # 主标题字体大小
    "CON_TITLE_BOLD": False,                    # 主标题是否加粗
    "CON_TITLE_OFFSET_X": 0.5,                  # 主标题水平位置
    "CON_TITLE_OFFSET_Y": 1.1,                  # 主标题垂直位置
    "CON_AXIS_LABEL_SIZE": 20,                  # 坐标轴标签字体大小
    "CON_AXIS_LABEL_BOLD": False,               # 坐标轴标签是否加粗
    "CON_TICK_LABEL_SIZE": 16,                  # 刻度值字体大小
    "CON_LINE_WIDTH": 2.0,                      # 一致性曲线宽度
    "CON_TICK_WIDTH": 1.2,                      # 刻度线宽度

    # ====================== 一致性图表标签样式 ======================
    "LABEL_POS_0mA": -14,                       # 0mA数据标签垂直偏移（负数向下）
    "LABEL_POS_OTHER": 8,                       # 其他电流标签垂直偏移（正数向上）
    "CON_LABEL_SIZE": 10,                       # 误差标注字体大小
    "CON_CURR_TEXT_SIZE": 20,                   # 电流值标注字体大小
    
    # ====================== 【功能运行总开关】 ======================
    "RUN_PERFORMANCE_PLOTS": True,       # 生成单文件性能图表
    "RUN_PQ_COMPARE_PLOT": True,         # 生成所有产品PQ总对比图
    "RUN_HYSTERESIS_PLOT": True,         # 生成PNG格式迟滞图
    "RUN_CONSISTENCY_PLOT": True,        # 生成一致性分析图
    "RUN_AVG_PRESSURE_PLOT": True,       # 生成压差平均值图
    "RUN_EXCEL_CHARTS": True,            # 在Excel中嵌入迟滞+一致性图表
}

# ---------------------- 图表颜色配置（固定标准，无需修改） ----------------------
# 曲线颜色映射：超标点数越多，颜色越红
COLOR_MAP = {0: '#70AD47', 1: '#00B0F0', 2: '#FFC000', 3: '#FF0000', 4: '#C00000'}
COLOR_EXCEL = ["70AD47", "00B0F0", "FFC000", "FF0000", "C00000"]  # Excel图表专用色值
# 多文件对比时自动分配的颜色列表
FILE_COLORS = {
    0: '#00B0F0', 1: '#FFC000', 2: '#70AD47', 3: '#FF0000', 4: '#C00000',
    5: '#9933FF', 6: '#00FF00', 7: '#00FFFF', 8: '#FF9900', 9: '#999999'
}

# ======================================
# ====================== 程序核心功能区 ======================
# ======================================

# ====================== 输出文件夹自动创建（防重名） ======================
# 功能：自动创建结果文件夹，若已存在则自动编号重命名
OUTPUT_FOLDER = "检测结果输出"
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)
else:
    count = 1
    while True:
        new_folder = f"{OUTPUT_FOLDER}_{count}"
        if not os.path.exists(new_folder):
            OUTPUT_FOLDER = new_folder
            os.makedirs(new_folder)
            break
        count += 1

# ====================== 图表安全保存函数（防覆盖） ======================
# 功能：保存图片时自动重命名，避免覆盖已有文件
# 参数：fig=图表对象, filename=保存名称, dpi=清晰度, bbox_inches=紧凑布局
def save_safe_plot(fig, filename, dpi=150, bbox_inches='tight'):
    path = os.path.join(OUTPUT_FOLDER, filename)
    base, ext = os.path.splitext(path)
    counter = 1
    # 文件已存在则自动添加(1)(2)后缀
    while os.path.exists(path):
        path = f"{base}({counter}){ext}"
        counter += 1
    fig.savefig(path, dpi=dpi, bbox_inches=bbox_inches)
    return path

# ====================== 通用工具函数 ======================
# 功能：在数组中找到最接近目标值的索引
def get_closest_index(arr, target):
    return np.argmin(np.abs(arr - target))

# 功能：安全计算百分比，避免除0报错
def safe_percent(numerator, denominator, default=0):
    try:
        return round((numerator / denominator) * 100, 1)
    except ZeroDivisionError:
        return default

# 功能：按电流分组，仅对压差值做滑动平均滤波，流量保持原始数据
def filter_press_by_current(df, window=201):
    if not CONFIG["FILTER_ON"]:
        return df
    df_out = df.copy()
    # 每个电流值独立滤波，保证数据准确性
    for curr, group in df.groupby("电流值"):
        filtered_press = group["压差值"].rolling(window=window, center=True, min_periods=1).mean()
        df_out.loc[group.index, "压差值滤波"] = filtered_press
    return df_out

# 功能：清洗无效数据（空值、无穷值、非数值）
def clean_invalid_data(df):
    df = df.replace([False, None, np.inf, -np.inf], np.nan)
    df = df.dropna(subset=["流量值", "压差值"])
    df["流量值"] = pd.to_numeric(df["流量值"], errors="coerce")
    df["压差值"] = pd.to_numeric(df["压差值"], errors="coerce")
    return df

# ====================== 加权一致性影响分析（核心算法） ======================
# 功能：计算每个产品对整体一致性的影响权重，定位异常产品
# 输出：产品偏差得分、影响占比、最差产品、修正系数
def analyze_consistency_influence_final(df_valid):
    influence = []
    KEY_CURRENTS = CONFIG["KEY_CURRENTS"]
    KEY_FLOW = CONFIG["KEY_FLOW"]
    HIGH_FLOW = CONFIG["HIGH_FLOW"]

    # 1. 计算每个工况的全局平均值（作为基准值）
    group_mean = df_valid.groupby(["电流值", "参考流量值", "数据分区"])["实际压差值"].mean().reset_index()
    group_mean.rename(columns={"实际压差值": "工况均值"}, inplace=True)
    df_merge = df_valid.merge(group_mean, on=["电流值", "参考流量值", "数据分区"], how="left")

    # 2. 逐点计算偏差并加权
    for _, row in df_merge.iterrows():
        I, rf, part = row["电流值"], row["参考流量值"], row["数据分区"]
        f = row["数据源文件"]
        P = row["实际压差值"]
        mean_ir = row["工况均值"]
        dev = P - mean_ir

        # 核心规则：0mA/300mA的5L/min、10L/min不参与加权计算
        if I in [0.0, 300.0] and rf in [5.0, 10.0]:
            continue

        # 配置加权系数
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
            # 关键流量点权重加倍
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
    # 3. 计算每个文件的加权总偏差
    file_wdev = df_inf.groupby("file")["weighted_dev"].sum()
    file_weight = df_inf.groupby("file")["weight"].sum()
    file_avg_dev = file_wdev / file_weight
    file_score = file_avg_dev.abs().sort_values(ascending=False)

    # 4. 计算产品修正系数
    global_mean_all = df_inf["mean_ir"].mean()
    suggest = {}
    for f in file_avg_dev.index:
        f_mean = df_inf[df_inf["file"] == f]["P"].mean()
        adj_mean = f_mean + file_avg_dev[f]
        scale = round(global_mean_all / adj_mean, 3) if adj_mean != 0 else 1.0
        suggest[f] = scale

    # 5. 计算影响占比
    total_abs = file_score.sum()
    impact_rate = {f: round(s / total_abs * 100, 1) for f, s in file_score.items()}
    worst_file = file_score.index[0]

    return file_score, impact_rate, worst_file, suggest

# ====================== 提取0mA 60L/min真实压差值 ======================
# 功能：提取0mA、60L/min、前半段数据，用于生成修正报告
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

# ====================== 主程序入口 ======================
print("="*70)
print("🚀 PQ 产品一致性检测分析程序 启动成功")
print("="*70)

# 确保程序在脚本所在目录运行
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# ---------------------- 步骤1：扫描数据文件 ----------------------
print("\n【1/12】正在扫描当前目录的CSV数据文件...")
csv_files = glob.glob("*.csv")
print(f"✅ 扫描完成：共找到 {len(csv_files)} 个有效数据文件")

raw_data = []            # 存储原始检测数据
pq_comparison_data = []  # 存储PQ对比图数据

# ---------------------- 步骤2：读取并处理数据 ----------------------
print("\n【2/12】正在读取数据并执行滤波处理...")
for file in csv_files:
    try:
        # 按配置读取电流、流量、压差三列数据
        df = pd.read_csv(file, usecols=[1,4,10], encoding="gbk")
        df.columns = ["电流值","流量值","压差值"]
    except:
        continue  # 读取失败则跳过该文件

    # 数据清洗
    df = clean_invalid_data(df)
    if df.empty:
        continue

    # 保留原始数据，仅对压差滤波
    df["流量值原始"] = df["流量值"].copy()
    df["压差值原始"] = df["压差值"].copy()
    df["压差值滤波"] = df["压差值"].copy()
    df = filter_press_by_current(df, CONFIG["FILTER_WINDOW"])

    df = clean_invalid_data(df)
    if df.empty:
        continue

    # 使用滤波后的压差数据
    df["压差值"] = df["压差值滤波"]
    file_short = os.path.splitext(file)[0]

    # 提取PQ对比图数据
    for I, g in df.groupby("电流值"):
        I = round(I, 2)
        mid = len(g) // 2
        first_half = g.iloc[:mid]  # 仅使用前50%数据
        first_half = clean_invalid_data(first_half)
        if first_half.empty:
            continue

        f_vals = first_half["流量值"].abs().values
        p_vals = first_half["压差值"].abs().values
        flow_points = []
        press_points = []
        # 匹配标准流量点
        for ref in CONFIG["REF_FLOW"]:
            idx = get_closest_index(f_vals, ref)
            fp = round(abs(f_vals[idx]),4)
            pp = round(abs(p_vals[idx]),4)
            flow_points.append(fp)
            press_points.append(pp)

        pq_comparison_data.append({
            "file": file_short, "current": I,
            "flow": flow_points, "press": press_points
        })

    # 生成完整检测数据表
    for I, g in df.groupby("电流值"):
        I = round(I, 2)
        mid = len(g) // 2
        # 拆分去程(前50%)、回程(后50%)数据
        for part, part_data in [("前50%", g.iloc[:mid]), ("后50%", g.iloc[mid:])]:
            part_data = clean_invalid_data(part_data)
            if part_data.empty:
                continue
            f_vals = part_data["流量值"].abs().values
            p_vals = part_data["压差值"].abs().values
            for ref in CONFIG["REF_FLOW"]:
                idx = get_closest_index(f_vals, ref)
                real_flow = round(abs(f_vals[idx]),4)
                real_press = round(abs(p_vals[idx]),4)
                diff = round(abs(real_flow - ref),4)
                check = "TRUE" if diff <= CONFIG["TOLERANCE"] else "FALSE"
                raw_data.append([file, I, part, ref, real_press, real_flow, diff, check])

# 构建原始数据表格
df_raw = pd.DataFrame(raw_data, columns=[
    "数据源文件","电流值","数据分区","参考流量值","实际压差值","实际流量值","差值数据","校验结果"
])

# ---------------------- 步骤3：筛选全文件有效数据 ----------------------
print("\n【3/12】正在筛选所有文件均合格的有效检测点位...")
total_files = df_raw["数据源文件"].nunique()
# 仅保留所有文件校验都为TRUE的检测点
all_valid_groups = df_raw.groupby(["电流值", "参考流量值", "数据分区"])["校验结果"].apply(
    lambda x: (x == "TRUE").all()
)
valid_groups = all_valid_groups[all_valid_groups].index.tolist()
df_valid = df_raw[df_raw.set_index(["电流值", "参考流量值", "数据分区"]).index.isin(valid_groups)].copy()
print("✅ 数据预处理全部完成")

# ---------------------- 步骤4：加权一致性分析 ----------------------
print("\n【4/12】正在执行加权一致性影响分析...")
file_score, impact_rate, worst_file, scale_suggest = analyze_consistency_influence_final(df_valid)
print("✅ 一致性分析完成，已定位异常产品")

# ---------------------- 步骤5：提取关键点位数据 ----------------------
print("\n【5/12】正在读取0mA 60L/min 基准压差值...")
all_files = df_valid["数据源文件"].unique()
orig_0mA60L = get_real_0mA60L_press(df_valid, all_files)
scale_0mA60L = {}
for f in all_files:
    orig = orig_0mA60L.get(f, 0.0)
    scale = scale_suggest.get(f, 1.0)
    new_p = round(orig * scale, 2)
    scale_0mA60L[f] = (orig, new_p)
print("✅ 关键点位数据提取完成")

# Matplotlib中文显示配置
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

# ---------------------- 步骤6：生成单文件性能图表 ----------------------
print("\n【6/12】正在生成单产品性能曲线图...")
if CONFIG["RUN_PERFORMANCE_PLOTS"]:
    file_list = df_valid["数据源文件"].unique()
    # 创建性能图表专用文件夹
    perf_folder = os.path.join(OUTPUT_FOLDER, "性能图表")
    if not os.path.exists(perf_folder):
        os.makedirs(perf_folder)

    for fig_idx, file in enumerate(file_list,1):
        df_file = df_valid[(df_valid["数据源文件"]==file) & (df_valid["数据分区"]=="前50%")].copy()
        if df_file.empty:
            continue
        fig, ax = plt.subplots(figsize=(12,6), dpi=100)
        curr_list = sorted(df_file["电流值"].unique())
        for curr in curr_list:
            sub = df_file[df_file["电流值"]==curr].sort_values("参考流量值")
            ax.plot(sub["参考流量值"], sub["实际压差值"], marker='o', ms=4, 
                    color='#00B0F0', lw=CONFIG["PERF_LINE_WIDTH"])
            # 曲线末端标注电流值
            x_last = sub["参考流量值"].iloc[-1]
            y_last = sub["实际压差值"].iloc[-1]
            ax.text(
                x_last + CONFIG["PERF_TEXT_OFFSET_X"],
                y_last + CONFIG["PERF_TEXT_OFFSET_Y"],
                f"{curr}mA",
                fontsize=CONFIG["PERF_CURR_TEXT_SIZE"],
                weight="bold" if CONFIG["PERF_CURR_TEXT_BOLD"] else "normal",
                va='center'
            )
        # 图表样式设置
        title_text = f"{os.path.splitext(file)[0]}{CONFIG['PERF_TITLE_SUFFIX']}"
        ax.set_title(title_text, fontsize=CONFIG["PERF_TITLE_SIZE"],
                     weight="bold" if CONFIG["PERF_TITLE_BOLD"] else "normal",
                     x=CONFIG["PERF_TITLE_OFFSET_X"], y=CONFIG["PERF_TITLE_OFFSET_Y"])  
        ax.set_xlabel("流量 L/min", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"],
                      weight="bold" if CONFIG["PERF_AXIS_LABEL_BOLD"] else "normal")
        ax.set_ylabel("压差 bar", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"],
                      weight="bold" if CONFIG["PERF_AXIS_LABEL_BOLD"] else "normal")
        ax.tick_params(axis='both', labelsize=CONFIG["PERF_TICK_LABEL_SIZE"], 
                       width=CONFIG["PERF_TICK_WIDTH"])
        ax.set_xlim(right=70)
        if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
        plt.tight_layout()
        save_safe_plot(fig, os.path.join("性能图表", f"性能图表-{fig_idx}.png"))
        plt.close()
    print("✅ 单产品性能图表生成完成")
else:
    print("ℹ️ 已跳过：单产品性能图表")

# ---------------------- 步骤7：生成PQ总对比图 ----------------------
print("\n【7/12】正在生成所有产品PQ总对比图...")
if CONFIG["RUN_PQ_COMPARE_PLOT"] and pq_comparison_data:
    fig, ax = plt.subplots(figsize=(18, 10), dpi=300)
    fig.set_facecolor('white')
    unique_files = sorted(list(set([item["file"] for item in pq_comparison_data])))
    unique_currents = sorted(list(set([item["current"] for item in pq_comparison_data])))
    
    # 配色配置
    import matplotlib.cm as cm
    colors = cm.nipy_spectral(np.linspace(0, 1, len(unique_files)))
    file_color_map = {f: colors[i] for i, f in enumerate(unique_files)}
    
    line_dict = {}
    plotted_files = set()

    # 绘制所有产品PQ曲线
    for f in unique_files:
        dat = [d for d in pq_comparison_data if d["file"] == f]
        for curr in unique_currents:
            cdat = next((d for d in dat if d["current"] == curr), None)
            if cdat:
                show_label = f not in plotted_files
                label = f if show_label else ""
                if show_label:
                    plotted_files.add(f)
                
                line = ax.plot(
                    cdat["flow"], cdat["press"],
                    marker='o', ms=3, linewidth=0.6,
                    color=file_color_map[f], label=label, alpha=0.8
                )[0]
                if show_label:
                    line_dict[line] = f

    # 统一标注电流值
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
                        x_last + CONFIG["PERF_TEXT_OFFSET_X"],
                        y_last + CONFIG["PERF_TEXT_OFFSET_Y"],
                        f"{curr}mA",
                        fontsize=CONFIG["CON_CURR_TEXT_SIZE"],
                        weight="bold" if CONFIG["CON_TITLE_BOLD"] else "normal",
                        va='center'
                    )

    # 图表样式
    ax.set_title("所有产品滤波后前50% PQ曲线对比", fontsize=16, weight='bold')
    ax.set_xlabel("流量 L/min", fontsize=14)
    ax.set_ylabel("压差 bar", fontsize=14)
    ax.grid(alpha=0.3)
    ax.set_xlim(0, max(CONFIG["REF_FLOW"]) + 10)
    ax.tick_params(axis='both', labelsize=12)
    if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

    # 图例优化
    ncol = min(3, max(1, len(unique_files) // 12))
    leg = ax.legend(loc="center left", bbox_to_anchor=(1.03, 0.5), ncol=ncol, fontsize=10)
    leg.handlelength = 1.2
    plt.tight_layout(rect=[0, 0, 0.87, 1])
    save_safe_plot(fig, "所有产品PQ曲线对比图.png", dpi=300)
    plt.close()
    print("✅ PQ总对比图生成完成")
else:
    print("ℹ️ 已跳过：PQ总对比图")

# ---------------------- 步骤8：计算迟滞数据 ----------------------
print("\n【8/12】正在计算产品迟滞数据...")
hys_data = []
for (f,c,rf),g in df_valid.groupby(["数据源文件","电流值","参考流量值"]):
    b = g[g["数据分区"]=="前50%"]["实际压差值"]
    a = g[g["数据分区"]=="后50%"]["实际压差值"]
    if len(b)==1 and len(a)==1:
        hys_data.append([f,c,rf,round(abs(b.iloc[0]-a.iloc[0]),2)])
df_hys = pd.DataFrame(hys_data, columns=["数据源文件","电流值","参考流量值","迟滞"])
print("✅ 迟滞数据计算完成")

# ---------------------- 步骤9：计算一致性数据 ----------------------
print("\n【9/12】正在计算产品一致性数据...")
con_data = []
for (c,rf,p),g in df_valid.groupby(["电流值","参考流量值","数据分区"]):
    vals = g["实际压差值"].dropna()
    if len(vals)<1: continue
    maxv = vals.max()
    minv = vals.min()
    avg = (maxv+minv)/2
    dif = (maxv-minv)/2
    con_data.append([round(c,2), round(rf,1), p,
                     round(maxv,2), g.loc[vals.idxmax(),"数据源文件"],
                     round(minv,2), g.loc[vals.idxmin(),"数据源文件"],
                     round(avg,2), round(dif,2), round(safe_percent(dif,avg), 2)])
df_con = pd.DataFrame(con_data, columns=[
    "电流值","参考流量值","数据分区","最大值","最大值文件名","最小值","最小值文件名","平均值","diff(半差值)","一致性%"
])
print("✅ 一致性数据计算完成")

# ---------------------- 步骤10：生成Excel检测报告 ----------------------
print("\n【10/12】正在生成Excel完整检测报告...")
wb = Workbook()
if "Sheet" in wb.sheetnames:
    del wb["Sheet"]

# 工作表1：原始检测数据表
ws1 = wb.create_sheet("原始检测数据")
ws1.append(["数据源文件","电流值","数据分区","参考流量值","实际压差值","实际流量值","差值数据","校验结果"])
for _,r in df_raw.iterrows(): 
    ws1.append(list(r))

# 工作表2：迟滞数据+Excel迟滞图表
ws2 = wb.create_sheet("产品迟滞")
ws2.append(["数据源文件","电流值","参考流量值","迟滞"])
for r in hys_data: 
    ws2.append(r)
# 提取绘图数据
plot_hys = df_hys[df_hys["参考流量值"]==CONFIG["PLOT_FLOW"]].copy().sort_values(["数据源文件","电流值"])
ws2.cell(1,8,"绘图文件名")
ws2.cell(1,9,"绘图电流")
ws2.cell(1,10,"绘图迟滞")
plot_last_row = len(plot_hys) + 1
for i,(_,r) in enumerate(plot_hys.iterrows(),2):
    ws2.cell(i,8,r["数据源文件"])
    ws2.cell(i,9,r["电流值"])
    ws2.cell(i,10,r["迟滞"])
ws2.cell(plot_last_row+1,8,"绘图数据结束")

# Excel内嵌迟滞图表
if CONFIG["RUN_EXCEL_CHARTS"]:
    chart_hys = ScatterChart()
    chart_hys.scatterStyle = "lineMarker"
    chart_hys.title = "P-Q滞环@20L/min"
    chart_hys.x_axis.title = "电流值"
    chart_hys.y_axis.title = "迟滞/bar"

    # 读取产品列表
    prod_list = []
    for row in range(2, plot_last_row+1):
        prod = ws2.cell(row,8).value
        if prod and prod not in prod_list:
            prod_list.append(prod)

    # 逐产品绘制曲线
    for prod in prod_list:
        sr = er = None
        for row in range(2, plot_last_row+1):
            if ws2.cell(row,8).value == prod:
                sr = row if sr is None else sr
                er = row
        if sr is None or er is None:
            continue
        
        x = Reference(ws2, 9, sr, 9, er)
        y = Reference(ws2, 10, sr, 10, er)
        
        # 按超标点数设置颜色
        exceed_num = 0
        for row in range(sr, er+1):
            hys_val = ws2.cell(row,10).value
            if hys_val and float(hys_val) > CONFIG["STD_LIMIT"]:
                exceed_num += 1
        color = COLOR_EXCEL[min(exceed_num,4)]
        
        ser = Series(y,x,title=prod)
        ser.marker = Marker(size=3)
        ser.graphicalProperties.line = LineProperties(solidFill=color, w=10000)
        chart_hys.series.append(ser)
    ws2.add_chart(chart_hys, "L2")

# 工作表3：一致性数据+Excel一致性图表
ws3 = wb.create_sheet("一致性对比")
ws3.append([
    "电流值","参考流量值","数据分区",
    "最大值","最大值文件名","最小值","最小值文件名",
    "平均值","diff(半差值)","一致性%"
])
for r in con_data: 
    ws3.append(r)

# 绘制一致性图表函数
def draw_con_chart(part, title, col, pos):
    data = df_con[df_con["数据分区"]==part].sort_values(["电流值","参考流量值"])
    currs = sorted(data["电流值"].unique())
    headers = ["电流值","参考流量值","最大值","最大值文件名","最小值","最小值文件名","平均值","diff(半差值)","一致性%","数据标签"]
    for i,h in enumerate(headers): 
        ws3.cell(1,col+i,h)
    
    data_last_row = 1
    for i,(_,r) in enumerate(data.iterrows(),2):
        curr = r["电流值"]
        flow = r["参考流量值"]
        dif = r["diff(半差值)"]
        pct = r["一致性%"]
        lab = f"±{dif:.2f}bar" if (curr in currs[:2] and flow in [5,10]) else f"±{pct:.2f}%"
        ws3.cell(i,col+0,curr)
        ws3.cell(i,col+1,flow)
        ws3.cell(i,col+2,r["最大值"])
        ws3.cell(i,col+3,r["最大值文件名"])
        ws3.cell(i,col+4,r["最小值"])
        ws3.cell(i,col+5,r["最小值文件名"])
        ws3.cell(i,col+6,r["平均值"])
        ws3.cell(i,col+7,dif)
        ws3.cell(i,col+8,pct)
        ws3.cell(i,col+9,lab)
        data_last_row = i
    ws3.cell(data_last_row+1, col, "绘图数据结束")
    
    ch = ScatterChart()
    ch.scatterStyle = "lineMarker"
    ch.title = f"一致性 {title}"
    ch.x_axis.title = "流量/Lpm"
    ch.y_axis.title = "压差/bar"
    ch.x_axis.majorGridlines = None
    ch.y_axis.majorGridlines = None
    
    # 读取电流值
    curr_list = []
    for row in range(2, data_last_row+1):
        curr_val = ws3.cell(row, col+0).value
        if curr_val and curr_val not in curr_list:
            curr_list.append(curr_val)
    curr_list = sorted(curr_list)
    
    # 逐电流绘图
    for curr in curr_list:
        sr = er = None
        for row in range(2, data_last_row+1):
            if ws3.cell(row, col+0).value == curr:
                sr = row if sr is None else sr
                er = row
        if sr is None or er is None:
            continue
        
        x = Reference(ws3, col+1, sr, col+1, er)
        y = Reference(ws3, col+6, sr, col+6, er)
        label_ref = Reference(ws3, col+9, sr, col+9, er)
        
        ser = Series(y,x,title=f"{curr}mA")
        ser.marker = Marker(size=4)
        ser.graphicalProperties.line = LineProperties(solidFill="00B0F0", w=8000)
        ser.marker.graphicalProperties.solidFill = "00B0F0"
        ser.dLbl = True
        ser.labelRef = label_ref
        ch.series.append(ser)
    ws3.add_chart(ch, pos)

if CONFIG["RUN_EXCEL_CHARTS"]:
    draw_con_chart("前50%", "前50%", 15, "AQ2")
    draw_con_chart("后50%", "后50%", 28, "AQ20")

# 工作表4：加权一致性分析报告
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
        rate = impact_rate.get(f,0)
        scale = scale_suggest.get(f,1.0)
        orig_p, new_p = scale_0mA60L.get(f, (0.0, 0.0))
        correct_msg = f"{orig_p} → {new_p}"
        note = f"全量程P值 × {scale}"
        ws4.append([f, round(score,2), rate, scale, correct_msg, note])

# 保存Excel文件（防重名）
excel_path = os.path.join(OUTPUT_FOLDER, CONFIG["OUTPUT_EXCEL"])
base, ext = os.path.splitext(excel_path)
cnt = 1
while os.path.exists(excel_path):
    excel_path = f"{base}({cnt}){ext}"
    cnt +=1
wb.save(excel_path)
print("✅ Excel检测报告生成完成")

# ---------------------- 步骤11：生成PNG分析图表 ----------------------
print("\n【11/12】正在生成迟滞、一致性、压差平均值图表...")
# 迟滞图
if CONFIG["RUN_HYSTERESIS_PLOT"]:
    fig, ax = plt.subplots(figsize=(15,6), dpi=100)
    for prod in plot_hys["数据源文件"].unique():
        sub = plot_hys[plot_hys["数据源文件"]==prod]
        exceed_num = sum(sub["迟滞"] > CONFIG["STD_LIMIT"])
        color_idx = min(exceed_num, 4)
        color = COLOR_MAP[color_idx]
        ax.plot(sub["电流值"], sub["迟滞"], marker='o', ms=3, lw=1, color=color, label=prod)
    if CONFIG["STD_LINE1"]["show"]:
        ax.plot(plot_hys["电流值"].unique(), [10]*len(plot_hys["电流值"].unique()), 'r--', lw=1, label="标准线10")
    ax.set_title("P-Q滞环@20L/min")
    ax.set_xlabel("电流值")
    ax.set_ylabel("迟滞/bar")
    ax.grid(axis='y')
    # 图例优化
    unique_prods = plot_hys["数据源文件"].unique()
    ncol = min(3, len(unique_prods) // 15 + 1)
    ax.legend(loc="center left", bbox_to_anchor=(1.02, 0.5), ncol=ncol, fontsize=8)
    if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
    plt.tight_layout()
    save_safe_plot(fig, "迟滞图.png")
    plt.close()
    print("✅ 迟滞图生成完成")
else:
    print("ℹ️ 已跳过：迟滞图")

# 一致性图
if CONFIG["RUN_CONSISTENCY_PLOT"]:
    con_plot = df_con[df_con["数据分区"]=="前50%"].copy()
    currs = sorted(con_plot["电流值"].unique())
    min_two_currents = sorted(con_plot["电流值"].unique())[:2]
    fig, ax = plt.subplots(figsize=(15,6), dpi=100)

    for idx,curr in enumerate(currs):
        sub = con_plot[con_plot["电流值"]==curr].sort_values("参考流量值")
        ax.plot(sub["参考流量值"], sub["平均值"], marker='o', ms=4, color='#00B0F0')
        
        # 电流标注
        x_last = sub["参考流量值"].iloc[-1]
        y_last = sub["平均值"].iloc[-1]
        ax.text(
            x_last + CONFIG["PERF_TEXT_OFFSET_X"],
            y_last + CONFIG["PERF_TEXT_OFFSET_Y"],
            f"{curr}mA",
            fontsize=CONFIG["CON_CURR_TEXT_SIZE"],
            weight="bold" if CONFIG["CON_TITLE_BOLD"] else "normal",
            va='center'
        )
        # 误差标注
        for x,y,dif,pct in zip(sub["参考流量值"], sub["平均值"], sub["diff(半差值)"], sub["一致性%"]):
            lab = f"±{dif:.2f}bar" if (curr in min_two_currents and x in [5.0, 10.0]) else f"±{pct:.2f}%"
            offset = CONFIG["LABEL_POS_0mA"] if curr == 0 else CONFIG["LABEL_POS_OTHER"]
            ax.annotate(lab, (x,y), xytext=(0, offset), textcoords='offset points', 
                        ha='center', fontsize=CONFIG["CON_LABEL_SIZE"])

    # 样式设置
    ax.set_title(CONFIG["CON_TITLE"], fontsize=CONFIG["CON_TITLE_SIZE"],
                 weight="bold" if CONFIG["CON_TITLE_BOLD"] else "normal",
                 x=CONFIG["CON_TITLE_OFFSET_X"], y=CONFIG["CON_TITLE_OFFSET_Y"])
    ax.set_xlabel("流量/Lpm", fontsize=CONFIG["CON_AXIS_LABEL_SIZE"],
                  weight="bold" if CONFIG["CON_AXIS_LABEL_BOLD"] else "normal")
    ax.set_ylabel("压差 bar", fontsize=CONFIG["CON_AXIS_LABEL_SIZE"],
                  weight="bold" if CONFIG["CON_AXIS_LABEL_BOLD"] else "normal")
    ax.tick_params(axis='both', labelsize=CONFIG["CON_TICK_LABEL_SIZE"], width=CONFIG["CON_TICK_WIDTH"])
    ax.set_xlim(right=ax.get_xlim()[1]+CONFIG["X_AXIS_EXTEND"])
    for line in ax.get_lines():
        line.set_linewidth(CONFIG["CON_LINE_WIDTH"])
    if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
    plt.tight_layout()
    save_safe_plot(fig, "一致性图.png")
    plt.close()
    print("✅ 一致性图生成完成")
else:
    print("ℹ️ 已跳过：一致性图")

# 压差平均值图
if CONFIG["RUN_AVG_PRESSURE_PLOT"]:
    con_plot_avg = df_con[df_con["数据分区"]=="前50%"].copy()
    currs_avg = sorted(con_plot_avg["电流值"].unique())
    min_two_currents_avg = sorted(con_plot_avg["电流值"].unique())[:2]

    fig_avg, ax_avg = plt.subplots(figsize=(15,6), dpi=100)
    for idx,curr in enumerate(currs_avg):
        sub_avg = con_plot_avg[con_plot_avg["电流值"]==curr].sort_values("参考流量值")
        ax_avg.plot(sub_avg["参考流量值"], sub_avg["平均值"], marker='o', ms=4, color='#00B0F0')
        
        # 电流标注
        x_last_avg = sub_avg["参考流量值"].iloc[-1]
        y_last_avg = sub_avg["平均值"].iloc[-1]
        ax_avg.text(
            x_last_avg + CONFIG["PERF_TEXT_OFFSET_X"],
            y_last_avg + CONFIG["PERF_TEXT_OFFSET_Y"],
            f"{curr}mA",
            fontsize=CONFIG["CON_CURR_TEXT_SIZE"],
            weight="bold" if CONFIG["CON_TITLE_BOLD"] else "normal",
            va='center'
        )
        # 平均值±误差标注
        for x,y,avg_val,diff_val in zip(sub_avg["参考流量值"], sub_avg["平均值"], sub_avg["平均值"], sub_avg["diff(半差值)"]):
            lab_avg = f"{avg_val:.2f}±{diff_val:.2f} bar"
            offset_avg = CONFIG["LABEL_POS_0mA"] if curr == 0 else CONFIG["LABEL_POS_OTHER"]
            ax_avg.annotate(lab_avg, (x,y), xytext=(0, offset_avg), textcoords='offset points', 
                        ha='center', fontsize=CONFIG["CON_LABEL_SIZE"])

    # 样式设置
    ax_avg.set_title("压差平均值", fontsize=CONFIG["CON_TITLE_SIZE"],
                     weight="bold" if CONFIG["CON_TITLE_BOLD"] else "normal",
                     x=CONFIG["CON_TITLE_OFFSET_X"], y=CONFIG["CON_TITLE_OFFSET_Y"])
    ax_avg.set_xlabel("流量/Lpm", fontsize=CONFIG["CON_AXIS_LABEL_SIZE"],
                      weight="bold" if CONFIG["CON_AXIS_LABEL_BOLD"] else "normal")
    ax_avg.set_ylabel("压差 bar", fontsize=CONFIG["CON_AXIS_LABEL_SIZE"],
                      weight="bold" if CONFIG["CON_AXIS_LABEL_BOLD"] else "normal")
    ax_avg.tick_params(axis='both', labelsize=CONFIG["CON_TICK_LABEL_SIZE"], width=CONFIG["CON_TICK_WIDTH"])
    ax_avg.set_xlim(right=ax_avg.get_xlim()[1]+CONFIG["X_AXIS_EXTEND"])
    for line in ax_avg.get_lines():
        line.set_linewidth(CONFIG["CON_LINE_WIDTH"])
    if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
        ax_avg.spines['top'].set_visible(False)
        ax_avg.spines['right'].set_visible(False)

    plt.tight_layout()
    save_safe_plot(fig_avg, "压差平均值图.png")
    plt.close()
    print("✅ 压差平均值图生成完成")
else:
    print("ℹ️ 已跳过：压差平均值图")

# ---------------------- 步骤12：生成去程/回程PQ图表 ----------------------
print("\n【12/12】正在生成去程、回程、合并PQ曲线图...")
# 创建专用文件夹
folder_go = os.path.join(OUTPUT_FOLDER, "01_去程PQ")
folder_back = os.path.join(OUTPUT_FOLDER, "02_回程PQ")
folder_both = os.path.join(OUTPUT_FOLDER, "03_去程+回程PQ")
for f in [folder_go, folder_back, folder_both]:
    if not os.path.exists(f):
        os.makedirs(f)

LINE_COLOR = "#00B0F0"
all_go_data = []
all_back_data = []
all_both_data = []

# 逐文件生成三张图表
for file_idx, file in enumerate(csv_files, 1):
    file_short = os.path.splitext(file)[0]
    try:
        df = pd.read_csv(file, usecols=[1,4,10], encoding="gbk")
        df.columns = ["电流值","流量值","压差值"]
        df = clean_invalid_data(df)
        if df.empty: continue

        df["压差值滤波"] = df["压差值"].copy()
        df = filter_press_by_current(df, CONFIG["FILTER_WINDOW"])
        df["压差值"] = df["压差值滤波"]
        df = clean_invalid_data(df)
        if df.empty: continue

        curr_list = sorted(df["电流值"].unique())
        file_go = []
        file_back = []

        # 去程PQ图
        fig1, ax1 = plt.subplots(figsize=(12,6), dpi=100)
        for curr in curr_list:
            sub = df[df["电流值"]==curr].copy()
            sub_go = sub.iloc[:len(sub)//2].sort_values("流量值")
            x = sub_go["流量值"].abs()
            y = sub_go["压差值"].abs()
            if len(x)>=2:
                ax1.plot(x, y, lw=CONFIG["PERF_LINE_WIDTH"], color=LINE_COLOR)
                ax1.text(x.iloc[-1]+CONFIG["PERF_TEXT_OFFSET_X"], y.iloc[-1]+CONFIG["PERF_TEXT_OFFSET_Y"],
                        f"{round(curr)}mA", fontsize=CONFIG["PERF_CURR_TEXT_SIZE"], va="center")
                file_go.append((x,y))
        ax1.set_title(f"{file_short} 去程PQ", fontsize=CONFIG["PERF_TITLE_SIZE"], 
                      x=CONFIG["PERF_TITLE_OFFSET_X"], y=CONFIG["PERF_TITLE_OFFSET_Y"])
        ax1.set_xlabel("流量 L/min", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"])
        ax1.set_ylabel("压差 bar", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"])
        ax1.tick_params(labelsize=CONFIG["PERF_TICK_LABEL_SIZE"])
        ax1.set_xlim(right=70)
        if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
            ax1.spines['top'].set_visible(False)
            ax1.spines['right'].set_visible(False)
        plt.tight_layout()
        save_safe_plot(fig1, os.path.join("01_去程PQ", f"去程PQ_{file_idx}.png"))
        plt.close()
        all_go_data.append(file_go)

        # 回程PQ图
        fig2, ax2 = plt.subplots(figsize=(12,6), dpi=100)
        for curr in curr_list:
            sub = df[df["电流值"]==curr].copy()
            sub_back = sub.iloc[len(sub)//2:].sort_values("流量值")
            x = sub_back["流量值"].abs()
            y = sub_back["压差值"].abs()
            if len(x)>=2:
                ax2.plot(x, y, lw=CONFIG["PERF_LINE_WIDTH"], color=LINE_COLOR)
                ax2.text(x.iloc[-1]+CONFIG["PERF_TEXT_OFFSET_X"], y.iloc[-1]+CONFIG["PERF_TEXT_OFFSET_Y"],
                        f"{round(curr)}mA", fontsize=CONFIG["PERF_CURR_TEXT_SIZE"], va="center")
                file_back.append((x,y))
        ax2.set_title(f"{file_short} 回程PQ", fontsize=CONFIG["PERF_TITLE_SIZE"], 
                      x=CONFIG["PERF_TITLE_OFFSET_X"], y=CONFIG["PERF_TITLE_OFFSET_Y"])
        ax2.set_xlabel("流量 L/min", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"])
        ax2.set_ylabel("压差 bar", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"])
        ax2.tick_params(labelsize=CONFIG["PERF_TICK_LABEL_SIZE"])
        ax2.set_xlim(right=70)
        if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
            ax2.spines['top'].set_visible(False)
            ax2.spines['right'].set_visible(False)
        plt.tight_layout()
        save_safe_plot(fig2, os.path.join("02_回程PQ", f"回程PQ_{file_idx}.png"))
        plt.close()
        all_back_data.append(file_back)

        # 去程+回程合并图
        fig3, ax3 = plt.subplots(figsize=(12,6), dpi=100)
        for curr in curr_list:
            sub = df[df["电流值"]==curr].copy()
            sub_go = sub.iloc[:len(sub)//2].sort_values("流量值")
            xg = sub_go["流量值"].abs()
            yg = sub_go["压差值"].abs()
            if len(xg)>=2:
                ax3.plot(xg, yg, lw=CONFIG["PERF_LINE_WIDTH"], color=LINE_COLOR)
            sub_back = sub.iloc[len(sub)//2:].sort_values("流量值")
            xb = sub_back["流量值"].abs()
            yb = sub_back["压差值"].abs()
            if len(xb)>=2:
                ax3.plot(xb, yb, lw=CONFIG["PERF_LINE_WIDTH"], color=LINE_COLOR)
            if len(xg)>=2:
                ax3.text(xg.iloc[-1]+CONFIG["PERF_TEXT_OFFSET_X"], yg.iloc[-1]+CONFIG["PERF_TEXT_OFFSET_Y"],
                        f"{round(curr)}mA", fontsize=CONFIG["PERF_CURR_TEXT_SIZE"], va="center")
        ax3.set_title(f"{file_short} 去程+回程PQ", fontsize=CONFIG["PERF_TITLE_SIZE"], 
                      x=CONFIG["PERF_TITLE_OFFSET_X"], y=CONFIG["PERF_TITLE_OFFSET_Y"])
        ax3.set_xlabel("流量 L/min", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"])
        ax3.set_ylabel("压差 bar", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"])
        ax3.tick_params(labelsize=CONFIG["PERF_TICK_LABEL_SIZE"])
        ax3.set_xlim(right=70)
        if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
            ax3.spines['top'].set_visible(False)
            ax3.spines['right'].set_visible(False)
        plt.tight_layout()
        save_safe_plot(fig3, os.path.join("03_去程+回程PQ", f"去程回程对比_{file_idx}.png"))
        plt.close()
        all_both_data.append((file_go, file_back))

    except Exception as e:
        print(f"⚠️ {file} 图表生成失败：{str(e)}")
        plt.close('all')

# 生成总对比图
print("\n正在生成所有文件总对比图...")
# 去程总对比
fig_go, ax_go = plt.subplots(figsize=(15,8), dpi=150)
for file_data in all_go_data:
    for x,y in file_data:
        ax_go.plot(x, y, color=LINE_COLOR, lw=1.0)
ax_go.set_title("所有文件 去程PQ总对比", fontsize=20, pad=20)
ax_go.set_xlabel("流量 L/min", fontsize=14)
ax_go.set_ylabel("压差 bar", fontsize=14)
ax_go.set_xlim(right=70)
if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
    ax_go.spines['top'].set_visible(False)
    ax_go.spines['right'].set_visible(False)
plt.tight_layout()
save_safe_plot(fig_go, os.path.join("01_去程PQ", "所有文件_去程总对比.png"))
plt.close()

# 回程总对比
fig_back, ax_back = plt.subplots(figsize=(15,8), dpi=150)
for file_data in all_back_data:
    for x,y in file_data:
        ax_back.plot(x, y, color=LINE_COLOR, lw=1.0)
ax_back.set_title("所有文件 回程PQ总对比", fontsize=20, pad=20)
ax_back.set_xlabel("流量 L/min", fontsize=14)
ax_back.set_ylabel("压差 bar", fontsize=14)
ax_back.set_xlim(right=70)
if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
    ax_back.spines['top'].set_visible(False)
    ax_back.spines['right'].set_visible(False)
plt.tight_layout()
save_safe_plot(fig_back, os.path.join("02_回程PQ", "所有文件_回程总对比.png"))
plt.close()

# 去程+回程总对比
fig_both, ax_both = plt.subplots(figsize=(15,8), dpi=150)
for go_data, back_data in all_both_data:
    for x,y in go_data:
        ax_both.plot(x, y, color=LINE_COLOR, lw=1.0)
    for x,y in back_data:
        ax_both.plot(x, y, color=LINE_COLOR, lw=1.0)
ax_both.set_title("所有文件 去程+回程PQ总对比", fontsize=20, pad=20)
ax_both.set_xlabel("流量 L/min", fontsize=14)
ax_both.set_ylabel("压差 bar", fontsize=14)
ax_both.set_xlim(right=70)
if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
    ax_both.spines['top'].set_visible(False)
    ax_both.spines['right'].set_visible(False)
plt.tight_layout()
save_safe_plot(fig_both, os.path.join("03_去程+回程PQ", "所有文件_去程回程总对比.png"))
plt.close()

print("✅ 全部PQ图表生成完成！")

# ====================== 最终结果输出 ======================
print("\n" + "="*85)
print("📊 产品一致性检测最终结果（影响从大到小排序）")
print("="*85)

if file_score is not None and len(file_score) > 0:
    print(f"{'文件名':<25} {'影响占比':<10} {'缩放系数':<10} {'0mA 60L/min 修正值（原始→新）'}")
    print("-"*85)
    for f, score in file_score.items():
        rate = f"{impact_rate.get(f, 0):.1f}%"
        scale = f"×{scale_suggest.get(f, 1.0):.2f}"
        orig, new = scale_0mA60L.get(f, (0.0, 0.0))
        corr = f"{orig:.2f} → {new:.2f}"
        print(f"{f:<25} {rate:<10} {scale:<10} {corr}")
else:
    print("✅ 所有产品一致性均在合格范围内，无需调节")

print(f"\n🎉 程序执行完毕！所有检测结果已保存")
print(f"📂 结果保存路径：{OUTPUT_FOLDER}")

# 程序运行结束停留
input("\n按 回车键 退出程序...")