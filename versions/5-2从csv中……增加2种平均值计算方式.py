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

    # ====================== 【新增：一致性计算方式】 ======================
    # 可选模式：
    # 0 = 原始：(最大值+最小值)/2 作为均值（上下对称）
    # 1 = 所有文件真实平均值（上下偏差分开计算）
    # 2 = 指定基准文件作为均值（上下偏差分开计算）
    "CONSISTENCY_MODE": 0,

    # 当 CONSISTENCY_MODE = 2 时，填写基准文件全称（带.csv）
    "BASE_FILE_NAME": "2LHO-32.csv",

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
    "LABEL_POS_0mA": -16,                       # 0mA数据标签垂直偏移（负数向下）
    "LABEL_POS_OTHER": 9,                       # 其他电流标签垂直偏移（正数向上）
    "CON_LABEL_SIZE": 15,                       # 误差标注字体大小
    "CON_CURR_TEXT_SIZE": 20,                   # 电流值标注字体大小
    
    # ====================== 【功能运行总开关】 ======================
    "RUN_PERFORMANCE_PLOTS": True,       # 生成单文件性能图表
    "RUN_PQ_COMPARE_PLOT": True,         # 生成所有产品PQ总对比散点图
    "RUN_HYSTERESIS_PLOT": True,         # 生成PNG格式迟滞图
    "RUN_CONSISTENCY_PLOT": True,        # 生成一致性分析图
    "RUN_AVG_PRESSURE_PLOT": True,       # 生成压差平均值图
    "RUN_EXCEL_CHARTS": True,            # 在Excel中嵌入迟滞+一致性图表
    "RUN_EXCEL_REPORT": True,            # 生成完整Excel检测报告
    "RUN_SINGLE_GOBACK_PQ": True,        # 生成单文件去程/回程/合并PQ图
    "RUN_ALL_GOBACK_SUMMARY": True,      # 生成全体去程回程总对比图
}

# ---------------------- 图表颜色配置（固定标准，无需修改） ----------------------
COLOR_MAP = {0: '#70AD47', 1: '#00B0F0', 2: '#FFC000', 3: '#FF0000', 4: '#C00000'}
COLOR_EXCEL = ["70AD47", "00B0F0", "FFC000", "FF0000", "C00000"]
FILE_COLORS = {
    0: '#00B0F0', 1: '#FFC000', 2: '#70AD47', 3: '#FF0000', 4: '#C00000',
    5: '#9933FF', 6: '#00FF00', 7: '#00FFFF', 8: '#FF9900', 9: '#999999'
}

# ======================================
# ====================== 程序核心功能区 ======================
# ======================================

# ====================== 输出文件夹自动创建（防重名） ======================
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

# ====================== 图表安全保存函数（防覆盖） ======================
def save_safe_plot(fig, filename, dpi=150, bbox_inches='tight'):
    path = os.path.join(OUTPUT_FOLDER, filename)
    base, ext = os.path.splitext(path)
    counter = 1
    while os.path.exists(path):
        path = f"{base}({counter}){ext}"
        counter += 1
    fig.savefig(path, dpi=dpi, bbox_inches=bbox_inches)
    plt.close(fig)
    return path

# ====================== 通用工具函数 ======================
def get_closest_index(arr, target):
    return np.argmin(np.abs(arr - target))

def safe_percent(numerator, denominator, default=0):
    try:
        return round((numerator / denominator) * 100, 1)
    except ZeroDivisionError:
        return default

def filter_press_by_current(df, window=201):
    if not CONFIG["FILTER_ON"]:
        return df
    df_out = df.copy()
    for curr, group in df.groupby("电流值"):
        filtered_press = group["压差值"].rolling(window=window, center=True, min_periods=1).mean()
        df_out.loc[group.index, "压差值滤波"] = filtered_press
    return df_out

def clean_invalid_data(df):
    df = df.replace([False, None, np.inf, -np.inf], np.nan)
    df = df.dropna(subset=["流量值", "压差值"])
    df["流量值"] = pd.to_numeric(df["流量值"], errors="coerce")
    df["压差值"] = pd.to_numeric(df["压差值"], errors="coerce")
    return df

# ====================== 加权一致性影响分析（核心算法） ======================
def analyze_consistency_influence_final(df_valid):
    influence = []
    KEY_CURRENTS = CONFIG["KEY_CURRENTS"]
    KEY_FLOW = CONFIG["KEY_FLOW"]

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

# ====================== 提取0mA 60L/min真实压差值 ======================
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

# ====================== 一致性计算核心函数 ======================
def calculate_consistency(df_valid):
    con_data = []
    mode = CONFIG["CONSISTENCY_MODE"]
    base_file = CONFIG["BASE_FILE_NAME"]

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

        con_data.append([
            round(curr, 2), round(rf, 1), part,
            round(maxv, 2), max_file,
            round(minv, 2), min_file,
            round(mean_val, 2),
            round(pos_diff, 2), round(pos_pct, 2),
            round(neg_diff, 2), round(neg_pct, 2)
        ])

    df_con = pd.DataFrame(con_data, columns=[
        "电流值", "参考流量值", "数据分区",
        "最大值", "最大值文件名",
        "最小值", "最小值文件名",
        "平均值",
        "正偏差bar", "正偏差%",
        "负偏差bar", "负偏差%"
    ])
    return df_con

# ====================== 主程序入口 ======================
print("="*70)
print("🚀 PQ 产品一致性检测分析程序 启动成功")
print("="*70)

script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# ---------------------- 步骤1：扫描数据文件 ----------------------
print("\n【1/12】正在扫描当前目录的CSV数据文件...")
csv_files = glob.glob("*.csv")
print(f"✅ 扫描完成：共找到 {len(csv_files)} 个有效数据文件")

raw_data = []
pq_comparison_data = []
file_full_data = {}  # 全局存储所有文件已清洗数据，修复绘图bug

# ---------------------- 步骤2：读取并处理数据 ----------------------
print("\n【2/12】正在读取数据并执行滤波处理...")
for file in csv_files:
    try:
        df = pd.read_csv(file, usecols=[1,4,10], encoding="gbk")
        df.columns = ["电流值","流量值","压差值"]
    except:
        continue

    df = clean_invalid_data(df)
    if df.empty:
        continue

    df["流量值原始"] = df["流量值"].copy()
    df["压差值原始"] = df["压差值"].copy()
    df["压差值滤波"] = df["压差值"].copy()
    df = filter_press_by_current(df, CONFIG["FILTER_WINDOW"])

    df = clean_invalid_data(df)
    if df.empty:
        continue

    df["压差值"] = df["压差值滤波"]
    file_short = os.path.splitext(file)[0]
    file_full_data[file] = df.copy()  # 存储完整数据，供绘图使用

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

    for I, g in df.groupby("电流值"):
        I = round(I, 2)
        mid = len(g) // 2
        for part, part_data in [("前50%", g.iloc[:mid]), ("后50%", g.iloc[mid:])]:
            part_data = clean_invalid_data(part_data)
            if part_data.empty:
                continue
            f_vals = part_data["流量值"].abs().values
            p_vals = part_data["压差值"].abs().values
            for ref_flow in CONFIG["REF_FLOW"]:
                idx = get_closest_index(f_vals, ref_flow)
                real_flow = round(abs(f_vals[idx]),4)
                real_press = round(abs(p_vals[idx]),4)
                diff = round(abs(real_flow - ref_flow),4)
                check = "TRUE" if diff <= CONFIG["TOLERANCE"] else "FALSE"
                raw_data.append([file, I, part, ref_flow, real_press, real_flow, diff, check])

df_raw = pd.DataFrame(raw_data, columns=[
    "数据源文件","电流值","数据分区","参考流量值","实际压差值","实际流量值","差值数据","校验结果"
])

# ---------------------- 步骤3：筛选全文件有效数据 ----------------------
print("\n【3/12】正在筛选所有文件均合格的有效检测点位...")
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

plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

# ---------------------- 步骤6：生成单文件性能图表（已修复） ----------------------
print("\n【6/12】正在生成单产品性能曲线图...")
if CONFIG["RUN_PERFORMANCE_PLOTS"]:
    file_list = df_valid["数据源文件"].unique()
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
    
    print("✅ 单产品性能图表生成完成")
else:
    print("ℹ️ 已跳过：单产品性能图表")

# ---------------------- 步骤7：生成PQ总对比散点图 ----------------------
print("\n【7/12】正在生成所有产品PQ总对比散点图...")
if CONFIG["RUN_PQ_COMPARE_PLOT"] and pq_comparison_data:
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
                label = f if f not in plotted_files else ""
                if label: plotted_files.add(f)
                ax.plot(cdat["flow"], cdat["press"], marker='o', ms=3, linewidth=0.6,
                        color=file_color_map[f], label=label, alpha=0.8)

    if unique_currents:
        first_file = unique_files[0]
        first_data = [d for d in pq_comparison_data if d["file"] == first_file]
        for curr in unique_currents:
            cdat = next((d for d in first_data if d["current"] == curr), None)
            if cdat:
                x_last = cdat["flow"][-1]
                y_last = cdat["press"][-1]
                ax.text(x_last + CONFIG["PERF_TEXT_OFFSET_X"], y_last, f"{curr}mA",
                        fontsize=CONFIG["CON_CURR_TEXT_SIZE"])

    ax.set_title("所有产品滤波后前50% PQ曲线对比", fontsize=16, weight='bold')
    ax.set_xlabel("流量 L/min", fontsize=14)
    ax.set_ylabel("压差 bar", fontsize=14)
    ax.grid(alpha=0.3)
    ax.set_xlim(0, max(CONFIG["REF_FLOW"]) + 10)
    if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
    
    ncol = min(3, max(1, len(unique_files) // 12))
    ax.legend(loc="center left", bbox_to_anchor=(1.03, 0.5), ncol=ncol, fontsize=10)
    plt.tight_layout(rect=[0, 0, 0.87, 1])
    save_safe_plot(fig, "所有产品PQ曲线对比图.png", dpi=300)
    print("✅ PQ总对比散点图生成完成")
else:
    print("ℹ️ 已跳过：PQ总对比散点图")

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

# ---------------------- 步骤9：新版一致性计算 ----------------------
print("\n【9/12】正在计算产品一致性数据...")
df_con = calculate_consistency(df_valid)
print("✅ 一致性数据计算完成")

# ---------------------- 步骤10：生成Excel检测报告 ----------------------
print("\n【10/12】正在生成Excel完整检测报告...")
if CONFIG["RUN_EXCEL_REPORT"]:
    wb = Workbook()
    if "Sheet" in wb.sheetnames: del wb["Sheet"]

    ws1 = wb.create_sheet("原始检测数据")
    ws1.append(df_raw.columns.tolist())
    for _,r in df_raw.iterrows(): ws1.append(r.tolist())

    ws2 = wb.create_sheet("产品迟滞")
    ws2.append(df_hys.columns.tolist())
    for r in hys_data: ws2.append(r)
    
    plot_hys = df_hys[df_hys["参考流量值"]==CONFIG["PLOT_FLOW"]].copy().sort_values(["数据源文件","电流值"])
    ws2.cell(1,8,"绘图文件名"), ws2.cell(1,9,"绘图电流"), ws2.cell(1,10,"绘图迟滞")
    for i,(_,r) in enumerate(plot_hys.iterrows(),2):
        ws2.cell(i,8,r["数据源文件"]), ws2.cell(i,9,r["电流值"]), ws2.cell(i,10,r["迟滞"])

    if CONFIG["RUN_EXCEL_CHARTS"]:
        chart_hys = ScatterChart()
        chart_hys.scatterStyle = "lineMarker"
        chart_hys.title = "P-Q滞环@20L/min"
        chart_hys.x_axis.title = "电流值"
        chart_hys.y_axis.title = "迟滞/bar"
        
        for prod in plot_hys["数据源文件"].unique():
            sub_rows = plot_hys[plot_hys["数据源文件"]==prod]
            sr = sub_rows.index[0]+2
            er = sub_rows.index[-1]+2
            x = Reference(ws2, 9, sr, 9, er)
            y = Reference(ws2, 10, sr, 10, er)
            exceed = sum(sub_rows["迟滞"] > CONFIG["STD_LIMIT"])
            color = COLOR_EXCEL[min(exceed,4)]
            ser = Series(y,x,title=prod)
            ser.marker = Marker(size=3)
            ser.graphicalProperties.line = LineProperties(solidFill=color, w=10000)
            chart_hys.series.append(ser)
        ws2.add_chart(chart_hys, "L2")

    ws3 = wb.create_sheet("一致性对比")
    ws3.append(df_con.columns.tolist())
    for _,r in df_con.iterrows(): ws3.append(r.tolist())

    def draw_con_chart(part, title, col, pos):
        data = df_con[df_con["数据分区"]==part].sort_values(["电流值","参考流量值"])
        headers = ["电流值","参考流量值","平均值","正偏差bar","正偏差%","负偏差bar","负偏差%","标签"]
        for i,h in enumerate(headers): ws3.cell(1,col+i,h)
        
        for i,(_,r) in enumerate(data.iterrows(),2):
            lab = f"+{r['正偏差%']}% / -{r['负偏差%']}%"
            ws3.cell(i,col+0,r["电流值"]), ws3.cell(i,col+1,r["参考流量值"]), ws3.cell(i,col+2,r["平均值"])
            ws3.cell(i,col+3,r["正偏差bar"]), ws3.cell(i,col+4,r["正偏差%"]), ws3.cell(i,col+5,r["负偏差bar"])
            ws3.cell(i,col+6,r["负偏差%"]), ws3.cell(i,col+7,lab)
        
        ch = ScatterChart()
        ch.scatterStyle = "lineMarker"
        ch.title = f"一致性 {title}"
        ch.x_axis.title = "流量/Lpm"
        ch.y_axis.title = "压差/bar"
        ch.x_axis.majorGridlines = None
        ch.y_axis.majorGridlines = None
        
        for curr in sorted(data["电流值"].unique()):
            curr_data = data[data["电流值"]==curr]
            sr = curr_data.index[0]+2
            er = curr_data.index[-1]+2
            x = Reference(ws3, col+1, sr, col+1, er)
            y = Reference(ws3, col+2, sr, col+2, er)
            label_ref = Reference(ws3, col+7, sr, col+7, er)
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
        draw_con_chart("后50%", "后50%", 30, "AQ20")

    ws4 = wb.create_sheet("一致性影响分析(加权)")
    ws4.append(["文件名","加权偏差总分","影响占比(%)","建议缩放系数P","0mA 60L/min 修正说明","调节建议"])
    if file_score is not None:
        for f, score in file_score.items():
            orig, new = scale_0mA60L.get(f, (0,0))
            ws4.append([f, round(score,2), impact_rate.get(f,0), scale_suggest.get(f,1.0),
                        f"{orig}→{new}", f"全量程×{scale_suggest.get(f,1.0)}"])

    excel_path = os.path.join(OUTPUT_FOLDER, CONFIG["OUTPUT_EXCEL"])
    wb.save(excel_path)
    print("✅ Excel检测报告生成完成")
else:
    print("ℹ️ 已跳过：生成完整Excel检测报告")

# ---------------------- 步骤11：生成PNG分析图表 ----------------------
print("\n【11/12】正在生成迟滞、一致性、压差平均值图表...")
if CONFIG["RUN_HYSTERESIS_PLOT"]:
    fig, ax = plt.subplots(figsize=(15,6), dpi=100)
    for prod in plot_hys["数据源文件"].unique():
        sub = plot_hys[plot_hys["数据源文件"]==prod]
        exceed = sum(sub["迟滞"] > CONFIG["STD_LIMIT"])
        ax.plot(sub["电流值"], sub["迟滞"], marker='o', ms=3, lw=1, 
                color=COLOR_MAP[min(exceed,4)], label=prod)
    ax.axhline(CONFIG["STD_LIMIT"], color='r', linestyle='--', label='标准线')
    ax.set_title("P-Q滞环@20L/min")
    ax.set_xlabel("电流值"), ax.set_ylabel("迟滞/bar"), ax.grid(axis='y')
    ncol = min(3, len(plot_hys["数据源文件"].unique())//15+1)
    ax.legend(loc="center left", bbox_to_anchor=(1.02,0.5), ncol=ncol, fontsize=8)
    if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
        ax.spines['top'].set_visible(False), ax.spines['right'].set_visible(False)
    plt.tight_layout()
    save_safe_plot(fig, "迟滞图.png")
    print("✅ 迟滞图生成完成")

if CONFIG["RUN_CONSISTENCY_PLOT"]:
    con_plot = df_con[df_con["数据分区"]=="前50%"].copy()
    fig, ax = plt.subplots(figsize=(15,6), dpi=100)
    for curr in sorted(con_plot["电流值"].unique()):
        sub = con_plot[con_plot["电流值"]==curr].sort_values("参考流量值")
        ax.plot(sub["参考流量值"], sub["平均值"], marker='o', ms=4, color='#00B0F0', lw=CONFIG["CON_LINE_WIDTH"])
        ax.text(sub["参考流量值"].iloc[-1]+CONFIG["PERF_TEXT_OFFSET_X"], sub["平均值"].iloc[-1],
                f"{curr}mA", fontsize=CONFIG["CON_CURR_TEXT_SIZE"])
        for x,y,pp,np in zip(sub["参考流量值"], sub["平均值"], sub["正偏差%"], sub["负偏差%"]):
            offset = CONFIG["LABEL_POS_0mA"] if curr==0 else CONFIG["LABEL_POS_OTHER"]
            ax.annotate(f"+{pp}% / -{np}%", (x,y), xytext=(0,offset), textcoords='offset points',
                        ha='center', fontsize=CONFIG["CON_LABEL_SIZE"])
    ax.set_title(CONFIG["CON_TITLE"], fontsize=CONFIG["CON_TITLE_SIZE"], x=CONFIG["CON_TITLE_OFFSET_X"], y=CONFIG["CON_TITLE_OFFSET_Y"])
    ax.set_xlabel("流量/Lpm", fontsize=CONFIG["CON_AXIS_LABEL_SIZE"]), ax.set_ylabel("压差 bar", fontsize=CONFIG["CON_AXIS_LABEL_SIZE"])
    ax.set_xlim(right=ax.get_xlim()[1]+CONFIG["X_AXIS_EXTEND"])
    if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
        ax.spines['top'].set_visible(False), ax.spines['right'].set_visible(False)
    plt.tight_layout()
    save_safe_plot(fig, "一致性图.png")
    print("✅ 一致性图生成完成")

if CONFIG["RUN_AVG_PRESSURE_PLOT"]:
    con_plot = df_con[df_con["数据分区"]=="前50%"].copy()
    fig, ax = plt.subplots(figsize=(15,6), dpi=100)
    for curr in sorted(con_plot["电流值"].unique()):
        sub = con_plot[con_plot["电流值"]==curr].sort_values("参考流量值")
        ax.plot(sub["参考流量值"], sub["平均值"], marker='o', ms=4, color='#00B0F0', lw=CONFIG["CON_LINE_WIDTH"])
        ax.text(sub["参考流量值"].iloc[-1]+CONFIG["PERF_TEXT_OFFSET_X"], sub["平均值"].iloc[-1],
                f"{curr}mA", fontsize=CONFIG["CON_CURR_TEXT_SIZE"])
        for x,y,avg,pp,np in zip(sub["参考流量值"], sub["平均值"], sub["平均值"], sub["正偏差%"], sub["负偏差%"]):
            offset = CONFIG["LABEL_POS_0mA"] if curr==0 else CONFIG["LABEL_POS_OTHER"]
            ax.annotate(f"{avg:.1f} | +{pp}% / -{np}%", (x,y), xytext=(0,offset), textcoords='offset points',
                        ha='center', fontsize=CONFIG["CON_LABEL_SIZE"])
    ax.set_title("压差平均值", fontsize=CONFIG["CON_TITLE_SIZE"], x=CONFIG["CON_TITLE_OFFSET_X"], y=CONFIG["CON_TITLE_OFFSET_Y"])
    ax.set_xlabel("流量/Lpm", fontsize=CONFIG["CON_AXIS_LABEL_SIZE"]), ax.set_ylabel("压差 bar", fontsize=CONFIG["CON_AXIS_LABEL_SIZE"])
    ax.set_xlim(right=ax.get_xlim()[1]+CONFIG["X_AXIS_EXTEND"])
    if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
        ax.spines['top'].set_visible(False), ax.spines['right'].set_visible(False)
    plt.tight_layout()
    save_safe_plot(fig, "压差平均值图.png")
    print("✅ 压差平均值图生成完成")

# ---------------------- 步骤12：生成去程/回程PQ图表（已修复） ----------------------
print("\n【12/12】正在生成去程、回程、合并PQ曲线图...")
if CONFIG["RUN_SINGLE_GOBACK_PQ"]:
    folder_go = os.path.join(OUTPUT_FOLDER, "01_去程PQ")
    folder_back = os.path.join(OUTPUT_FOLDER, "02_回程PQ")
    folder_both = os.path.join(OUTPUT_FOLDER, "03_去程+回程PQ")
    for d in [folder_go, folder_back, folder_both]:
        os.makedirs(d, exist_ok=True)

    all_go, all_back, all_both = [], [], []
    LINE_COLOR = "#00B0F0"

    for idx, file in enumerate(csv_files, 1):
        if file not in file_full_data:
            continue
        df = file_full_data[file].copy()
        file_short = os.path.splitext(file)[0]
        curr_list = sorted(df["电流值"].unique())
        go_list, back_list = [], []

        # 去程
        fig1, ax1 = plt.subplots(figsize=(12,6), dpi=100)
        for curr in curr_list:
            sub = df[df["电流值"]==curr]
            go = sub.iloc[:len(sub)//2].sort_values("流量值")
            x, y = go["流量值"].abs(), go["压差值"].abs()
            if len(x)>=2:
                ax1.plot(x, y, lw=CONFIG["PERF_LINE_WIDTH"], color=LINE_COLOR)
                ax1.text(x.iloc[-1]+CONFIG["PERF_TEXT_OFFSET_X"], y.iloc[-1], f"{round(curr)}mA",
                        fontsize=CONFIG["PERF_CURR_TEXT_SIZE"], va="center")
                go_list.append((x,y))
        ax1.set_title(f"{file_short} 去程PQ", fontsize=CONFIG["PERF_TITLE_SIZE"],
                      x=CONFIG["PERF_TITLE_OFFSET_X"], y=CONFIG["PERF_TITLE_OFFSET_Y"])
        ax1.set_xlabel("流量 L/min", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"])
        ax1.set_ylabel("压差 bar", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"])
        ax1.set_xlim(right=70)
        if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
            ax1.spines['top'].set_visible(False), ax1.spines['right'].set_visible(False)
        plt.tight_layout()
        save_safe_plot(fig1, os.path.join("01_去程PQ", f"去程PQ_{idx}.png"))
        all_go.append(go_list)

        # 回程
        fig2, ax2 = plt.subplots(figsize=(12,6), dpi=100)
        for curr in curr_list:
            sub = df[df["电流值"]==curr]
            back = sub.iloc[len(sub)//2:].sort_values("流量值")
            x, y = back["流量值"].abs(), back["压差值"].abs()
            if len(x)>=2:
                ax2.plot(x, y, lw=CONFIG["PERF_LINE_WIDTH"], color=LINE_COLOR)
                ax2.text(x.iloc[-1]+CONFIG["PERF_TEXT_OFFSET_X"], y.iloc[-1], f"{round(curr)}mA",
                        fontsize=CONFIG["PERF_CURR_TEXT_SIZE"], va="center")
                back_list.append((x,y))
        ax2.set_title(f"{file_short} 回程PQ", fontsize=CONFIG["PERF_TITLE_SIZE"],
                      x=CONFIG["PERF_TITLE_OFFSET_X"], y=CONFIG["PERF_TITLE_OFFSET_Y"])
        ax2.set_xlabel("流量 L/min", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"])
        ax2.set_ylabel("压差 bar", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"])
        ax2.set_xlim(right=70)
        if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
            ax2.spines['top'].set_visible(False), ax2.spines['right'].set_visible(False)
        plt.tight_layout()
        save_safe_plot(fig2, os.path.join("02_回程PQ", f"回程PQ_{idx}.png"))
        all_back.append(back_list)

        # 合并
        fig3, ax3 = plt.subplots(figsize=(12,6), dpi=100)
        for curr in curr_list:
            sub = df[df["电流值"]==curr]
            go = sub.iloc[:len(sub)//2].sort_values("流量值")
            back = sub.iloc[len(sub)//2:].sort_values("流量值")
            if len(go)>=2:
                ax3.plot(go["流量值"].abs(), go["压差值"].abs(), lw=CONFIG["PERF_LINE_WIDTH"], color=LINE_COLOR)
            if len(back)>=2:
                ax3.plot(back["流量值"].abs(), back["压差值"].abs(), lw=CONFIG["PERF_LINE_WIDTH"], color=LINE_COLOR)
            if len(go)>=2:
                ax3.text(go["流量值"].abs().iloc[-1]+CONFIG["PERF_TEXT_OFFSET_X"], go["压差值"].abs().iloc[-1],
                        f"{round(curr)}mA", fontsize=CONFIG["PERF_CURR_TEXT_SIZE"], va="center")
        ax3.set_title(f"{file_short} 去程+回程PQ", fontsize=CONFIG["PERF_TITLE_SIZE"],
                      x=CONFIG["PERF_TITLE_OFFSET_X"], y=CONFIG["PERF_TITLE_OFFSET_Y"])
        ax3.set_xlabel("流量 L/min", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"])
        ax3.set_ylabel("压差 bar", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"])
        ax3.set_xlim(right=70)
        if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
            ax3.spines['top'].set_visible(False), ax3.spines['right'].set_visible(False)
        plt.tight_layout()
        save_safe_plot(fig3, os.path.join("03_去程+回程PQ", f"去程回程对比_{idx}.png"))
        all_both.append((go_list, back_list))

    print("✅ 单文件去程/回程/合并PQ图生成完成")

    if CONFIG["RUN_ALL_GOBACK_SUMMARY"]:
        fig, ax = plt.subplots(figsize=(15,8), dpi=150)
        for d in all_go:
            for x,y in d: ax.plot(x,y,color=LINE_COLOR,lw=1)
        ax.set_title("所有文件 去程PQ总对比",fontsize=20,pad=20)
        ax.set_xlabel("流量 L/min"), ax.set_ylabel("压差 bar")
        ax.set_xlim(right=70)
        ax.spines[['top','right']].set_visible(False)
        plt.tight_layout()
        save_safe_plot(fig, os.path.join("01_去程PQ","所有文件_去程总对比.png"))

        fig, ax = plt.subplots(figsize=(15,8), dpi=150)
        for d in all_back:
            for x,y in d: ax.plot(x,y,color=LINE_COLOR,lw=1)
        ax.set_title("所有文件 回程PQ总对比",fontsize=20,pad=20)
        ax.set_xlabel("流量 L/min"), ax.set_ylabel("压差 bar")
        ax.set_xlim(right=70)
        ax.spines[['top','right']].set_visible(False)
        plt.tight_layout()
        save_safe_plot(fig, os.path.join("02_回程PQ","所有文件_回程总对比.png"))

        fig, ax = plt.subplots(figsize=(15,8), dpi=150)
        for go_d, back_d in all_both:
            for x,y in go_d: ax.plot(x,y,color=LINE_COLOR,lw=1)
            for x,y in back_d: ax.plot(x,y,color=LINE_COLOR,lw=1)
        ax.set_title("所有文件 去程+回程PQ总对比",fontsize=20,pad=20)
        ax.set_xlabel("流量 L/min"), ax.set_ylabel("压差 bar")
        ax.set_xlim(right=70)
        ax.spines[['top','right']].set_visible(False)
        plt.tight_layout()
        save_safe_plot(fig, os.path.join("03_去程+回程PQ","所有文件_去程回程总对比.png"))
        print("✅ 全体去程回程总对比图生成完成！")

# ====================== 最终结果输出 ======================
print("\n" + "="*85)
print("📊 产品一致性检测最终结果（影响从大到小排序）")
print("="*85)

if file_score is not None and len(file_score) > 0:
    print(f"{'文件名':<25} {'影响占比':<10} {'缩放系数':<10} {'0mA 60L/min 修正值（原始→新）'}")
    print("-"*85)
    for f, _ in file_score.items():
        rate = f"{impact_rate.get(f, 0):.1f}%"
        scale = f"×{scale_suggest.get(f, 1.0):.2f}"
        orig, new = scale_0mA60L.get(f, (0.0, 0.0))
        print(f"{f:<25} {rate:<10} {scale:<10} {orig:.2f} → {new:.2f}")
else:
    print("✅ 所有产品一致性均在合格范围内，无需调节")

print(f"\n🎉 程序执行完毕！所有检测结果已保存")
print(f"📂 结果保存路径：{OUTPUT_FOLDER}")
input("\n按 回车键 退出程序...")