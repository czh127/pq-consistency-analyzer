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
# ====================== 运行规则约定 ======================
# 功能开关统一规则：0=关闭功能  1=开启功能
# 【强依赖关系】不满足前置条件时，对应功能自动静默跳过，不报错：
# 1. RUN_ALL_GOBACK_SUMMARY（全体去回程总对比图）→ 依赖 RUN_SINGLE_GOBACK_PQ=1
# 2. RUN_EXCEL_CHARTS（Excel内嵌图表）→ 依赖 RUN_EXCEL_REPORT=1
# 其余开关相互独立，无依赖
# 提示：开启全部图表会增加运行时长与文件体积，按需开启
# ======================================
warnings.filterwarnings('ignore')  # 忽略无关警告，保持控制台整洁
CONFIG = {
    # ---------------------- 输出文件基础配置 ----------------------
    "OUTPUT_EXCEL": "产品检测结果总表.xlsx",  # Excel报告主文件名，重名自动追加序号防覆盖
    "ENCODING": "gbk",                      # CSV读取编码：Windows默认gbk；Linux/Mac请改为 utf-8

    # ---------------------- CSV数据列索引配置（索引从0开始） ----------------------
    # 重要：列索引与实际CSV列不匹配会造成数据读取失败，修改后请验证
    "COL_CURRENT": 1,   # 电流值列 (单位: mA)
    "COL_FLOW": 4,      # 流量值列 (单位: L/min)
    "COL_PRESS": 10,    # 压差值列 (单位: bar)

    # ---------------------- 标准检测流量点 ----------------------
    "REF_FLOW": [5.0, 10.0, 20.0, 40.0, 60.0],  # 目标标准流量(L/min)，程序自动就近匹配原始数据

    # ---------------------- 迟滞 & 流量匹配容错配置 ----------------------
    "PLOT_FLOW": 20.0,                # 迟滞图表固定采用的流量点 (L/min)
    "STD_LIMIT": 10.0,                # 迟滞判定合格上限 (bar)，超过则判定超标
    "STD_LINE1": {"value": 10.0, "show": 1},  # 迟滞图参考标准线：数值+显示开关(0隐藏/1显示)
    "TOLERANCE": 1.00,                # 流量匹配最大允许误差(L/min)，超出则该点位判定无效

    # ---------------------- 一致性计算模式 & 判定参数 ----------------------
    "CONSISTENCY_LIMIT_PCT": 5.0,   # 一致性参考阈值(±%)，作为人工判定依据
    # 一致性均值计算模式（关键：直接影响图表标签样式、偏差计算逻辑）
    # 模式0：(最大值+最小值)/2 对称均值 → 【当前需求：全部标签显示 数值±偏差bar】
    # 模式1：全部文件真实算术平均值 → 通用统计分析、常规检测推荐
    # 模式2：以指定单个文件为基准均值 → 样机对标、单品参照场景使用
    "CONSISTENCY_MODE": 0,
    # 仅模式2生效：基准文件名（必须包含后缀 .csv，文件不存在则该工况跳过）
    "BASE_FILE_NAME": "产品1.csv",

    # ---------------------- 加权一致性分析参数 ----------------------
    "KEY_FLOW": 20.0,                  # 核心重点流量点 (L/min)，加权系数更高
    "HIGH_FLOW": 40.0,                 # 高流量参考点 (L/min)
    "KEY_CURRENTS": [0, 300, 600, 900],# 重点分析电流序列 (mA)，参与加权计算

    # ---------------------- 滑动平均滤波配置 ----------------------
    "FILTER_WINDOW": 201,    # 滤波窗口：强制要求奇数；数值越大曲线越平滑，细节丢失越多
    "FILTER_ON": True,       # 滤波总开关：True开启 / False关闭，直接使用原始数据

    # ---------------------- 图表通用全局样式 ----------------------
    "X_AXIS_EXTEND": 8,                # X轴右侧留白延长长度 (L/min)，防止文字超出边界
    "HIDE_TOP_RIGHT_BORDER": True,     # 隐藏图表上、右侧边框：True简洁风格 / False完整边框

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

    # ---------------------- 一致性/均值图 数据标签位置样式 ----------------------
    "LABEL_POS_0mA": -16,        # 0mA标签垂直偏移(像素)：负数向下，标签重叠可微调
    "LABEL_POS_OTHER": 9,        # 其余电流标签垂直偏移(像素)：正数向上
    "CON_LABEL_SIZE": 15,        # 偏差数值标签字号
    "CON_CURR_TEXT_SIZE": 20,    # 曲线末端电流标注字号
    
    # ====================== 【功能运行总开关 0=关闭 1=开启】 ======================
    # 独立开关：生成单文件性能图表
    "RUN_PERFORMANCE_PLOTS": 1,
    # 独立开关：生成所有产品PQ总对比散点图
    "RUN_PQ_COMPARE_PLOT": 1,
    # 独立开关：生成PNG格式迟滞图
    "RUN_HYSTERESIS_PLOT": 1,
    # 独立开关：生成一致性分析图
    "RUN_CONSISTENCY_PLOT": 1,
    # 独立开关：生成压差平均值图
    "RUN_AVG_PRESSURE_PLOT": 1,
    # 依赖开关：Excel内嵌图表 → 必须 RUN_EXCEL_REPORT=1 才生效
    "RUN_EXCEL_CHARTS": 1,
    # 独立开关：生成完整Excel检测报告
    "RUN_EXCEL_REPORT": 1,
    # 基础开关：生成单文件去程/回程/合并PQ图
    "RUN_SINGLE_GOBACK_PQ": 1,
    # 依赖开关：全体去程回程总对比图 → 必须 RUN_SINGLE_GOBACK_PQ=1 才生效
    "RUN_ALL_GOBACK_SUMMARY": 1,

    # ---------------------- 扩展通用配置（新增，便于后期维护） ----------------------
    "PLOT_DPI": 150,                # 基础图表DPI，越高图片越清晰、文件越大
    "SUFFIX_AUTO_NUM": True,        # 文件/文件夹重名时自动加序号，防止覆盖
    "PRINT_DEBUG_LOG": True         # 控制台打印运行日志，调试可开；正式运行可关
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
    df = df.copy()
    df = df.replace([np.inf, -np.inf], np.nan)
    df = df.dropna()
    return df

# ====================== 加权一致性影响分析（核心算法） ======================
def analyze_consistency_influence_final(df_valid):
    influence = []
    KEY_CURRENTS = CONFIG["KEY_CURRENTS"]
    KEY_FLOW = CONFIG["KEY_FLOW"]
    HIGH_FLOW = CONFIG["HIGH_FLOW"]

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

# ====================== 【新版：一致性计算核心函数】 ======================
def calculate_consistency(df_valid):
    con_data = []
    mode = CONFIG["CONSISTENCY_MODE"]
    base_file = CONFIG["BASE_FILE_NAME"]

    for (curr, rf, part), group in df_valid.groupby(["电流值", "参考流量值", "数据分区"]):
        vals = group["实际压差值"].dropna()
        files = group["数据源文件"]

        if len(vals) < 1:
            continue

        # 模式0：原始 (max+min)/2
        if mode == 0:
            maxv = vals.max()
            minv = vals.min()
            mean_val = (maxv + minv) / 2
            pos_diff = maxv - mean_val
            neg_diff = mean_val - minv
            max_file = group.loc[vals.idxmax(), "数据源文件"]
            min_file = group.loc[vals.idxmin(), "数据源文件"]

        # 模式1：全体真实平均值
        elif mode == 1:
            mean_val = vals.mean()
            maxv = vals.max()
            minv = vals.min()
            pos_diff = maxv - mean_val
            neg_diff = mean_val - minv
            max_file = group.loc[vals.idxmax(), "数据源文件"]
            min_file = group.loc[vals.idxmin(), "数据源文件"]

        # 模式2：指定文件为基准
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

# ====================== 📊 打印本次使用的平均值计算方式 ======================
mode = CONFIG["CONSISTENCY_MODE"]
print("\n📏 【平均值计算方式】")
if mode == 0:
    print("   模式 0 → (最大值 + 最小值) / 2  对称均值 ")
elif mode == 1:
    print("   模式 1 → 所有文件真实平均值  🎯 推荐")
elif mode == 2:
    print(f"   模式 2 → 以基准文件 [{CONFIG['BASE_FILE_NAME']}] 为均值")
else:
    print("   模式未知 → 使用默认真实平均值")
print("="*70)

script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# ---------------------- 步骤1：扫描数据文件 ----------------------
print("\n📂 【1/12】正在扫描当前目录的CSV数据文件...")
csv_files = glob.glob("*.csv")
print(f"✅ 扫描完成：共找到 {len(csv_files)} 个有效数据文件")

raw_data = []
pq_comparison_data = []
file_cleaned_data = {}

# ---------------------- 步骤2：读取并处理数据 ----------------------
print("\n🔍 【2/12】正在读取数据并执行滤波处理...")
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
            for ref in CONFIG["REF_FLOW"]:
                idx = get_closest_index(f_vals, ref)
                real_flow = round(abs(f_vals[idx]),4)
                real_press = round(abs(p_vals[idx]),4)
                diff = round(abs(real_flow - ref),4)
                check = "TRUE" if diff <= CONFIG["TOLERANCE"] else "FALSE"
                raw_data.append([file, I, part, ref, real_press, real_flow, diff, check])

df_raw = pd.DataFrame(raw_data, columns=[
    "数据源文件","电流值","数据分区","参考流量值","实际压差值","实际流量值","差值数据","校验结果"
])

# ---------------------- 步骤3：筛选全文件有效数据 ----------------------
print("\n✅ 【3/12】正在筛选所有文件均合格的有效检测点位...")
total_files = df_raw["数据源文件"].nunique()
all_valid_groups = df_raw.groupby(["电流值", "参考流量值", "数据分区"])["校验结果"].apply(
    lambda x: (x == "TRUE").all()
)
valid_groups = all_valid_groups[all_valid_groups].index.tolist()
df_valid = df_raw[df_raw.set_index(["电流值", "参考流量值", "数据分区"]).index.isin(valid_groups)].copy()
print("✅ 数据预处理全部完成")

# ---------------------- 步骤4：加权一致性分析 ----------------------
print("\n📊 【4/12】正在执行加权一致性影响分析...")
file_score, impact_rate, worst_file, scale_suggest = analyze_consistency_influence_final(df_valid)
print("✅ 一致性分析完成，已定位异常产品")

# ---------------------- 步骤5：提取关键点位数据 ----------------------
print("\n📌 【5/12】正在读取0mA 60L/min 基准压差值...")
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

# ---------------------- 步骤6：生成单文件性能图表 ----------------------
print("\n📈 【6/12】正在生成单产品性能曲线图...")
if CONFIG["RUN_PERFORMANCE_PLOTS"]:
    file_list = df_valid["数据源文件"].unique()
    perf_folder = "性能图表"
    perf_full = os.path.join(OUTPUT_FOLDER, perf_folder)
    if not os.path.exists(perf_full):
        os.makedirs(perf_full)

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
        save_safe_plot(fig, os.path.join(perf_folder, f"性能图表-{fig_idx}.png"))
    print("✅ 单产品性能图表生成完成")
else:
    print("ℹ️ 已跳过：单产品性能图表")

# ---------------------- 步骤7：生成PQ总对比散点图 ----------------------
print("\n📊 【7/12】正在生成所有产品PQ总对比散点图...")
if CONFIG["RUN_PQ_COMPARE_PLOT"] and pq_comparison_data:
    fig, ax = plt.subplots(figsize=(18, 10), dpi=300)
    fig.set_facecolor('white')
    unique_files = sorted(list(set([item["file"] for item in pq_comparison_data])))
    unique_currents = sorted(list(set([item["current"] for item in pq_comparison_data])))
    
    import matplotlib.cm as cm
    colors = cm.nipy_spectral(np.linspace(0, 1, len(unique_files)))
    file_color_map = {f: colors[i] for i, f in enumerate(unique_files)}
    
    line_dict = {}
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
                
                line = ax.plot(
                    cdat["flow"], cdat["press"],
                    marker='o', ms=3, linewidth=0.6,
                    color=file_color_map[f], label=label, alpha=0.8
                )[0]
                if show_label:
                    line_dict[line] = f

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

    ax.set_title("所有产品滤波后前50% PQ曲线对比", fontsize=16, weight='bold')
    ax.set_xlabel("流量 L/min", fontsize=14)
    ax.set_ylabel("压差 bar", fontsize=14)
    ax.grid(alpha=0.3)
    ax.set_xlim(0, max(CONFIG["REF_FLOW"]) + 10)
    ax.tick_params(axis='both', labelsize=12)
    if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

    ncol = min(3, max(1, len(unique_files) // 12))
    leg = ax.legend(loc="center left", bbox_to_anchor=(1.03, 0.5), ncol=ncol, fontsize=10)
    leg.handlelength = 1.2
    plt.tight_layout(rect=[0, 0, 0.87, 1])
    save_safe_plot(fig, "所有产品PQ曲线对比图.png", dpi=300)
    print("✅ PQ总对比散点图生成完成")
else:
    print("ℹ️ 已跳过：PQ总对比散点图")

# ---------------------- 步骤8：计算迟滞数据 ----------------------
print("\n🔁 【8/12】正在计算产品迟滞数据...")
hys_data = []
for (f,c,rf),g in df_valid.groupby(["数据源文件","电流值","参考流量值"]):
    b = g[g["数据分区"]=="前50%"]["实际压差值"]
    a = g[g["数据分区"]=="后50%"]["实际压差值"]
    if len(b)==1 and len(a)==1:
        hys_data.append([f,c,rf,round(abs(b.iloc[0]-a.iloc[0]),2)])
df_hys = pd.DataFrame(hys_data, columns=["数据源文件","电流值","参考流量值","迟滞"])
print("✅ 迟滞数据计算完成")

# ---------------------- 步骤9：新版一致性计算 ----------------------
print("\n📐 【9/12】正在计算产品一致性数据...")
df_con = calculate_consistency(df_valid)
print("✅ 一致性数据计算完成")

# ---------------------- 步骤10：生成Excel检测报告 ----------------------
print("\n📄 【10/12】正在生成Excel完整检测报告...")
if CONFIG["RUN_EXCEL_REPORT"]:
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    ws1 = wb.create_sheet("原始检测数据")
    ws1.append(["数据源文件","电流值","数据分区","参考流量值","实际压差值","实际流量值","差值数据","校验结果"])
    for _,r in df_raw.iterrows(): 
        ws1.append(list(r))

    ws2 = wb.create_sheet("产品迟滞")
    ws2.append(["数据源文件","电流值","参考流量值","迟滞"])
    for r in hys_data: 
        ws2.append(r)
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

    if CONFIG["RUN_EXCEL_CHARTS"]:
        chart_hys = ScatterChart()
        chart_hys.scatterStyle = "lineMarker"
        chart_hys.title = "P-Q滞环@20L/min"
        chart_hys.x_axis.title = "电流值"
        chart_hys.y_axis.title = "迟滞/bar"

        prod_list = []
        for row in range(2, plot_last_row+1):
            prod = ws2.cell(row,8).value
            if prod and prod not in prod_list:
                prod_list.append(prod)

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

    ws3 = wb.create_sheet("一致性对比")
    ws3.append([
        "电流值","参考流量值","数据分区",
        "最大值","最大值文件名","最小值","最小值文件名",
        "平均值","正偏差bar","正偏差%","负偏差bar","负偏差%"
    ])
    for _, r in df_con.iterrows():
        ws3.append([
            r["电流值"], r["参考流量值"], r["数据分区"],
            r["最大值"], r["最大值文件名"], r["最小值"], r["最小值文件名"],
            r["平均值"], r["正偏差bar"], r["正偏差%"], r["负偏差bar"], r["负偏差%"]
        ])

    def draw_con_chart(part, title, col, pos):
        data = df_con[df_con["数据分区"]==part].sort_values(["电流值","参考流量值"])
        currs = sorted(data["电流值"].unique())
        headers = ["电流值","参考流量值","平均值","正偏差bar","正偏差%","负偏差bar","负偏差%","标签"]
        for i,h in enumerate(headers): 
            ws3.cell(1,col+i,h)
        
        data_last_row = 1
        for i,(_,r) in enumerate(data.iterrows(),2):
            curr = r["电流值"]
            flow = r["参考流量值"]
            avg = r["平均值"]
            pos_bar = r["正偏差bar"]
            pos_pct = r["正偏差%"]
            neg_bar = r["负偏差bar"]
            neg_pct = r["负偏差%"]
            lab = f"+{pos_pct}% / -{neg_pct}%"
            ws3.cell(i,col+0,curr)
            ws3.cell(i,col+1,flow)
            ws3.cell(i,col+2,avg)
            ws3.cell(i,col+3,pos_bar)
            ws3.cell(i,col+4,pos_pct)
            ws3.cell(i,col+5,neg_bar)
            ws3.cell(i,col+6,neg_pct)
            ws3.cell(i,col+7,lab)
            data_last_row = i
        ws3.cell(data_last_row+1, col, "绘图数据结束")
        
        ch = ScatterChart()
        ch.scatterStyle = "lineMarker"
        ch.title = f"一致性 {title}"
        ch.x_axis.title = "流量/Lpm"
        ch.y_axis.title = "压差/bar"
        ch.x_axis.majorGridlines = None
        ch.y_axis.majorGridlines = None
        
        curr_list = []
        for row in range(2, data_last_row+1):
            curr_val = ws3.cell(row, col+0).value
            if curr_val and curr_val not in curr_list:
                curr_list.append(curr_val)
        curr_list = sorted(curr_list)
        
        for curr in curr_list:
            sr = er = None
            for row in range(2, data_last_row+1):
                if ws3.cell(row, col+0).value == curr:
                    sr = row if sr is None else sr
                    er = row
            if sr is None or er is None:
                continue
            
            x = Reference(ws3, col+1, sr, col+1, er)
            y = Reference(ws3, col+2, sr, col+2, er)
            label_ref = Reference(ws3, col+7, sr, col+7, er)
            
            ser = Series(y,x,title=f"{curr}mA")
            ser.marker = Marker(size=4)
            ser.graphicalProperties.line = LineProperties(solidFill="00B0F0", w=8000)
            ser.marker.graphicalProperties.solidFill="00B0F0"
            ser.dLbl = True
            ser.labelRef = label_ref
            ch.series.append(ser)
        ws3.add_chart(ch, pos)

    if CONFIG["RUN_EXCEL_CHARTS"]:
        draw_con_chart("前50%", "前50%", 15, "AQ2")
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
            rate = impact_rate.get(f,0)
            scale = scale_suggest.get(f,1.0)
            orig_p, new_p = scale_0mA60L.get(f, (0.0, 0.0))
            correct_msg = f"{orig_p} → {new_p}"
            note = f"全量程P值 × {scale}"
            ws4.append([f, round(score,2), rate, scale, correct_msg, note])

    excel_path = os.path.join(OUTPUT_FOLDER, CONFIG["OUTPUT_EXCEL"])
    base, ext = os.path.splitext(excel_path)
    cnt = 1
    while os.path.exists(excel_path):
        excel_path = f"{base}({cnt}){ext}"
        cnt +=1
    wb.save(excel_path)
    print("✅ Excel检测报告生成完成")
else:
    print("ℹ️ 已跳过：生成完整Excel检测报告")

# ---------------------- 步骤11：生成PNG分析图表 ----------------------
print("\n🎨 【11/12】正在生成迟滞、一致性、压差平均值图表...")
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
    unique_prods = plot_hys["数据源文件"].unique()
    ncol = min(3, len(unique_prods) // 15 + 1)
    ax.legend(loc="center left", bbox_to_anchor=(1.02, 0.5), ncol=ncol, fontsize=8)
    if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
    plt.tight_layout()
    save_safe_plot(fig, "迟滞图.png")
    print("✅ 迟滞图生成完成")
else:
    print("ℹ️ 已跳过：迟滞图")

# ==============================================================================
# 【✅ 已融合：一致性图 = 模式0用旧版标注，模式1/2用新版标注】
# ==============================================================================
if CONFIG["RUN_CONSISTENCY_PLOT"]:
    con_plot = df_con[df_con["数据分区"]=="前50%"].copy()
    currs = sorted(con_plot["电流值"].unique())
    fig, ax = plt.subplots(figsize=(15,6), dpi=100)

    for idx,curr in enumerate(currs):
        sub = con_plot[con_plot["电流值"]==curr].sort_values("参考流量值")
        ax.plot(sub["参考流量值"], sub["平均值"], marker='o', ms=4, color='#00B0F0')
        
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

        # ============== 核心融合逻辑 ==============
        if CONFIG["CONSISTENCY_MODE"] == 0:
            # 旧版标注：0mA/300mA + 5/10L → ±bar，其余→±%
            min_two_curr = currs[:2]
            for x, y, dif, pct in zip(sub["参考流量值"], sub["平均值"], sub["正偏差bar"], sub["正偏差%"]):
                if curr in min_two_curr and x in [5.0, 10.0]:
                    lab = f"±{dif:.2f}bar"
                else:
                    lab = f"±{pct:.2f}%"
                offset = CONFIG["LABEL_POS_0mA"] if curr == 0 else CONFIG["LABEL_POS_OTHER"]
                ax.annotate(lab, (x,y), xytext=(0, offset), textcoords='offset points', 
                            ha='center', fontsize=CONFIG["CON_LABEL_SIZE"])
        else:
            # 新版标注：正负百分比
            for x,y,pb,pp,nb,np in zip(sub["参考流量值"], sub["平均值"], sub["正偏差bar"], sub["正偏差%"], sub["负偏差bar"], sub["负偏差%"]):
                lab = f"+{pp}% / -{np}%"
                offset = CONFIG["LABEL_POS_0mA"] if curr == 0 else CONFIG["LABEL_POS_OTHER"]
                ax.annotate(lab, (x,y), xytext=(0, offset), textcoords='offset points', 
                            ha='center', fontsize=CONFIG["CON_LABEL_SIZE"])

    ax.set_title(CONFIG["CON_TITLE"], fontsize=CONFIG["CON_TITLE_SIZE"],
                 weight="bold" if CONFIG["CON_TITLE_BOLD"] else "normal",
                 x=CONFIG["CON_TITLE_OFFSET_X"], y=CONFIG["CON_TITLE_OFFSET_Y"])
    ax.set_xlabel("流量/Lpm", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"],
                  weight="bold" if CONFIG["PERF_AXIS_LABEL_BOLD"] else "normal")
    ax.set_ylabel("压差 bar", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"],
                  weight="bold" if CONFIG["PERF_AXIS_LABEL_BOLD"] else "normal")
    ax.tick_params(axis='both', labelsize=CONFIG["CON_TICK_LABEL_SIZE"], width=CONFIG["CON_TICK_WIDTH"])
    ax.set_xlim(right=ax.get_xlim()[1]+CONFIG["X_AXIS_EXTEND"])
    for line in ax.get_lines():
        line.set_linewidth(CONFIG["CON_LINE_WIDTH"])
    if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
    plt.tight_layout()
    save_safe_plot(fig, "一致性图.png")
    print("✅ 一致性图生成完成")
else:
    print("ℹ️ 已跳过：一致性图")

# ==============================================================================
# 【✅ 已修改：压差平均值图 模式0 全部统一显示 值±偏差bar】
# ==============================================================================
if CONFIG["RUN_AVG_PRESSURE_PLOT"]:
    con_plot_avg = df_con[df_con["数据分区"]=="前50%"].copy()
    currs_avg = sorted(con_plot_avg["电流值"].unique())

    fig_avg, ax_avg = plt.subplots(figsize=(15,6), dpi=100)
    for idx,curr in enumerate(currs_avg):
        sub_avg = con_plot_avg[con_plot_avg["电流值"]==curr].sort_values("参考流量值")
        ax_avg.plot(sub_avg["参考流量值"], sub_avg["平均值"], marker='o', ms=4, color='#00B0F0')
        
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

        # ============== 核心修改：模式0 全部统一 bar 标注 ==============
        if CONFIG["CONSISTENCY_MODE"] == 0:
            # 全部点位统一：平均值 ± 正偏差bar
            for x, y, avg_val, dif in zip(sub_avg["参考流量值"], sub_avg["平均值"], sub_avg["平均值"], sub_avg["正偏差bar"]):
                lab_avg = f"{avg_val:.2f}±{dif:.2f} bar"
                offset_avg = CONFIG["LABEL_POS_0mA"] if curr == 0 else CONFIG["LABEL_POS_OTHER"]
                ax_avg.annotate(lab_avg, (x,y), xytext=(0, offset_avg), textcoords='offset points', 
                            ha='center', fontsize=CONFIG["CON_LABEL_SIZE"])
        else:
            # 新版标注：平均值+正负百分比
            for x,y,avg_val,pb,pp,nb,np in zip(sub_avg["参考流量值"], sub_avg["平均值"], sub_avg["平均值"], sub_avg["正偏差bar"], sub_avg["正偏差%"], sub_avg["负偏差bar"], sub_avg["负偏差%"]):
                lab_avg = f"{avg_val:.1f} | +{pp}% / -{np}%"
                offset_avg = CONFIG["LABEL_POS_0mA"] if curr == 0 else CONFIG["LABEL_POS_OTHER"]
                ax_avg.annotate(lab_avg, (x,y), xytext=(0, offset_avg), textcoords='offset points', 
                            ha='center', fontsize=CONFIG["CON_LABEL_SIZE"])

    ax_avg.set_title("压差平均值", fontsize=CONFIG["CON_TITLE_SIZE"],
                     weight="bold" if CONFIG["CON_TITLE_BOLD"] else "normal",
                     x=CONFIG["CON_TITLE_OFFSET_X"], y=CONFIG["CON_TITLE_OFFSET_Y"])
    ax_avg.set_xlabel("流量/Lpm", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"],
                      weight="bold" if CONFIG["PERF_AXIS_LABEL_BOLD"] else "normal")
    ax_avg.set_ylabel("压差 bar", fontsize=CONFIG["PERF_AXIS_LABEL_SIZE"],
                      weight="bold" if CONFIG["PERF_AXIS_LABEL_BOLD"] else "normal")
    ax_avg.tick_params(axis='both', labelsize=CONFIG["CON_TICK_LABEL_SIZE"], width=CONFIG["CON_TICK_WIDTH"])
    ax_avg.set_xlim(right=ax_avg.get_xlim()[1]+CONFIG["X_AXIS_EXTEND"])
    for line in ax_avg.get_lines():
        line.set_linewidth(CONFIG["CON_LINE_WIDTH"])
    if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
        ax_avg.spines['top'].set_visible(False)
        ax_avg.spines['right'].set_visible(False)

    plt.tight_layout()
    save_safe_plot(fig_avg, "压差平均值图.png")
    print("✅ 压差平均值图生成完成")
else:
    print("ℹ️ 已跳过：压差平均值图")

# ---------------------- 步骤12：生成去程/回程/合并PQ曲线图 ----------------------
print("\n🔄 【12/12】正在生成去程、回程、合并PQ曲线图...")

all_go_data = []
all_back_data = []
all_both_data = []

if CONFIG["RUN_SINGLE_GOBACK_PQ"]:
    folder_go = "01_去程PQ"
    folder_back = "02_回程PQ"
    folder_both = "03_去程+回程PQ"
    
    os.makedirs(os.path.join(OUTPUT_FOLDER, folder_go), exist_ok=True)
    os.makedirs(os.path.join(OUTPUT_FOLDER, folder_back), exist_ok=True)
    os.makedirs(os.path.join(OUTPUT_FOLDER, folder_both), exist_ok=True)

    LINE_COLOR = "#00B0F0"

    for file_idx, file in enumerate(csv_files, 1):
        if file not in file_cleaned_data:
            continue

        df = file_cleaned_data[file].copy()
        file_short = os.path.splitext(file)[0]
        curr_list = sorted(df["电流值"].unique())
        file_go = []
        file_back = []

        # 去程
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
        save_safe_plot(fig1, os.path.join(folder_go, f"去程PQ_{file_idx}.png"))
        all_go_data.append(file_go)

        # 回程
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
        save_safe_plot(fig2, os.path.join(folder_back, f"回程PQ_{file_idx}.png"))
        all_back_data.append(file_back)

        # 合并
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
        save_safe_plot(fig3, os.path.join(folder_both, f"去程回程对比_{file_idx}.png"))
        all_both_data.append((file_go, file_back))

    print("✅ 单文件去程/回程/合并PQ图生成完成")
else:
    print("ℹ️ 已跳过：单文件去程/回程/合并PQ图")

# 全体去程回程总对比图
if CONFIG["RUN_ALL_GOBACK_SUMMARY"] and CONFIG["RUN_SINGLE_GOBACK_PQ"]:
    print("\n📊 正在生成所有文件PQ总对比图...")
    
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
    save_safe_plot(fig_go, os.path.join(folder_go, "所有文件_去程总对比.png"))

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
    save_safe_plot(fig_back, os.path.join(folder_back, "所有文件_回程总对比.png"))

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
    save_safe_plot(fig_both, os.path.join(folder_both, "所有文件_去程回程总对比.png"))
    print("✅ 全体去程回程总对比图生成完成！")
else:
    print("ℹ️ 已跳过：全体去程回程总对比图")

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

input("\n按 回车键 退出程序...")