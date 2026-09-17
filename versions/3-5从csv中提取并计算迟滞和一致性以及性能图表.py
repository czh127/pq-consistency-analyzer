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

# ======================================
# ====================== 自定义配置区 ======================
# 说明：这里所有参数都可以自己改，程序运行时自动生效
# ======================================
warnings.filterwarnings('ignore')  # 忽略运行警告，避免控制台干扰
CONFIG = {
    # ---------------------- 输出文件设置 ----------------------
    "OUTPUT_EXCEL": "产品检测结果总表.xlsx",  # 最终生成的Excel报告文件名
    "ENCODING": "gbk",                      # CSV文件编码格式（中文系统默认gbk）

    # ---------------------- 数据列索引 ----------------------
    "COL_CURRENT": 1,   # 电流值在CSV中的列索引（从0开始计数）
    "COL_FLOW": 4,      # 流量值在CSV中的列索引
    "COL_PRESS": 10,    # 压差值在CSV中的列索引

    # ---------------------- 标准流量点 ----------------------
    "REF_FLOW": [5.0, 10.0, 20.0, 40.0, 60.0],  # 程序自动匹配的标准流量点位(L/min)

    # ---------------------- 迟滞判定流量 ----------------------
    "PLOT_FLOW": 20.0,  # 迟滞图/一致性图固定使用的流量点

    # ---------------------- 迟滞标准阈值 ----------------------
    "STD_LIMIT": 10.0,          # 迟滞判断的标准上限(bar)
    "STD_LINE1": {"value": 10.0, "show": 1},  # 迟滞图标准线数值+是否显示
    "TOLERANCE": 1.00,          # 流量匹配允许误差范围(L/min)

    # ---------------------- 一致性判定 ----------------------
    "CONSISTENCY_LIMIT_PCT": 5.0,  # 一致性合格阈值(±5%)

    # ---------------------- 加权分析权重 ----------------------
    "KEY_FLOW": 20.0,        # 关键流量点
    "HIGH_FLOW": 40.0,       # 高流量点
    "KEY_CURRENTS": [0, 300, 600, 900],  # 重点关注电流值

    # ---------------------- 滤波参数 ----------------------
    "FILTER_WINDOW": 201,    # 滑动平均滤波窗口大小（奇数）
    "FILTER_ON": True,       # 是否开启数据滤波

    # ---------------------- 性能图表与一致性统一样式(电流标签位置和边框隐藏)----------------------
    "X_AXIS_EXTEND": 8,      # X轴右侧扩展长度
    "HIDE_TOP_RIGHT_BORDER": True,  # 是否隐藏图表上/右边框

    # ====================== 【性能图表自定义样式】 ======================
    "PERF_TITLE_SUFFIX": " 实测",          # 性能图标题后缀
    "PERF_TITLE_SIZE": 36,                # 标题字体大小
    "PERF_TITLE_BOLD": False,             # 标题是否加粗
    "PERF_TITLE_OFFSET_X": 0.5,    # 标题水平位置（0.5=居中，1=最右）
    "PERF_TITLE_OFFSET_Y": 1.1,    # 标题垂直位置（1.0=默认，越大越往上）
    "PERF_AXIS_LABEL_SIZE": 20,           # 坐标轴标签大小
    "PERF_AXIS_LABEL_BOLD": False,        # 坐标轴标签是否加粗
    "PERF_TICK_LABEL_SIZE": 20,           # 刻度字体大小
    "PERF_LINE_WIDTH": 2.0,               # 曲线宽度
    "PERF_TICK_WIDTH": 1.2,               # 刻度线宽度

    # ====================== 性能图表电流标签样式 ======================
    "PERF_TEXT_OFFSET_X": 1.5,            # 电流标注X偏移
    "PERF_TEXT_OFFSET_Y": 0.0,            # 电流标注Y偏移
    "PERF_CURR_TEXT_SIZE": 20,            # 电流标注字体大小
    "PERF_CURR_TEXT_BOLD": False,         # 电流标注是否加粗

    # ====================== 【一致性图表自定义样式】 ======================
    "CON_TITLE": "一致性",                      # 一致性图表主标题文字
    "CON_TITLE_SIZE": 24,                       # 主标题字体大小
    "CON_TITLE_BOLD": False,                    # 主标题是否加粗
    "CON_TITLE_OFFSET_X": 0.5,                  # 主标题水平偏移（0.5=居中，1=最右）
    "CON_TITLE_OFFSET_Y": 1.1,                  # 主标题垂直偏移（1.0=默认，越大越往上）
    "CON_AXIS_LABEL_SIZE": 20,                  # 坐标轴标签（流量/压差）字体大小
    "CON_AXIS_LABEL_BOLD": False,               # 坐标轴标签是否加粗
    "CON_TICK_LABEL_SIZE": 16,                  # 坐标轴刻度数字字体大小
    "CON_LINE_WIDTH": 2.0,                      # 一致性曲线线条宽度
    "CON_TICK_WIDTH": 1.2,                      # 坐标轴刻度线粗细

    # ====================== 一致性标签样式 ======================
    "LABEL_POS_0mA": -14,                       # 0mA 电流的数据标签垂直偏移（负数向下）
    "LABEL_POS_OTHER": 8,                       # 其他电流的数据标签垂直偏移（正数向上）
    "CON_LABEL_SIZE": 10,                       # 一致性误差标签字体大小（±xxbar / ±xx%）
    "CON_CURR_TEXT_SIZE": 20,                   # 一致性图表电流值字体大小（0mA / 600mA 等标注）
}

# ---------------------- 颜色配置（不用改） ----------------------
# 曲线颜色映射表：0=绿色,1=蓝色,2=黄色,3=红色,4=深红色
COLOR_MAP = {0: '#70AD47', 1: '#00B0F0', 2: '#FFC000', 3: '#FF0000', 4: '#C00000'}
COLOR_EXCEL = ["70AD47", "00B0F0", "FFC000", "FF0000", "C00000"]  # Excel图表颜色
FILE_COLORS = {  # 多文件对比时自动分配的颜色
    0: '#00B0F0', 1: '#FFC000', 2: '#70AD47', 3: '#FF0000', 4: '#C00000',
    5: '#9933FF', 6: '#00FF00', 7: '#00FFFF', 8: '#FF9900', 9: '#999999'
}

# ======================================
# ====================== 以下为程序核心代码 ======================
# ======================================

# ====================== 输出文件夹（自动创建，防重名） ======================
# 功能：自动创建结果输出文件夹，若已存在则自动重命名（检测结果输出_1/2/3...）
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

# ====================== 图片安全保存 ======================
# 功能：防止图片重名覆盖，自动添加(1)(2)后缀
# 输入：fig=图表对象, filename=保存文件名, dpi=清晰度, bbox_inches=紧凑布局
def save_safe_plot(fig, filename, dpi=150, bbox_inches='tight'):
    path = os.path.join(OUTPUT_FOLDER, filename)
    base, ext = os.path.splitext(path)
    counter = 1
    while os.path.exists(path):
        path = f"{base}({counter}){ext}"
        counter += 1
    fig.savefig(path, dpi=dpi, bbox_inches=bbox_inches)
    return path

# ====================== 工具函数 ======================
# 功能：在数组中找到最接近目标值的索引
# 示例：arr=[5,10,20], target=18 → 返回索引2（值20）
def get_closest_index(arr, target):
    return np.argmin(np.abs(arr - target))

# 功能：安全计算百分比，避免除0错误
# 输入：分子,分母,默认值
def safe_percent(numerator, denominator, default=0):
    try:
        return round((numerator / denominator) * 100, 1)
    except ZeroDivisionError:
        return default

# 功能：滑动平均滤波，平滑数据曲线
# 输入：原始数据序列, 窗口大小
def moving_average_filter(series, window=201):
    if not CONFIG["FILTER_ON"] or window <= 1:
        return series.copy()
    return series.rolling(window=window, center=True, min_periods=1).mean()

# 功能：清洗无效数据（空值、无穷值、非数值）
def clean_invalid_data(df):
    df = df.replace([False, None, np.inf, -np.inf], np.nan)
    df = df.dropna(subset=["流量值", "压差值"])
    df["流量值"] = pd.to_numeric(df["流量值"], errors="coerce")
    df["压差值"] = pd.to_numeric(df["压差值"], errors="coerce")
    return df

# ====================== 加权一致性分析 ======================
# 功能：核心算法 → 计算每个文件对整体一致性的影响权重，找出异常文件
# 输出：文件得分、影响占比、最差文件、修正系数
def analyze_consistency_influence_final(df_valid):
    influence = []
    KEY_CURRENTS = CONFIG["KEY_CURRENTS"]
    KEY_FLOW = CONFIG["KEY_FLOW"]
    HIGH_FLOW = CONFIG["HIGH_FLOW"]

    # 按【电流+参考流量+数据分区】分组计算一致性
    groups = df_valid.groupby(["电流值", "参考流量值", "数据分区"])
    for (I, rf, part), g in groups:
        vals = g["实际压差值"].dropna()
        if len(vals) < 2:
            continue

        mean_all = vals.mean()       # 组内平均值
        maxv = vals.max()            # 组内最大值
        minv = vals.min()            # 组内最小值
        half_diff = (maxv - minv) / 2  # 半极差
        consistency_pct = safe_percent(half_diff, mean_all)  # 一致性百分比

        # 合格点位跳过，只分析不合格
        if abs(consistency_pct) <= CONFIG["CONSISTENCY_LIMIT_PCT"]:
            continue

        # ====================== 加权规则 ======================
        # 0mA权重最高(10) → 300mA(8) → 600mA(6) → 900mA(4)
        # 关键流量20L权重翻倍
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

            if rf == KEY_FLOW:
                weight = base_w * 2.0
            else:
                weight = base_w
        elif I > 900:
            weight = 1.0
        else:
            weight = 2.0 if rf >= HIGH_FLOW else 1.0

        # 统计每个文件的偏差
        for f in g["数据源文件"].unique():
            g_f = g[g["数据源文件"] == f]
            if g_f.empty:
                continue
            p = g_f["实际压差值"].iloc[0]
            dev = abs(p - mean_all)               # 绝对偏差
            weighted_dev = dev * weight           # 加权偏差

            influence.append({
                "file": f, "I": I, "rf": rf,
                "P": p, "mean": mean_all,
                "weight": weight,
                "weighted_dev": weighted_dev,
                "consistency_pct": consistency_pct
            })

    if not influence:
        return None, None, None, None

    # 按文件汇总加权得分
    df_inf = pd.DataFrame(influence)
    file_score = df_inf.groupby("file")["weighted_dev"].sum().sort_values(ascending=False)
    total = file_score.sum()
    impact_rate = {f: round(s / total * 100, 1) for f, s in file_score.items()}  # 影响占比
    worst_file = file_score.index[0]  # 影响最大的异常文件

    # 计算修正缩放系数（全局均值/文件均值）
    suggest = {}
    for f in df_inf["file"]:
        sub = df_inf[df_inf["file"] == f]
        f_mean = sub["P"].mean()
        global_mean = df_inf["mean"].mean()
        scale = round(global_mean / f_mean, 3) if f_mean != 0 else 1.0
        suggest[f] = scale

    return file_score, impact_rate, worst_file, suggest

# ====================== 读取真实 0mA 60L/min 压差值 ======================
# 功能：提取0mA、60L/min、前半段数据的压差，用于修正报告
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

# ====================== 主程序 ======================
print("="*60)
print("🚀 开始运行 PQ 一致性检测分析程序")
print("="*60)

# 切换到脚本所在目录，确保读取同文件夹下的CSV
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

print("\n📂 步骤1：扫描当前目录 CSV 文件...")
csv_files = glob.glob("*.csv")  # 匹配所有.csv文件
print(f"✅ 找到 {len(csv_files)} 个数据文件")

raw_data = []            # 原始检测数据存储列表
pq_comparison_data = []  # PQ曲线对比数据存储列表

print("\n📊 步骤2：读取并滤波处理数据...")
for file in csv_files:
    try:
        # 按配置列索引读取电流、流量、压差三列
        df = pd.read_csv(file, usecols=[1,4,10], encoding="gbk")
        df.columns = ["电流值","流量值","压差值"]
    except:
        continue  # 读取失败跳过

    df = clean_invalid_data(df)
    if df.empty:
        continue

    # 保存原始值+生成滤波后值
    df["流量值原始"] = df["流量值"].copy()
    df["压差值原始"] = df["压差值"].copy()
    df["流量值滤波"] = moving_average_filter(df["流量值"])
    df["压差值滤波"] = moving_average_filter(df["压差值"])
    df = clean_invalid_data(df)
    if df.empty:
        continue

    # 使用滤波后数据参与计算
    df["流量值"] = df["流量值滤波"]
    df["压差值"] = df["压差值滤波"]

    file_short = os.path.splitext(file)[0]  # 去掉.csv后缀
    # 按电流分组，提取标准流量点数据（用于PQ对比图）
    for I, g in df.groupby("电流值"):
        I = round(I, 2)
        mid = len(g) // 2
        first_half = g.iloc[:mid]  # 只使用前50%数据
        first_half = clean_invalid_data(first_half)
        if first_half.empty:
            continue

        f_vals = first_half["流量值"].abs().values
        p_vals = first_half["压差值"].abs().values
        flow_points = []
        press_points = []
        # 匹配5/10/20/40/60L标准流量
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

    # 生成完整原始检测表（前50%+后50%）
    for I, g in df.groupby("电流值"):
        I = round(I, 2)
        mid = len(g) // 2
        # 拆分前后半段数据
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

# 构建原始数据DataFrame
df_raw = pd.DataFrame(raw_data, columns=[
    "数据源文件","电流值","数据分区","参考流量值","实际压差值","实际流量值","差值数据","校验结果"
])
df_valid = df_raw[df_raw["校验结果"] == "TRUE"].copy()  # 只保留流量匹配合格数据

print("✅ 数据读取与滤波完成")

print("\n⚖️ 步骤3：执行加权一致性影响分析...")
file_score, impact_rate, worst_file, scale_suggest = analyze_consistency_influence_final(df_valid)
print("✅ 加权分析完成")

print("\n📌 步骤4：读取 0mA 60L/min 真实压差值...")
all_files = df_valid["数据源文件"].unique()
orig_0mA60L = get_real_0mA60L_press(df_valid, all_files)
scale_0mA60L = {}
for f in all_files:
    orig = orig_0mA60L.get(f, 0.0)
    scale = scale_suggest.get(f, 1.0)
    new_p = round(orig * scale, 2)  # 计算修正后压差
    scale_0mA60L[f] = (orig, new_p)
print("✅ 0mA 60L/min 数值读取完成")

# Matplotlib中文显示设置
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

# ====================== 【性能图：自定义大字体 + 标签靠右】 ======================
# 功能：为每个CSV文件生成独立的性能曲线图
print("\n📈 步骤5：生成单文件性能图表...")
file_list = df_valid["数据源文件"].unique()
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
        # 在曲线末端标注电流值
        ax.text(
            x_last + CONFIG["PERF_TEXT_OFFSET_X"],
            y_last + CONFIG["PERF_TEXT_OFFSET_Y"],
            f"{curr}mA",
            fontsize=CONFIG["PERF_CURR_TEXT_SIZE"],
            weight="bold" if CONFIG["PERF_CURR_TEXT_BOLD"] else "normal",
            va='center'
        )
    title_text = f"{os.path.splitext(file)[0]}{CONFIG['PERF_TITLE_SUFFIX']}"
    ax.set_title(
        title_text,
        fontsize=CONFIG["PERF_TITLE_SIZE"],
        weight="bold" if CONFIG["PERF_TITLE_BOLD"] else "normal",
        x=CONFIG["PERF_TITLE_OFFSET_X"],  # 新增：水平偏移
        y=CONFIG["PERF_TITLE_OFFSET_Y"]   # 新增：垂直偏移
    )  
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
    save_safe_plot(fig, f"性能图表-{fig_idx}.png")
    plt.close()

print("✅ 单文件性能图表生成完成")

# ====================== PQ 总对比图 ======================
# 功能：所有文件PQ曲线叠加对比图
print("\n📊 步骤6：生成所有产品 PQ 对比图...")
if pq_comparison_data:
    fig, ax = plt.subplots(figsize=(15,8), dpi=100)
    unique_files = list(set([item["file"] for item in pq_comparison_data]))
    unique_currents = sorted(list(set([item["current"] for item in pq_comparison_data])))
    file_color_map = {f:FILE_COLORS[i%len(FILE_COLORS)] for i,f in enumerate(unique_files)}
    
    plotted_files = set()
    for f in unique_files:
        dat = [d for d in pq_comparison_data if d["file"]==f]
        for curr in unique_currents:
            cdat = next((d for d in dat if d["current"]==curr), None)
            if cdat:
                show_label = f not in plotted_files
                label = f if show_label else ""
                if show_label:
                    plotted_files.add(f)
                
                ax.plot(cdat["flow"], cdat["press"], marker='o', ms=3, linewidth=0.6,
                        color=file_color_map[f], label=label, alpha=0.8)
    
    ax.set_title("所有产品滤波后前50% PQ曲线对比", fontsize=16, weight='bold')
    ax.set_xlabel("流量 L/min")
    ax.set_ylabel("压差 bar")
    ax.grid(alpha=0.3)
    ax.set_xlim(0, max(CONFIG["REF_FLOW"])+10)
    if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
    
    ax.legend(bbox_to_anchor=(1.05,1), loc='upper left', fontsize=10)
    plt.tight_layout()
    save_safe_plot(fig, "所有产品PQ曲线对比图.png")
    plt.close()

print("✅ PQ 总对比图生成完成")

# ====================== 迟滞数据 ======================
# 功能：计算前后半段压差差值 → 迟滞值
print("\n🔄 步骤7：计算迟滞数据...")
hys_data = []
for (f,c,rf),g in df_valid.groupby(["数据源文件","电流值","参考流量值"]):
    b = g[g["数据分区"]=="前50%"]["实际压差值"]
    a = g[g["数据分区"]=="后50%"]["实际压差值"]
    if len(b)==1 and len(a)==1:
        hys_data.append([f,c,rf,round(abs(b.iloc[0]-a.iloc[0]),2)])
df_hys = pd.DataFrame(hys_data, columns=["数据源文件","电流值","参考流量值","迟滞"])
print("✅ 迟滞计算完成")

# ====================== 一致性数据 ======================
# 功能：计算多文件同一工况下的最大/最小/平均/一致性百分比
print("\n📏 步骤8：计算一致性数据...")
con_data = []
for (c,rf,p),g in df_valid.groupby(["电流值","参考流量值","数据分区"]):
    vals = g["实际压差值"].dropna()
    if len(vals)<1: continue
    maxv = vals.max()
    minv = vals.min()
    avg = (maxv+minv)/2
    dif = (maxv-minv)/2  # 半差值
    con_data.append([round(c,2), round(rf,1), p,
                     round(maxv,2), g.loc[vals.idxmax(),"数据源文件"],
                     round(minv,2), g.loc[vals.idxmin(),"数据源文件"],
                     round(avg,2), round(dif,2), round(safe_percent(dif,avg), 2)])
df_con = pd.DataFrame(con_data, columns=[
    "电流值","参考流量值","数据分区","最大值","最大值文件名","最小值","最小值文件名","平均值","diff(半差值)","一致性%"
])
print("✅ 一致性计算完成")

# ====================== Excel 报告 ======================
# 功能：生成包含4个工作表的完整检测报告
print("\n📝 步骤9：生成 Excel 报告...")
wb = Workbook()
if "Sheet" in wb.sheetnames:
    del wb["Sheet"]

# 工作表1：原始检测数据
ws1 = wb.create_sheet("原始检测数据")
ws1.append(["数据源文件","电流值","数据分区","参考流量值","实际压差值","实际流量值","差值数据","校验结果"])
for _,r in df_raw.iterrows(): ws1.append(list(r))

# 工作表2：产品迟滞 + 迟滞折线图
ws2 = wb.create_sheet("产品迟滞")
ws2.append(["数据源文件","电流值","参考流量值","迟滞"])
for r in hys_data: ws2.append(r)
plot_hys = df_hys[df_hys["参考流量值"]==CONFIG["PLOT_FLOW"]].copy().sort_values(["数据源文件","电流值"])
ws2.cell(1,8,"绘图文件名")
ws2.cell(1,9,"绘图电流")
ws2.cell(1,10,"绘图迟滞")
for i,(_,r) in enumerate(plot_hys.iterrows(),2):
    ws2.cell(i,8,r["数据源文件"])
    ws2.cell(i,9,r["电流值"])
    ws2.cell(i,10,r["迟滞"])
# 新增：标记绘图数据的最后一行（便于绘图时直接读取）
plot_last_row = len(plot_hys) + 1  # 表头行=1，数据从2开始
ws2.cell(plot_last_row+1,8,"绘图数据结束")  # 边界标记，可选
# 插入Excel散点图
chart_hys = ScatterChart()
chart_hys.scatterStyle = "lineMarker"
chart_hys.title = "P-Q滞环@20L/min"
chart_hys.x_axis.title = "电流值"
chart_hys.y_axis.title = "迟滞/bar"

chart_hys = ScatterChart()
chart_hys.scatterStyle = "lineMarker"
chart_hys.title = "P-Q滞环@20L/min"
chart_hys.x_axis.title = "电流值"
chart_hys.y_axis.title = "迟滞/bar"

# 步骤1：读取Excel中所有绘图文件名（第8列），去重得到产品列表
prod_list = []
for row in range(2, plot_last_row+1):
    prod = ws2.cell(row,8).value
    if prod and prod not in prod_list:
        prod_list.append(prod)

# 步骤2：逐产品读取对应的电流（9列）、迟滞（10列）单元格范围
for prod in prod_list:
    # 找到该产品在Excel中的行范围
    sr = None
    er = None
    for row in range(2, plot_last_row+1):
        if ws2.cell(row,8).value == prod:
            if sr is None:
                sr = row
            er = row
    if sr is None or er is None:
        continue
    
    # 仅读取Excel单元格数据（不再依赖df_hys）
    x = Reference(ws2, 9, sr, 9, er)  # 绘图电流列
    y = Reference(ws2, 10, sr, 10, er)  # 绘图迟滞列
    
    # 计算超标数（从Excel迟滞列读取值）
    exceed_num = 0
    for row in range(sr, er+1):
        hys_val = ws2.cell(row,10).value
        if hys_val and float(hys_val) > CONFIG["STD_LIMIT"]:
            exceed_num += 1
    color = COLOR_EXCEL[min(exceed_num,4)]
    
    # 绘制折线
    ser = Series(y,x,title=prod)
    ser.marker = Marker(size=3)
    ser.graphicalProperties.line = LineProperties(solidFill=color, w=10000)
    chart_hys.series.append(ser)

ws2.add_chart(chart_hys, "L2")


# 工作表3：一致性对比 + 一致性曲线图
ws3 = wb.create_sheet("一致性对比")
ws3.append([
    "电流值","参考流量值","数据分区",
    "最大值","最大值文件名","最小值","最小值文件名",
    "平均值","diff(半差值)","一致性%"
])
for r in con_data: ws3.append(r)

# 绘制一致性图表函数（前50%/后50%）
def draw_con_chart(part, title, col, pos):
    # 步骤1：写入一致性绘图数据到Excel（原逻辑保留，确保数据完整）
    data = df_con[df_con["数据分区"]==part].sort_values(["电流值","参考流量值"])
    currs = sorted(data["电流值"].unique())
    headers = ["电流值","参考流量值","最大值","最大值文件名","最小值","最小值文件名","平均值","diff(半差值)","一致性%","数据标签"]
    for i,h in enumerate(headers): ws3.cell(1,col+i,h)
    
    # 写入数据并记录最后一行
    data_last_row = 1  # 表头行
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
        data_last_row = i  # 记录数据最后一行
    ws3.cell(data_last_row+1, col, "绘图数据结束")  # 边界标记
    
    # 步骤2：修改绘图逻辑——仅读取Excel单元格数据
    ch = ScatterChart()
    ch.scatterStyle = "lineMarker"
    ch.title = f"一致性 {title}"
    ch.x_axis.title = "流量/Lpm"
    ch.y_axis.title = "压差/bar"
    ch.x_axis.majorGridlines = None
    ch.y_axis.majorGridlines = None
    
    # 从Excel读取所有电流值（去重）
    curr_list = []
    for row in range(2, data_last_row+1):
        curr_val = ws3.cell(row, col+0).value
        if curr_val and curr_val not in curr_list:
            curr_list.append(curr_val)
    curr_list = sorted(curr_list)
    
    # 逐电流读取Excel数据绘图
    for curr in curr_list:
        # 找到该电流对应的行范围
        sr = None
        er = None
        for row in range(2, data_last_row+1):
            if ws3.cell(row, col+0).value == curr:
                if sr is None:
                    sr = row
                er = row
        if sr is None or er is None:
            continue
        
        # 仅读取Excel单元格（流量列、平均值列、标签列）
        x = Reference(ws3, col+1, sr, col+1, er)  # 参考流量值
        y = Reference(ws3, col+6, sr, col+6, er)  # 平均值
        label_ref = Reference(ws3, col+9, sr, col+9, er)  # 数据标签
        
        # 绘制折线
        ser = Series(y,x,title=f"{curr}mA")
        ser.marker = Marker(size=4)
        ser.graphicalProperties.line = LineProperties(solidFill="00B0F0", w=8000)
        ser.marker.graphicalProperties.solidFill = "00B0F0"
        ser.dLbl = True
        ser.labelRef = label_ref
        ch.series.append(ser)
    
    ws3.add_chart(ch, pos)

draw_con_chart("前50%", "前50%", 15, "AQ2")
draw_con_chart("后50%", "后50%", 28, "AQ20")

# 工作表4：加权一致性影响分析 + 修正建议
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

# 保存Excel（防重名）
excel_path = os.path.join(OUTPUT_FOLDER, CONFIG["OUTPUT_EXCEL"])
base, ext = os.path.splitext(excel_path)
cnt = 1
while os.path.exists(excel_path):
    excel_path = f"{base}({cnt}){ext}"
    cnt +=1
wb.save(excel_path)
print("✅ Excel 报告生成完成")

# ====================== 【迟滞图：原版颜色判定逻辑】 ======================
# 功能：生成PNG迟滞对比图，超标点数越多颜色越红
print("\n🎨 步骤10：生成迟滞与一致性图表...")
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
ax.legend(bbox_to_anchor=(1,1))
if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
plt.tight_layout()
save_safe_plot(fig, "迟滞图.png")
plt.close()

# ====================== 一致性图表 ======================
con_plot = df_con[df_con["数据分区"]=="前50%"].copy()
currs = sorted(con_plot["电流值"].unique())
# 自动获取最小的 2 个电流（动态逻辑，不写死固定值）
min_two_currents = sorted(con_plot["电流值"].unique())[:2]
fig, ax = plt.subplots(figsize=(15,6), dpi=100)

for idx,curr in enumerate(currs):
    sub = con_plot[con_plot["电流值"]==curr].sort_values("参考流量值")
    ax.plot(sub["参考流量值"], sub["平均值"], marker='o', ms=4, color='#00B0F0')
    
    # 电流值标注（独立字体配置）
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

    for x,y,dif,pct in zip(sub["参考流量值"], sub["平均值"], sub["diff(半差值)"], sub["一致性%"]):
        # ✅ 动态逻辑：最小2个电流 + 最小2个流量 → bar
        lab = f"±{dif:.2f}bar" if (curr in min_two_currents and x in [5.0, 10.0]) else f"±{pct:.2f}%"
        offset = CONFIG["LABEL_POS_0mA"] if curr == 0 else CONFIG["LABEL_POS_OTHER"]
        ax.annotate(lab, (x,y), xytext=(0, offset), textcoords='offset points', 
                    ha='center', fontsize=CONFIG["CON_LABEL_SIZE"])

# 一致性图表样式（完全对标性能图）
ax.set_title(
    CONFIG["CON_TITLE"],
    fontsize=CONFIG["CON_TITLE_SIZE"],
    weight="bold" if CONFIG["CON_TITLE_BOLD"] else "normal",
    x=CONFIG["CON_TITLE_OFFSET_X"],
    y=CONFIG["CON_TITLE_OFFSET_Y"]
)
ax.set_xlabel(
    "流量/Lpm",
    fontsize=CONFIG["CON_AXIS_LABEL_SIZE"],
    weight="bold" if CONFIG["CON_AXIS_LABEL_BOLD"] else "normal"
)
ax.set_ylabel(
    "压差 bar",
    fontsize=CONFIG["CON_AXIS_LABEL_SIZE"],
    weight="bold" if CONFIG["CON_AXIS_LABEL_BOLD"] else "normal"
)
ax.tick_params(
    axis='both',
    labelsize=CONFIG["CON_TICK_LABEL_SIZE"],
    width=CONFIG["CON_TICK_WIDTH"]
)
ax.set_xlim(right=ax.get_xlim()[1]+CONFIG["X_AXIS_EXTEND"])
# 线条宽度
for line in ax.get_lines():
    line.set_linewidth(CONFIG["CON_LINE_WIDTH"])
# 边框
if CONFIG["HIDE_TOP_RIGHT_BORDER"]:
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
plt.tight_layout()
save_safe_plot(fig, "一致性图.png")
plt.close()
print("✅ 迟滞 & 一致性图表生成完成")

# ====================== 最终输出 ======================
# 控制台打印最终结论：影响占比+修正建议
print("\n" + "="*85)
print("📊 全部文件一致性影响占比 & 修正建议（按影响从大到小）")
print("="*85)

if file_score is not None and len(file_score) > 0:
    print(f"{'文件名':<25} {'影响占比':<10} {'缩放系数':<10} {'0mA 60L/min 修正（原始→新）'}")
    print("-"*85)
    for f, score in file_score.items():
        rate = f"{impact_rate.get(f, 0):.1f}%"
        scale = f"×{scale_suggest.get(f, 1.0):.2f}"
        orig, new = scale_0mA60L.get(f, (0.0, 0.0))
        corr = f"{orig:.2f} → {new:.2f}"
        print(f"{f:<25} {rate:<10} {scale:<10} {corr}")
else:
    print("✅ 所有点位一致性均在 ±5% 以内，无需调节")

print(f"\n🎉 程序全部运行完成！")
print(f"📂 所有结果保存在文件夹：{OUTPUT_FOLDER}")