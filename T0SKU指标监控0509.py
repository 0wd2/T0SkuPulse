import streamlit as st
import plotly.graph_objects as go
import polars as pl
import pandas as pd
from io import BytesIO
from datetime import date, timedelta
from openpyxl import load_workbook
from plotly.subplots import make_subplots
import plotly.express as px
import numpy as np

# 1、设置页面标题
st.set_page_config(page_title="T0SKU指标监控", layout="wide", page_icon="📊")

# --- CSS 注入：实现表头与筛选面板固定 ---
# st.markdown("""
#     <style>
#     /* 强制固定顶部容器 */
#     [data-testid="stVerticalBlock"] > div:has(#fixed-header-anchor) {
#         position: sticky;
#         top: 2.8rem;
#         background-color: white;
#         z-index: 1000;
#         border-bottom: 2px solid #f0f2f6;
#         padding-bottom: 10px;
#     }
#     #fixed-header-anchor { display: none; }

#     /* 指标卡美化 */
#     div[data-testid="metric-container"] {
#         background-color: #f8f9fa;
#         border: 1px solid #eef0f2;
#         padding: 10px;
#         border-radius: 8px;
#     }
#     </style>
#     """, unsafe_allow_html=True)

# 2、初始化 Session State
if "df_fahuo" not in st.session_state: st.session_state.df_fahuo = None
if "df_yuce" not in st.session_state: st.session_state.df_yuce = None
if "df_ganyu" not in st.session_state: st.session_state.df_ganyu = None
if "df_ganyu_bi" not in st.session_state: st.session_state.df_ganyu_bi = None
if "filter_ver" not in st.session_state: st.session_state.filter_ver = 0
if "t0_date" not in st.session_state: st.session_state.t0_date = None
if "fahuo_frequency" not in st.session_state: st.session_state.fahuo_frequency = None
if "target_week" not in st.session_state: st.session_state.target_week = None
if "committed_filters" not in st.session_state:
    st.session_state.committed_filters = {"sub_market": [], "category": [], "mrpsku": []}
if "stock_turnover" not in st.session_state: st.session_state.stock_turnover = None
if "df_country_stock" not in st.session_state: st.session_state.df_country_stock = None
if "filter_market" not in st.session_state: st.session_state.filter_market = None
if "df_country_turnover" not in st.session_state: st.session_state.df_country_turnover = None
if "df_stock_turnover" not in st.session_state: st.session_state.df_stock_turnover = None
st.markdown("""
        <style>
        /* 1. 强制放大 Label (指标名称) */
        [data-testid="stMetricLabel"] div {
            font-size: 24px !important;
            font-weight: bold !important;
            font-family: "Microsoft YaHei", sans-serif;
            color: #31333F !important; /* 确保颜色清晰 */
        }

        /* 2. 强制放大 Delta (下方目标/变动文字) */
        [data-testid="stMetricDelta"] div {
            font-size: 19px !important;
        }
        </style>
        """, unsafe_allow_html=True)

# 3、侧边栏配置
with st.sidebar:
    st.header("⚙️ 配置参数")
    today = date.today()
    default_monday = today - timedelta(days=today.weekday()) - timedelta(days=7)
    t0_date_val = st.date_input("🗓️ 当前周周一", value=default_monday)
    st.session_state.t0_date = t0_date_val

    up_folder_btn = st.file_uploader("请选择上传文件：", type=['xlsx','parquet'],accept_multiple_files=True)
    if up_folder_btn is not None:
        for uploaded_file in up_folder_btn:
            try:
                file_name = uploaded_file.name
                file_bytes = uploaded_file.getvalue()
                file_type = file_name.split('.')[-1].lower()
                df_mapping = {} 
                if file_type == 'xlsx':
                    workbook = load_workbook(BytesIO(file_bytes), read_only=True)
                    sn = workbook.sheetnames
                    def read_and_format(sheet):
                        df = pl.read_excel(file_bytes, sheet_name=sheet, raise_if_empty=False).to_pandas()
                        if '日期' in df.columns: df['日期'] = pd.to_datetime(df['日期'])
                        return df
                    if '发货指标' in sn: st.session_state.df_fahuo = read_and_format('发货指标')
                    if '预测偏差' in sn: st.session_state.df_yuce = read_and_format('预测偏差')
                    if '干预偏差' in sn: st.session_state.df_ganyu = read_and_format('干预偏差')
                    if '干预比例' in sn: st.session_state.df_ganyu_bi = read_and_format('干预比例')
                    if '库存与周转' in sn: st.session_state.df_stock_turnover = read_and_format('库存与周转')
                    if '国内库存数据' in sn: st.session_state.df_country_stock = read_and_format('国内库存数据')
                    if '国内库存周转' in sn: st.session_state.df_country_turnover = read_and_format('国内库存周转')
                elif file_type == 'parquet':
                    st.session_state.stock_turnover = pd.read_parquet(file_bytes)

                st.success("✅ 数据加载成功")
            except Exception as e:
                st.error(f"❌ 读取失败: {e}")
    
    st.divider() # --- 分割线 ---

    # --- 第三部分：目录大纲 ---
    st.sidebar.subheader("📌 快速导航")

    # 使用 HTML 和 CSS 打造更具设计感的菜单
    toc_html = """
    <style>
        .toc-container {
            padding: 10px 0;
        }
        .toc-item {
            display: flex;
            align-items: center;
            padding: 12px 15px;
            margin: 8px 0;
            background-color: #f8f9fa;
            border-radius: 10px;
            text-decoration: none;
            color: #31333F !important;
            font-weight: 500;
            transition: all 0.3s ease;
            border: 1px solid #e0e0e0;
        }
        .toc-item:hover {
            background-color: #ffffff;
            border-color: #ff4b4b;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
            transform: translateX(5px);
            color: #ff4b4b !important;
        }
        .toc-icon {
            margin-right: 12px;
            font-size: 1.2rem;
        }
    </style>
    <div class="toc-container">
        <a href="#1" class="toc-item">
            <span class="toc-icon">📉</span> 海外库存周转
        </a>
        <a href="#2" class="toc-item">
            <span class="toc-icon">🚚</span> 发货指标监控
        </a>
        <a href="#3" class="toc-item">
            <span class="toc-icon">📈</span> 预测偏差
        </a>
        <a href="#4" class="toc-item">
            <span class="toc-icon">🛠️</span> 干预偏差
        </a>
    </div>
    """
    st.sidebar.markdown(toc_html, unsafe_allow_html=True)

# =========================================================
# 4、固定区域：Head、KPI卡片、筛选面板
# =========================================================
fixed_container = st.container()

with fixed_container:
    # 锚点
    # st.markdown('<div id="fixed-header-anchor"></div>', unsafe_allow_html=True)
    st.markdown("<h2 style='text-align: center; color: #1E3A8A; margin:0;'>📊 T0SKU 全链路指标监控</h2>",
                unsafe_allow_html=True)
    # 显示当前周周一,把日期变为年周
    current_week = st.session_state.t0_date.strftime("%Yw%V")
    # 年份只显示后两位，yyWww
    current_week = current_week[2:]
    head_1,head_2 = st.columns([2,3])
    with head_1:
        st.markdown(f"## 指标观测时间：{current_week}({st.session_state.t0_date.strftime('%m-%d')})")
    with head_2:
        toc_html = """
            <style>
                .toc-container {
                    display: flex;           /* 开启弹性布局 */
                    flex-direction: row;     /* 横向排列 */
                    flex-wrap: wrap;         /* 换行排版，防止手机端溢出 */
                    gap: 12px;               /* 按钮之间的间距 */
                    padding: 10px 0;
                    justify-content: flex-start; /* 靠左对齐，也可改为 space-between */
                }
                .toc-item {
                    display: flex;
                    align-items: center;
                    justify-content: center; /* 内容居中 */
                    padding: 8px 16px;       /* 调整内边距，使其更像水平标签 */
                    background-color: #f8f9fa;
                    border-radius: 8px;
                    text-decoration: none;
                    color: #31333F !important;
                    font-weight: 500;
                    font-size: 14px;         /* 稍微缩小字体以适应横排 */
                    transition: all 0.2s ease;
                    border: 1px solid #e0e0e0;
                    white-space: nowrap;     /* 禁止文字换行 */
                    flex: 1;                 /* 选填：让所有按钮等宽 */
                    min-width: 120px;        /* 按钮最小宽度 */
                    max-width: 200px;        /* 按钮最大宽度 */
                }
                .toc-item:hover {
                    background-color: #ffffff;
                    border-color: #ff4b4b;
                    box-shadow: 0 4px 8px rgba(255, 75, 75, 0.1);
                    transform: translateY(-2px); /* 悬停改为向上微动，横向更自然 */
                    color: #ff4b4b !important;
                }
                .toc-icon {
                    margin-right: 8px;
                    font-size: 1.1rem;
                }
                /* 适配移动端：如果屏幕太窄，自动变成两行 */
                @media (max-width: 768px) {
                    .toc-item {
                        flex: 1 1 45%; 
                    }
                }
            </style>
            <div class="toc-container">
                <a href="#0" class="toc-item">
                    <span class="toc-icon">📊</span> 指标概况
                </a>
                <a href="#1" class="toc-item">
                    <span class="toc-icon">📉</span> 海外库存周转
                </a>
                <a href="#2" class="toc-item">
                    <span class="toc-icon">🚚</span> 发货监控
                </a>
                <a href="#3" class="toc-item">
                    <span class="toc-icon">📈</span> 预测偏差
                </a>
                <a href="#4" class="toc-item">
                    <span class="toc-icon">🛠️</span> 干预偏差
                </a>
            </div>
            """
        st.markdown(toc_html, unsafe_allow_html=True)

    # 筛选面板逻辑
    if st.session_state.df_fahuo is not None:
        df_ref = st.session_state.df_fahuo
        ver = st.session_state.filter_ver
        c1, c2, c3, c4, c5 = st.columns([3, 3, 3, 1, 1], vertical_alignment="bottom")
        with c1:
            ui_market = st.multiselect("🌍 子市场", sorted(df_ref['子市场'].unique().tolist()), key=f"m_{ver}")
        with c2:
            cat_opts = df_ref[df_ref['子市场'].isin(ui_market)]['品类'].unique() if ui_market else df_ref[
                '品类'].unique()
            ui_category = st.multiselect("📦 品类", sorted(cat_opts.tolist()), key=f"c_{ver}")
        with c3:
            sku_opts = df_ref[df_ref['品类'].isin(ui_category)]['主料mrpsku'].unique() if ui_category else df_ref[
                '主料mrpsku'].unique()
            ui_sku = st.multiselect("🔑 MRPSKU", sorted(sku_opts.tolist()), key=f"s_{ver}")
        with c4:
            if st.button("🚀 确认", width='stretch', type="primary"):
                st.session_state.committed_filters = {"sub_market": ui_market, "category": ui_category,
                                                      "mrpsku": ui_sku}
                st.rerun()
        with c5:
            if st.button("重置", width='stretch'):
                st.session_state.filter_ver += 1
                st.session_state.committed_filters = {"sub_market": [], "category": [], "mrpsku": []}
                st.rerun()
                
    if st.session_state.df_stock_turnover is not None:
        filters = st.session_state.committed_filters
        df_stock_turnover_kpi = st.session_state.df_stock_turnover.copy()
        if filters["sub_market"]: df_stock_turnover_kpi = df_stock_turnover_kpi[df_stock_turnover_kpi["子市场"].isin(filters["sub_market"])]
        if filters["category"]: df_stock_turnover_kpi = df_stock_turnover_kpi[df_stock_turnover_kpi["品类"].isin(filters["category"])]
        if filters["mrpsku"]: df_stock_turnover_kpi = df_stock_turnover_kpi[df_stock_turnover_kpi["主料mrpsku"].isin(filters["mrpsku"])]

        df_yuce_kpi = st.session_state.df_yuce.copy()
        if filters["sub_market"]: df_yuce_kpi = df_yuce_kpi[df_yuce_kpi["子市场"].isin(filters["sub_market"])]
        if filters["category"]: df_yuce_kpi = df_yuce_kpi[df_yuce_kpi["品类"].isin(filters["category"])]
        if filters["mrpsku"]: df_yuce_kpi = df_yuce_kpi[df_yuce_kpi["主料mrpsku"].isin(filters["mrpsku"])]

        df_ganyu_kpi = st.session_state.df_ganyu.copy()
        if filters["sub_market"]: df_ganyu_kpi = df_ganyu_kpi[df_ganyu_kpi["子市场"].isin(filters["sub_market"])]
        if filters["category"]: df_ganyu_kpi = df_ganyu_kpi[df_ganyu_kpi["品类"].isin(filters["category"])]
        if filters["mrpsku"]: df_ganyu_kpi = df_ganyu_kpi[df_ganyu_kpi["主料mrpsku"].isin(filters["mrpsku"])]

        df_ganyu_bi = st.session_state.df_ganyu_bi.copy()
        
        df_country_turnover = st.session_state.df_country_turnover.copy()
        if filters["sub_market"]: df_country_turnover = df_country_turnover[df_country_turnover["子市场"].isin(filters["sub_market"])]
        if filters["category"]: df_country_turnover = df_country_turnover[df_country_turnover["品类"].isin(filters["category"])]
        if filters["mrpsku"]: df_country_turnover = df_country_turnover[df_country_turnover["主料mrpsku"].isin(filters["mrpsku"])]

        with st.container(border=True):
            st.markdown("**🎯 指标概况**")
            curr_avg_yuce = df_yuce_kpi['单周预测偏差率'].abs().mean()
            curr_avg_huanbiyuce = df_yuce_kpi['环比预测偏差率'].abs().mean()
            curr_avg_ganyu = df_ganyu_kpi['单周干预偏差率'].abs().mean()
            curr_avg_huanbiganyu = df_ganyu_kpi['环比干预偏差率'].abs().mean()
            ganyu_intervention_rate = df_ganyu_bi["有干预样本数"].sum() / df_ganyu_bi["总样本数"].sum()
            last_dt_history = st.session_state.t0_date.strftime("%Yw%V")
            历史国内在库周转 = ((((df_country_turnover['当周期初在库'] * df_country_turnover['单价']).sum() + (df_country_turnover['下周期初在库'] * df_country_turnover['单价']).sum()) / 2) / ((df_country_turnover['当周周销']*df_country_turnover['单价'])/7).sum()).round(1)
            df_temp_history = df_stock_turnover_kpi[df_stock_turnover_kpi['周数']==last_dt_history]
            历史海外在库周转 = (((((df_temp_history['当周期初在库'] * df_temp_history['单价']).sum() + (df_temp_history['下周期初在库'] * df_temp_history['单价']).sum()) / 2) / ((df_temp_history['当周周销']*df_temp_history['单价'])/7).sum())).round(1)
            历史海外在途周转 = ((((df_temp_history['当周期初在途'] * df_temp_history['单价']).sum() + (df_temp_history['下周期初在途'] * df_temp_history['单价']).sum()) / 2) / ((df_temp_history['当周周销']*df_temp_history['单价'])/7).sum()).round(1)
            c1, c2, c3, c4,c5,c6,c7,c8 = st.columns(8)
            st.markdown(f"""
                <style>
                /* 1. 定位整个 Metric 容器，让所有子元素在纵向上居中对齐 */
                [data-testid="stMetric"] {{
                    display: flex;
                    flex-direction: column;
                    align-items: center;
                    justify-content: center;
                    width: 100%;
                }}

                /* 2. 强制 Label (标题) 居中 */
                [data-testid="stMetricLabel"] {{
                    display: flex;
                    justify-content: center;
                    width: 100%;
                }}

                /* 3. 强制 Value (大数字) 居中并处理变色 */
                [data-testid="stMetricValue"] {{
                    display: flex;
                    justify-content: center;
                    width: 100%;
                }}
                /* 4. 强制 Delta (目标气泡) 居中 */
                [data-testid="stMetricDelta"] {{
                    display: flex;
                    justify-content: center;
                    width: 100%;
                    margin-left: 0 !important;
                }}

                /* 5. 隐藏箭头 */
                [data-testid="stMetricDelta"] svg {{
                    display: none !important;
                }}

                /* 6. 移除气泡内部文字可能的左间距 */
                [data-testid="stMetricDelta"] > div {{
                    margin-left: 0 !important;
                }}
                </style>
                """, unsafe_allow_html=True)
            c1.metric("海外在库周转",历史海外在库周转,"目标: P1:45天,P2:30天",delta_color="green",help="(期初在库x单价+期末在库x单价)/2/(周销x单价/7)")
            c2.metric("海外在途周转",历史海外在途周转,"目标: 60天",delta_color="green",help="(期初在途x单价+期末在途x单价)/2/(周销x单价/7)")
            c3.metric("国内在库周转", f"{历史国内在库周转:.1f}", delta=f"目标: 60天", delta_color="green",help="(国内期初在库x单价+国内期末在库x单价)/2/(周销x单价/7)")
            c4.metric("预测偏差率", f"{curr_avg_yuce:.1%}",delta="目标: 35%",delta_color="green")
            c5.metric("预测偏差率(环比)", f"{curr_avg_huanbiyuce:.0%}",delta="目标: 5%",delta_color="green")
            c6.metric("干预SKU占比", f"{ganyu_intervention_rate:.0%}",delta="目标: 15%",delta_color="green")
            c7.metric("干预偏差率", f"{curr_avg_ganyu:.0%}",delta="目标: 30%",delta_color="green")
            c8.metric("干预偏差率(环比)", f"{curr_avg_huanbiganyu:.0%}",delta="目标: 5%",delta_color="green")    
            



# =========================================================
# 5、核心图表函数 (函数名保持不变，日期格式 yyyy-mm-dd)
# =========================================================
def apply_filters(df, filters):
    if df is None: return None
    if filters is None: return df
    df = df.copy()
    if filters.get("sub_market"): df = df[df["子市场"].isin(filters["sub_market"])]
    if filters.get("category"): df = df[df["品类"].isin(filters["category"])]
    if filters.get("mrpsku"): df = df[df["主料mrpsku"].isin(filters["mrpsku"])]
    return df


def plot_inventorySales_rate(df_kuxiao, filters):
    df = apply_filters(df_kuxiao, filters)
    if df is None or df.empty: return
    st.markdown("### 📦 海外库存周转")
    df_avg = df.groupby(['日期', '数据类型'])['库销比'].mean().reset_index().sort_values('日期')
    fig = go.Figure()
    colors = {"在库库销比": "#92ABDF", "发货前_在途库销比": "#F4B483", "发货后_在途库销比": "#FEDB63"}
    for dtype in df_avg['数据类型'].unique():
        sub = df_avg[df_avg['数据类型'] == dtype]
        fig.add_trace(go.Scatter(x=sub['日期'], y=sub['库销比'], mode='lines+markers', name=dtype,
                                 line=dict(width=3, color=colors.get(dtype))
        ))
        # ================= 新增：添加基准线 =================
        # 添加在库库销比基准线 (y=4)
        fig.add_hline(
            y=4,
            line_dash="dash",  # 设置为虚线
            line_color=colors.get("在库库销比", "#92ABDF"),  # 保持颜色一致
            annotation_text="在库基准线: 4",
            annotation_position="bottom right",  # 文字显示位置
            opacity=0.7  # 设置透明度，避免抢了主数据的视觉焦点
        )

        # 添加在途库销比基准线 (y=8)
        fig.add_hline(
            y=8,
            line_dash="dash",
            line_color=colors.get("发货前_在途库销比", "#FEDB63"),
            annotation_text="在途基准线: 8",
            annotation_position="top right",
            opacity=0.7
        )
    fig.update_layout(
        title=dict(text="在库在途库销比趋势", x=0.01, font=dict(size=16)),
        hovermode="x unified",
        margin=dict(l=50, r=20, t=60, b=50),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        plot_bgcolor='white',
        xaxis=dict(
            showgrid=True, 
            gridcolor='#F3F4F6', 
            type='date', 
            tickformat='%Y-%m-%d', 
            tickangle=-70,
            dtick=86400000.0, 
            tickvals=df_avg['日期'].unique().tolist(),
            automargin=True 
            ),
        yaxis=dict(
            showgrid=True, 
            gridcolor='#F3F4F6', 
            title="库销比", 
            tickformat=".1f",
            range=[0,15]
        ),
        font=dict(family="Microsoft YaHei"),
    )
    st.plotly_chart(fig, width='stretch')


def plot_delivery_bar_line(df_fahuo, filters):
    # 1. 应用筛选逻辑
    df = apply_filters(df_fahuo, filters)
    if df is None or df.empty: 
        st.warning("所选条件下无数据")
        return

    qty_cols = ["计划发货量", "配货数量", "排单数量", "拣货量", "实际出库量"]
    rate_cols = ["计划达成率", "配货达成率", "排单达成率", "拣货达成率", "出库达成率"]

    stage_names = ["计划", "配货", "排单", "拣货", "出库"]
    total_qty = df[qty_cols].sum().tolist()
    avg_rates = df[rate_cols].mean().tolist()
    display_rates = [r * 100 if max(avg_rates) <= 1.0 else r for r in avg_rates]

    st.metric(
        label="🎯 计划达成率", 
        value=f"{display_rates[0]:.2f}%"
    )
    fig = make_subplots(specs=[[{"secondary_y": False}]])
    fig.add_trace(
        go.Bar(
            y=stage_names,
            x=total_qty,
            name="执行数量",
            marker_color="#64b5f6",
            text=total_qty,
            textposition='auto',
            hovertemplate="数量: %{y:,.0f}<extra></extra>",
            textfont=dict(
                size=18,          
                color="black" 
            ),
            orientation='h'
        ),
        secondary_y=False,
    )

    fig.update_layout(
        title={
            'text': "发货过程执行数量",
            'y': 0.95, 'x': 0.5, 'xanchor': 'center', 'yanchor': 'top',
            'font': dict(size=20)
        },
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=50, r=50, t=100, b=50),
        hovermode="x unified",
        font=dict(family="Microsoft YaHei"),
    )

    fig.update_yaxes(title_text="<b>数量</b>", secondary_y=False, gridcolor='lightgrey',categoryorder='total ascending')
    st.plotly_chart(fig, width='stretch')

def plot_delivery_plan_rate(df_fahuo,filters):
    df = apply_filters(df_fahuo, filters)
    if df is None or df.empty: return
    cols = ["计划发货量", "实际出库量","计划达成率"]
    df_avg = df.groupby("日期")[cols].mean().reset_index().sort_values("日期")
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    for col in ["计划发货量", "实际出库量"]:
        fig.add_trace(
            go.Bar(x=df_avg["日期"], y=df_avg[col], name=col),
            secondary_y=False
        )
    fig.add_trace(
        go.Scatter(
            x=df_avg["日期"], 
            y=df_avg['计划达成率'], 
            mode='lines+markers', 
            name='计划达成率',
            line=dict(width=3, color="#3B82F6"),
        ), 
        secondary_y=True
    )
    fig.update_layout(
        title='计划达成率',
        xaxis_title='日期',
        hovermode="x unified",
        xaxis=dict(showgrid=True, gridcolor='#F3F4F6', type='date', tickformat='%Y-%m-%d', tickangle=-30),
        yaxis=dict(
        title_text="数量",
        tickformat=".2f" 
        ),
        yaxis2=dict(
        title_text="数量",
        tickformat=".2%" 
        ),
        font=dict(family="Microsoft YaHei"),
    )
    fig.update_yaxes(title_text="数量", secondary_y=False)
    fig.update_yaxes(title_text="计划达成率 (%)", secondary_y=True)
    st.plotly_chart(fig)

def plot_delivery_market_num(df_fahuo, filters):
    df = apply_filters(df_fahuo, filters)
    if df is None or df.empty: return
    market_list = sorted(df['子市场'].unique().tolist())
    selected_market = st.selectbox("请选择要分析的子市场", market_list)
    df_filtered = df[df['子市场'] == selected_market]
    stage_cols = ['计划发货量', '配货数量', '排单数量', '拣货量', '实际出库量']
    df_sum = df_filtered[stage_cols].sum().reset_index()
    df_sum.columns = ['执行环节', '数量']

    fig = px.bar(
        df_sum,
        x="数量",
        y="执行环节",
        orientation='h',
        text='数量',
        color='执行环节',
        category_orders={"执行环节": ['计划发货量', '配货数量', '排单数量', '拣货量', '实际出库量']},
        title=f"子市场 [{selected_market}] 发货全链路执行情况",
        color_discrete_sequence=px.colors.qualitative.Pastel # 使用较柔和的颜色
    )

    # 5. 美化调整
    fig.update_traces(
        textposition='inside', # 数值放在条形图内部
        texttemplate='%{text}', 
        hovertemplate="环节: %{y}<br>数量: %{x}<extra></extra>"
    )

    fig.update_layout(
        xaxis_title="数量",
        yaxis_title="",
        showlegend=False,
        height=450,
        plot_bgcolor='rgba(0,0,0,0)', # 透明背景
        margin=dict(l=150), # 留出左边文字空间
        font=dict(family="Microsoft YaHei"),
    )

    # 添加网格线辅助观察
    fig.update_xaxes(showgrid=True, gridcolor='LightGrey')

    # --- 第六步：展示 ---
    st.plotly_chart(fig, width='stretch')


def table_delivery_allmarket_rate(df_fahuo, filters):
    df = apply_filters(df_fahuo, filters)
    if df is None or df.empty: return
    
    stage_cols = ["计划达成率", "配货达成率", "排单达成率", "拣货达成率", "出库达成率"]
    df_avg = df.groupby("子市场").agg({
        "计划达成率": "mean",
        "配货达成率": "mean",
        "排单达成率": "mean",
        "拣货达成率": "mean",
        "出库达成率": "mean",
        "主料mrpsku": "count"
    }).reset_index()
    df_avg = df_avg.rename(columns={"主料mrpsku": "SKU数量"})
    df_avg = df_avg.sort_values(by="SKU数量", ascending=False)
    # 计算全市场平均值
    df_avg_allmarket = df[stage_cols].mean().reset_index()
    df_avg_allmarket[0] = df_avg_allmarket[0].astype(float)
    
    col1, col2 = st.columns([2.5, 2])
    thresholds = {
        "计划达成率": 0.9, 
        "配货达成率": 0.9,
        "排单达成率": 1.0,
        "拣货达成率": 1.0,
        "出库达成率": 1.0
    }
    def style_threshold_clean(col):
        if col.name in thresholds:
            target = thresholds[col.name]
            styles = []
            for v in col:
                if v >= target:
                    styles.append('color: #2E7D32;') 
                else:
                    styles.append('background-color: #FFF5F5; color: #C62828; font-weight: bold;')
            return styles
        return [''] * len(col)

    styled_df = (
        df_avg.style
        .apply(style_threshold_clean)
        .set_properties(**{
            'text-align': 'center',
        })
        .format("{:.1%}", subset=list(thresholds.keys()))
    )
    with col1:
        st.markdown("#### 📦 全市场发货过程指标全貌")
        c1,c2,c3,c4,c5 = st.columns(5)
        # 如果指标值小于目标就为红色，反之为绿色
        c1.metric("计划达成率", f"{df_avg_allmarket.iloc[0,1]:.1%}",delta="目标: 90%",delta_color="green")
        c2.metric("配货达成率", f"{df_avg_allmarket.iloc[1,1]:.1%}",delta="目标: 90%",delta_color="green")
        c3.metric("排单达成率", f"{df_avg_allmarket.iloc[2,1]:.1%}",delta="目标: 100%",delta_color="green")
        c4.metric("拣货达成率", f"{df_avg_allmarket.iloc[3,1]:.1%}",delta="目标: 100%",delta_color="green")
        c5.metric("出库达成率", f"{df_avg_allmarket.iloc[4,1]:.1%}",delta="目标: 100%",delta_color="green")
        st.dataframe(
            styled_df, 
            width='stretch', 
            column_config={
                "子市场": st.column_config.TextColumn("子市场", width="small"),
                **{
                    col: st.column_config.TextColumn(
                        label=f"{col}\n(目标:{thresholds[col]:.0%})", 
                        help=f"目标: {thresholds[col]:.1%}",
                        alignment="center"
                    ) for col in thresholds.keys()
                },
                "SKU数量": st.column_config.TextColumn("SKU数量", width="small",alignment="center"),
            },
            height=480,
            hide_index=True,
        )

    with col2:
        st.markdown("#### 📦 发货指标下钻分析-子市场")
        market_list =  ["全部市场"] + sorted(df['子市场'].unique().tolist()) 
        selected_market = st.selectbox("🎯 选择下钻的子市场", market_list)
        if selected_market == "全部市场":
            df_filtered=df.copy()
        else:
            df_filtered = df[df['子市场'] == selected_market]
        st.session_state.filter_market=selected_market
        stage_cols = ['计划发货量', '配货数量', '排单数量', '拣货量', '实际出库量']
        df_sum = df_filtered[stage_cols].sum().reset_index()
        df_sum.columns = ['执行环节', '数量']

        fig = px.bar(
            df_sum,
            x="数量",
            y="执行环节",
            orientation='h',
            text='数量',
            color='执行环节',
            category_orders={"执行环节": ['计划发货量', '配货数量', '排单数量', '拣货量', '实际出库量']},
            title=f"子市场 [{selected_market}] 发货全链路执行情况",
            color_discrete_sequence=px.colors.qualitative.Pastel
        )

        # 5. 美化调整
        fig.update_traces(
            textposition='inside', 
            # 【修改】使用 HTML 标签 <b> 让数值加粗，并设置 textfont 变大
            # texttemplate='<b>%{text}</b>', 
            texttemplate='<b>%{text:,}</b>', 
            textfont=dict(
                size=16,             # 设置条形图内部字体大小
                family="Microsoft YaHei",
                color="black",      # 根据背景调整颜色，通常内部用白色或黑色
            ),
            hovertemplate="环节: %{y}<br>数量: %{x}<extra></extra>"
        )

        fig.update_layout(
            xaxis_title="<b>数量</b>",  # 轴标题加粗
            yaxis_title="",
            showlegend=False,
            height=450,
            plot_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=150),
            # 【新增】全局字体设置
            font=dict(family="Microsoft YaHei"),
            # 【新增】设置标题加粗
            title_font=dict(size=20, family="Microsoft YaHei"),
        )

        # 【修改】坐标轴刻度字体加粗
        fig.update_xaxes(
            showgrid=True, 
            gridcolor='LightGrey',
            tickfont=dict(size=14, family="Microsoft YaHei", color="black"), # 刻度加粗
            title_font=dict(size=16, family="Microsoft YaHei")              # 轴标题加粗
        )
        
        fig.update_yaxes(
            tickfont=dict(size=15, family="Microsoft YaHei", color="black"), # 纵轴环节名称加粗变大
        )

        st.plotly_chart(fig, width='stretch')
        # plot_delivery_category_num(df_filtered)


def plot_top_categories_by_rate(df_fahuo, filters,filter_market):
    if filters is None:
        if filter_market == "全部市场":
            df=df_fahuo.copy()
        else:
            df=df_fahuo[df_fahuo['子市场'] == filter_market].copy()
    else:
        df = apply_filters(df_fahuo, filters)

    # 1. 指标与排序配置
    st.write(filter_market)
    st.write("没有姿势差")
    col_ctrl1, col_ctrl2 = st.columns([2.5,3.5])
    with col_ctrl1:
        # st.markdown(f"#### 📊 发货指标下钻分析-[{filter_market}] 品类")
        
        st.markdown(
            f"""
            #### 📊 发货指标下钻分析 - <span style='color: #ff4b4b;'>{filter_market}</span> 品类
            """, 
            unsafe_allow_html=True
        )
        st.write("没有姿势差")
        metric = st.selectbox(
            "🎯 选择分析指标", 
            ["计划达成率", "配货达成率", "排单达成率", "拣货达成率", "出库达成率"]
        )
        
        # 1. 数据处理
        df_cat_rate = df.groupby("品类")[metric].mean().reset_index()

        # 排序：按达成率降序排列，最好的在上面
        df_category_tab = df_cat_rate.sort_values(metric, ascending=False).reset_index(drop=True)
        
        thresholds = {
            "计划达成率": 0.9, "配货达成率": 0.9, "排单达成率": 1.0, "拣货达成率": 1.0, "出库达成率": 1.0
        }
        target_val = thresholds[metric]
        
        # 2. 计算配色与差值
        df_category_tab['是否达标'] = df_category_tab[metric] >= target_val
        df_category_tab['主颜色'] = df_category_tab['是否达标'].apply(lambda x: '#4CAABF' if x else '#FF6B6B')
        df_category_tab['背景色'] = '#EAEAEA'

        # 【优化】动态高度计算：每行40px + 顶部40px，取消max(500)防止品类少时留白过多
        chart_real_height = (len(df_category_tab) * 40) + 40

        # 3. 创建子图
        fig = make_subplots(
            rows=1, cols=2, 
            shared_yaxes=True, 
            column_widths=[0.65, 0.35], # 稍微拓宽右侧空间显示差值
            horizontal_spacing=0.01
        )

        # 4. 添加目标背景条
        fig.add_trace(go.Bar(
            y=df_category_tab['品类'], 
            x=[target_val] * len(df_category_tab),
            orientation='h',
            marker_color=df_category_tab['背景色'],
            name='目标要求',
            width=0.7,
            hoverinfo='skip'
        ), row=1, col=1)

        # 5. 添加实际达成条
        fig.add_trace(go.Bar(
            y=df_category_tab['品类'], 
            x=df_category_tab[metric],
            orientation='h',
            marker_color=df_category_tab['主颜色'],
            name='实际达成',
            width=0.4,
            text=[f"<b>{x:.1%}</b>" for x in df_category_tab[metric]],
            textposition='inside',
            textfont=dict(size=12, color='white'),
        ), row=1, col=1)

        # 6. 【关键修改】处理右侧文本：显示实际值及与目标的差值
        status_text = []
        for v, d in zip(df_category_tab[metric], df_category_tab['是否达标']):
            diff = v - target_val # 计算差值
            color = "#4CAABF" if d else "#FF6B6B"
            icon = "✔" if d else "✘"
            # 格式化差值：+1.2% 或 -3.5%
            diff_str = f"{diff:+.1%}"
            status_text.append(f"<b>{v:.1%}</b> <span style='font-size:12px; color:{color}'>({diff_str}) {icon}</span>")

        fig.add_trace(go.Scatter(
            y=df_category_tab['品类'], 
            x=[0.01] * len(df_category_tab), 
            mode='markers+text',
            marker=dict(symbol='circle-open', size=18, color=df_category_tab['主颜色'], line_width=2),
            text=status_text,
            textposition='middle right',
            textfont=dict(size=13, color='#333333', family="Microsoft YaHei"),
            showlegend=False
        ), row=1, col=2)

        # 7. 【优化】布局减少留白
        fig.update_layout(
            barmode='overlay',
            plot_bgcolor='white',
            height=chart_real_height,
            # 大幅减小上下左留白：l(左)=80, r=10, t(上)=30, b(下)=0
            margin=dict(l=60, r=10, t=0, b=0), 
            showlegend=False,
            font=dict(family="Microsoft YaHei"),
            xaxis=dict(visible=False, range=[0, max(df_category_tab[metric].max(), target_val) * 1.1]),
            xaxis2=dict(visible=False),
            yaxis=dict(
                autorange="min reversed",  # 保持升序排列
                tickmode="linear",
                showgrid=False,
                tickfont=dict(size=12, weight='bold', color='#333333')
            )
        )

        # 8. 在 Streamlit 容器中显示
        # 容器高度自适应，最高500
        with st.container(height=min(500, chart_real_height + 50)):
            st.plotly_chart(fig, width='stretch', config={'displayModeBar': False})
    
    with col_ctrl2:
        st.markdown(f"#### 🔍发货指标下钻分析-[{filter_market}] MRPSKU")
        if st.session_state.filter_market=="全部市场":
            df_market=df.copy()
        else:
            df_market=df[df['子市场'] == filter_market]
        
        # st.session_state.filter_market = filter_market
        # 1. 获取该市场下的所有品类，供用户选择
        # df_market = df[df['子市场'] == filter_market]
        # df_market = df.copy()
        category_list = sorted(df_market['品类'].unique().tolist())
        
        col1, col2 = st.columns([1, 3])
        with col1:
            selected_cat = st.selectbox("🎯 选定下钻的品类", ["全部品类"] + category_list)

        # 2. 过滤数据
        if selected_cat == "全部品类":
            df_sku = df_market.copy()
        else:
            df_sku = df_market[df_market['品类'] == selected_cat].copy()

        # 3. 计算关键指标（可选：为了让用户知道在这个品类下有多少坑）
        total_skus = len(df_sku)
        problem_skus = len(df_sku[df_sku['实际出库量'] < df_sku['计划发货量']])
        
        with col2:
            # 用小组件显示概况
            m1, m2, m3 = st.columns(3)
            m1.metric("SKU 总数", total_skus)
            m2.metric("异常 SKU 数", problem_skus)
            m3.write("💡 *异常定义：实际出库量 < 计划发货量*")
        display_cols = [
            "子市场","主料mrpsku", "品类", "运输方式","计划发货量", "配货数量", "排单数量", "拣货量", "实际出库量"
        ]
        df_sku['缺口'] = df_sku['计划发货量'] - df_sku['实际出库量']
        df_sku_sorted = df_sku.sort_values("缺口", ascending=False)
        rate_cols = ["计划达成率", "配货达成率", "排单达成率", "拣货达成率", "出库达成率"]
        
        styled_sku = (
            df_sku_sorted[display_cols].style
        )

        # 6. 展示表格
        st.dataframe(
            styled_sku, 
            width='stretch', 
            height=500,
            column_config={
                "主料mrpsku": st.column_config.TextColumn("SKU 编码", width="medium"),
                "计划发货量": st.column_config.NumberColumn("计划发货量", format="%d", alignment="center"),
                "配货数量": st.column_config.NumberColumn("配货数量", format="%d", alignment="center"),
                "排单数量": st.column_config.NumberColumn("排单数量", format="%d", alignment="center"),
                "拣货量": st.column_config.NumberColumn("拣货量", format="%d", alignment="center"),
                "实际出库量": st.column_config.NumberColumn("实际出库量", format="%d", alignment="center"),
            },
            hide_index=True
        )


def inventorySales_rate_area(df_stock_turnover, curr_filters):
    df = apply_filters(df_stock_turnover, curr_filters)
    last_dt_history = st.session_state.t0_date.strftime("%Yw%V")
    
    # 未来库存与周转
    
    market_filter, category_filter, sku_filter = st.columns(3)
    with market_filter:
        market_list_stock = ["全部市场"] + sorted(df["子市场"].unique().tolist())
        selected_market = st.selectbox("选择要查看的子市场", market_list_stock,key="selectbox_market_stock")
        st.session_state.stock_filter_market=selected_market
    with category_filter:
        category_list_stock = ["全部品类"] + sorted(df["品类"].unique().tolist())
        selected_category = st.selectbox("选择要查看的品类", category_list_stock,key="selectbox_category_stock")
        st.session_state.stock_filter_category=selected_category
    with sku_filter:
        sku_list_stock = ["全部SKU"] + sorted(df["主料mrpsku"].unique().tolist())
        selected_sku = st.selectbox("选择要查看的SKU", sku_list_stock,key="selectbox_sku_stock")
        st.session_state.stock_filter_sku=selected_sku
    if selected_market == "全部市场":
        df_filtered = df
    elif selected_market != "全部市场" and selected_category == "全部品类":
        df_filtered = df[(df["子市场"] == selected_market)]
    elif selected_market != "全部市场" and selected_category != "全部品类":
        df_filtered = df[(df["子市场"] == selected_market) & (df["品类"] == selected_category)]
    elif selected_market != "全部市场" and selected_category != "全部品类" and selected_sku != "全部SKU":
        df_filtered = df[(df["子市场"] == selected_market) & (df["品类"] == selected_category) & (df["主料mrpsku"] == selected_sku)]

    df_temp_future = df_filtered[(df_filtered['周数']>last_dt_history)].copy()
    df_temp_future['期初在库金额'] = df_temp_future['当周期初在库']*df_temp_future['单价']
    df_temp_future['期末在库金额'] = df_temp_future['下周期初在库']*df_temp_future['单价']
    df_temp_future['期初在途金额'] = df_temp_future['当周期初在途']*df_temp_future['单价']
    df_temp_future['期末在途金额'] = df_temp_future['下周期初在途']*df_temp_future['单价']
    df_temp_future['当周周销金额'] = df_temp_future['当周周销']*df_temp_future['单价']
    df_temp_future_group = df_temp_future.groupby(['周数'])[['当周期初在库','下周期初在库','当周期初在途',"下周期初在途","SLT平均周销","单价","期初在库金额","期末在库金额","期初在途金额","期末在途金额","当周到货","目标平均周销","SLT+1周期初在库","SLT+1周期初在途","当周周销金额","当周周销"]].sum().reset_index().sort_values('周数')
    df_temp_future_group['未来海外在库周转'] = (((df_temp_future_group['期初在库金额'] + df_temp_future_group['期末在库金额'])/2) / (df_temp_future_group['当周周销金额'] / 7)).round(1)
    df_temp_future_group['未来海外在途周转'] = (((df_temp_future_group['期初在途金额'] + df_temp_future_group['期末在途金额'])/2) / (df_temp_future_group['当周周销金额'] / 7)).round(1)
    df_temp_future_group['DOS1'] = ((df_temp_future_group['当周期初在库'] / df_temp_future_group['目标平均周销']) * 7).round(1)
    df_temp_future_group['DOS2'] = ((df_temp_future_group['当周期初在库'] + df_temp_future_group['当周到货']) / df_temp_future_group['目标平均周销'] * 7).round(1)
    df_temp_future_group['库销比天数'] = ((df_temp_future_group['当周期初在库'] / df_temp_future_group['当周周销']) * 7).round(1)

    st.markdown("### 看未来")
    df_temp_future_group=df_temp_future_group[(df_temp_future_group['周数']<='2026w53')]
    fig = make_subplots(
            rows=1, cols=2, 
            shared_yaxes=False,    
            horizontal_spacing=0.08,
            subplot_titles=("海外在库周转", "海外在途周转")
        )

    metrics = ['未来海外在库周转', '未来海外在途周转']
    colors = ['#1f77b4', '#ff7f0e']

    week_values = df_temp_future_group["周数"].values
    week_label = [str(x)[2:] for x in week_values]

    # 2. 循环添加曲线 (row 固定为 1，col 随循环变化)
    for i, col_name in enumerate(metrics):
        col_idx = i + 1
        fig.add_trace(
            go.Scatter(
                x=week_values, 
                y=df_temp_future_group[col_name],
                name=col_name,
                mode='lines+markers',
                line=dict(color=colors[i], width=2),
                marker=dict(size=6),
                hovertemplate="周数: %{x}<br>数值: %{y}<extra></extra>"
            ),
            row=1, col=col_idx
        )

    for i, annotation in enumerate(fig['layout']['annotations']):
        annotation['font'] = dict(family="Microsoft YaHei", size=20, color="black")
    
    fig.add_hline(y=45,line_dash="dash",line_color="#1f77b4", line_width=2,row=1, col=1,annotation_text="P1目标：45天",annotation_position="bottom right")
    fig.add_hline(y=30,line_dash="dash",line_color="#1f77b4", line_width=2,row=1, col=1,annotation_text="P2目标：30天",annotation_position="bottom right")
    fig.add_hline(y=60,line_dash="dash",line_color="#ff7f0e", line_width=2,row=1, col=2,annotation_text="目标：60天",annotation_position="bottom right")
    # 3. 布局优化
    fig.update_layout(
        height=400,
        showlegend=False,
        template="simple_white",
        hovermode="x unified",
        margin=dict(t=60, b=50, l=40, r=40),
        font=dict(family="Microsoft YaHei",size=10),
    )
    fig.update_xaxes(
        ticktext=week_label,
        tickvals=week_values,
        tickfont=dict(size=12, color="gray"),
        title_text="周数",
        showline=True,
        linewidth=1,
        linecolor='lightgray'
    )

    # 5. 美化所有 Y 轴
    fig.update_yaxes(
        showgrid=True, 
        gridcolor='whitesmoke',
        # zeroline=True,
        # zerolinecolor='lightgray',
        # 设定范围为[0,100]
        range=[0, 100]
    )

    # 6. 在 Streamlit 中显示
    st.plotly_chart(fig, width='stretch')
    st.markdown("### 未来库存、销量与到货")

    fig = go.Figure()
    df_temp_future_group=df_temp_future_group[(df_temp_future_group['周数']<='2026w53')]

    color_stock = "#1f77b4"
    color_replenish = "#ff7f0e"
    color_sales = "#009966"
    fig.add_trace(go.Scatter(
        x=df_temp_future_group["周数"], 
        y=df_temp_future_group["当周期初在库"],
        mode='lines+text+markers',
        name='库存数量',
        line=dict(color=color_stock, width=3, shape='spline'),
        text=[f"{row:,.0f}" for row in df_temp_future_group['当周期初在库']],
        textposition="bottom center",
        textfont=dict(color=color_stock)
    ))

    fig.add_trace(go.Scatter(
        x=df_temp_future_group["周数"], 
        y=df_temp_future_group["当周到货"],
        mode='lines+text+markers',
        name='到货数量',
        line=dict(color=color_replenish, width=3, shape='spline'),
        text=[f"{row:,.0f}" for row in df_temp_future_group['当周到货']],
        textposition="top center",
        textfont=dict(color=color_replenish)
    ))
    fig.add_trace(go.Scatter(
        x=df_temp_future_group["周数"], 
        y=df_temp_future_group["当周周销"],
        mode='lines+text+markers',
        name='周销数量',
        line=dict(color=color_sales, width=3, shape='spline'),
        text=[f"{row:,.0f}" for row in df_temp_future_group['当周周销']],
        textposition="bottom center",
        textfont=dict(color=color_sales)
    ))
    top_y1 = max(df_temp_future_group["当周期初在库"].max(), df_temp_future_group["当周周销"].max()) * 1.4
    for i, row in df_temp_future_group.iterrows():
        start_y = max(row["当周期初在库"], row["当周周销"])  
        fig.add_shape(
            type="line",
            x0=row["周数"], y0=start_y + 10,
            x1=row["周数"], y1=top_y1 - 20,
            line=dict(color="gray", width=1, dash="dot"),
        )
        fig.add_trace(go.Scatter(
            x=[row["周数"]],
            y=[top_y1],
            mode="markers+text",
            marker=dict(color="white", size=10, line=dict(color="#52c41a", width=2)),
            text=[f"{row['库销比天数']:.0f}"],
            textposition="top center",
            textfont=dict(color="#52c41a", size=14, family="Arial Black"),
            showlegend=i == 0,
            hoverinfo='skip',
            name='库销比天数'
        ))
    new_labels = [str(x)[2:] for x in df_temp_future["周数"].values]
    fig.update_layout(
        plot_bgcolor="white",
        height=500,
        margin=dict(t=80, b=20, l=20, r=20),
        xaxis=dict(
            showgrid=False,
            showline=True,
            linecolor="lightgray",
            ticktext=new_labels,  
            tickvals=df_temp_future["周数"].values,
            tickfont=dict(size=14, color="gray")
        ),
        yaxis=dict(
            showgrid=True,
            showticklabels=True
        ),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hovermode="x unified",
        font=dict(family="Microsoft YaHei",size=10),
    )
    st.plotly_chart(fig, width='stretch')

    stock_col1,stock_col2 = st.columns([1,2])
    df_future_single = df_filtered[(df_filtered['周数']>last_dt_history)].copy()
    df_future_single['期初在库金额'] = df_future_single['当周期初在库']*df_future_single['单价']
    df_future_single['期末在库金额'] = df_future_single['下周期初在库']*df_future_single['单价']
    df_future_single['期初在途金额'] = df_future_single['当周期初在途']*df_future_single['单价']
    df_future_single['期末在途金额'] = df_future_single['下周期初在途']*df_future_single['单价']
    df_future_single['当周周销金额'] = df_future_single['当周周销']*df_future_single['单价']
    df_future_single['未来海外在库周转'] = (((df_future_single['期初在库金额'] + df_future_single['期末在库金额'])/2) / (df_future_single['当周周销金额'] / 7)).round(1)
    df_future_single['未来海外在途周转'] = (((df_future_single['期初在途金额'] + df_future_single['期末在途金额'])/2) / (df_future_single['当周周销金额'] / 7)).round(1)
    # df_future_single_oneweek = df_future_single[df_future_single['周数']==min(df_future_single['周数'])]
    df_future_single_oneweek=df_future_single.groupby(['子市场','品类','主料mrpsku'])[['未来海外在库周转','未来海外在途周转']].mean().reset_index()
    周转SKU总数 = len(df_future_single_oneweek)
    df_future_single_oneweek['未来海外在库周转区间'] = pd.cut(
        df_future_single_oneweek['未来海外在库周转'], 
        bins=[0, 15, 30, 45, float('inf')], 
        labels=['[0,15]', '[15,30]', '[30,45]', '[45,+inf]']
    )
    df_future_single_oneweek['未来海外在途周转区间'] = pd.cut(
        df_future_single_oneweek['未来海外在途周转'], 
        bins=[0, 30,45, 60, float('inf')], 
        labels=['[0,30]','[30,45]', '[45,60]', '[60,+inf]']
    )
    df_future_single_oneweek['下一步动作']="调整"
    with stock_col1:
        st.markdown("### 海外在库在途周转区间明细")
        # stock_counts = df_future_single_oneweek['未来海外在库周转区间'].value_counts()
        # transit_counts = df_future_single_oneweek['未来海外在途周转区间'].value_counts()


        # # 提取数值
        # stock_0_45 = stock_counts.get('达标', 0)
        # stock_45_rate = stock_0_45 / 周转SKU总数
        # stock_45_plus = stock_counts.get('不达标', 0)
        # stock_45_rate_plus = stock_45_plus / 周转SKU总数

        # transit_0_60 = transit_counts.get('达标', 0)
        # transit_60_rate = transit_0_60 / 周转SKU总数
        # transit_60_plus = transit_counts.get('不达标', 0)
        # transit_60_rate_plus = transit_60_plus / 周转SKU总数

        # # 2. 构建图形
        # fig_turnover = go.Figure()
        # fig_turnover.add_trace(go.Bar(
        #     name='SKU总数',
        #     x=['SKU总数'],
        #     y=[周转SKU总数],
        #     text=[周转SKU总数],
        #     textposition='outside',
        #     marker_color="#A6ACAF",
        #     textfont=dict(size=14,family="Microsoft YaHei")
        # ))
        # # 同时显示个数和比例
        # fig_turnover.add_trace(go.Bar(
        #     name='达标区间',
        #     x=['在库周转', '在途周转'],
        #     y=[stock_0_45, transit_0_60],
        #     # text=[stock_0_45, transit_0_60],
        #     text=[f"{stock_0_45}, {stock_45_rate:.0%}", f"{transit_0_60}, {transit_60_rate:.0%}"],
        #     textposition='inside',
        #     marker_color='#5da9c4',
        #     textfont=dict(size=14,family="Microsoft YaHei")
        # ))
        # fig_turnover.add_trace(go.Bar(
        #     name='不达标区间',
        #     x=['在库周转', '在途周转'],
        #     y=[stock_45_plus, transit_60_plus],
        #     # text=[stock_45_plus, transit_60_plus],
        #     text=[f"{stock_45_plus}, {stock_45_rate_plus:.0%}", f"{transit_60_plus}, {transit_60_rate_plus:.0%}"],
        #     textposition='outside',
        #     marker_color='#ff9d4f',
        #     textfont=dict(size=14,family="Microsoft YaHei")
        # ))
        # def add_proportion_annotation(fig, x_pos, y_val, text, ay):
        #     fig.add_annotation(
        #         x=x_pos, y=y_val,
        #         text=text,
        #         showarrow=True,
        #         arrowhead=0,     # 0表示直线，没有箭头尖
        #         arrowwidth=1,
        #         arrowcolor="#666",
        #         ax=-60,          # 线条向左延伸40个单位
        #         ay=ay,           # 垂直偏移
        #         font=dict(size=10, color="#666"),
        #         xanchor="right"  # 文字在短线左侧
        #     )

        # # 为“在库周转”添加比例标注
        # # 标注不达标比例（指向红段中间）
        # add_proportion_annotation(fig_turnover, '在库周转',周转SKU总数/2 - (stock_45_plus/2), f"{stock_45_plus/周转SKU总数*100:.1f}%", 0)
        # # 标注达标比例（指向绿段中间）
        # add_proportion_annotation(fig_turnover, '在库周转', stock_0_45/2, f"{stock_0_45/周转SKU总数*100:.1f}%", 0)

        # 为“在途周转”添加比例标注
        # add_proportion_annotation(fig_turnover, '在途周转',周转SKU总数/2 - (transit_60_plus/2), f"{transit_60_plus/周转SKU总数*100:.1f}%", 0)
        # add_proportion_annotation(fig_turnover, '在途周转', transit_0_60/2, f"{transit_0_60/周转SKU总数*100:.1f}%", 0)

        # 3. 设置为堆积模式并添加总数标注（可选）
        # fig_turnover.update_layout(
        #     barmode='stack',  # 关键：设置为堆积模式
        #     margin=dict(l=10, r=10, t=30, b=10),
        #     xaxis_title="周转类型",
        #     yaxis_title="SKU数量",
        #     legend_title="统计区间",
        #     hovermode='x unified',
        #     bargap=0.3, # 取值 0-1，越小柱子越宽
        #     legend=dict(
        #         orientation="h",   # 将图例改为水平显示
        #         yanchor="bottom",
        #         y=1.02,           # 把图例放到图表上方，腾出左右空间
        #         xanchor="right",
        #         x=1
        #     ),
        #     font=dict(size=16,family="Microsoft YaHei"),
        # )
        stock_counts = df_future_single_oneweek['未来海外在库周转区间'].value_counts().sort_index().reset_index()
        stock_counts.columns = ['区间', '数量']
        stock_counts['维度'] = '未来海外在库'  # 标记为在库维度

        transit_counts = df_future_single_oneweek['未来海外在途周转区间'].value_counts().sort_index().reset_index()
        transit_counts.columns = ['区间', '数量']
        transit_counts['维度'] = '未来海外在途' 
        plot_df = pd.concat([stock_counts, transit_counts]).dropna(subset=['区间'])
        color_map = {
            '[0,15]': '#a569bd', '[15,30]': '#5dade2', '[30,45]': '#58D68D', '[45,+inf]': '#e74c3c',
            '[0,30]': '#a569bd', '[45,60]': '#58D68D', '[60,+inf]': '#e74c3c'
        }
        # 将颜色映射到 DataFrame 中
        plot_df['颜色'] = plot_df['区间'].map(color_map)
        fig_turnover=go.Figure()
        fig_turnover.add_trace(go.Bar(
            name='SKU总数',
            x=['SKU总数'],
            y=[周转SKU总数],
            text=[周转SKU总数],
            textposition='outside',
            marker_color="#A6ACAF",
            textfont=dict(size=14,family="Microsoft YaHei", color="black")
        ))
        color_map = {
            '[0,15]': '#a569bd', '[15,30]': '#5dade2', '[30,45]': '#58D68D', '[45,+inf]': '#e74c3c',
            '[0,30]': '#a569bd', '[45,60]': '#58D68D', '[60,+inf]': '#e74c3c'
        }
        for interval in plot_df['区间'].unique():
            interval_data = plot_df[plot_df['区间'] == interval]
            fig_turnover.add_trace(go.Bar(
                x=interval_data['维度'],       # X轴：在库、在途
                y=interval_data['数量'],       # Y轴：对应的数量
                name=interval,                 # 命名：当前区间名（图例会显示这个）
                text=interval_data['数量'],    # 显示数量文字
                textposition='auto',         # 堆积图建议文字放内部
                marker_color=color_map.get(interval, '#999999'), # 获取当前区间的专属颜色
                textfont=dict(size=14, family="Microsoft YaHei", color="black"),
        ))

        fig_turnover.update_layout(
            barmode='stack',
            # title="未来海外周转区间分布",
            xaxis_title="",
            yaxis_title="SKU数量",
            legend_title="周转区间",
            height=500
        )
        
        st.plotly_chart(fig_turnover, width='stretch',height=500)
    with stock_col2:
        st.markdown("### 海外在库在途周转SKU明细")
        st.dataframe(
            df_future_single_oneweek,
            column_config={
                "子市场": st.column_config.TextColumn("子市场", width=20),
                "品类": st.column_config.TextColumn("品类", width=20),
                "主料mrpsku": st.column_config.TextColumn("MRPSKU", width=30),
                "未来海外在库周转": st.column_config.NumberColumn("海外在库周转(目标:45天)", format="%d", width=20, alignment="center"),
                "未来海外在途周转": st.column_config.NumberColumn("海外在途周转(目标:60天)", format="%d", width=20, alignment="center"),
                "未来海外在库周转区间": st.column_config.TextColumn("海外在库周转区间", width=20),
                "未来海外在途周转区间": st.column_config.TextColumn("海外在途周转区间", width=20)
            },
            hide_index=True,
            width="stretch",
            height=500
        )

  
# 预测指标区域
def predictSales_rate_area(df_yuce, curr_filters):
    df = apply_filters(df_yuce, curr_filters)
    if df is None or df.empty:
        st.warning("无数据")
        return

    # st.subheader("📊 预测准确度监控看板")
    # --- 1. 顶部 KPI 总览 ---
    
    market_filter, category_filter, sku_filter,outStock_filter,week_filter = st.columns(5)
    with market_filter:
        market_list = ["全部市场"] + sorted(df["子市场"].unique().tolist())
        selected_market = st.selectbox("选择要查看的子市场", market_list,key="selectbox_market_yuce")
        st.session_state.yuce_filter_market=selected_market
    with category_filter:
        category_list = ["全部品类"] + sorted(df["品类"].unique().tolist())
        selected_category = st.selectbox("选择要查看的品类", category_list,key="selectbox_category_yuce")
        st.session_state.yuce_filter_category=selected_category
    with sku_filter:
        SKU_list = ["全部SKU"] + sorted(df["主料mrpsku"].unique().tolist())
        selected_sku = st.selectbox("选择要查看的SKU", SKU_list,key="selectbox_sku_yuce")
        st.session_state.yuce_filter_sku=selected_sku
    with outStock_filter:
        outStock_status_list = ["全部状态","断货","非断货"]
        select_stockStatus = st.selectbox("选择SKU的状态", outStock_status_list,key="selectbox_outStock_yuce")
        st.session_state.yuce_filter_outStockStatus=select_stockStatus
    with week_filter:
        select_week = st.date_input("选择要查看的周数", value=default_monday,key="date_input_week_yuce")
        select_week = select_week.strftime("%Yw%V")
        st.session_state.yuce_filter_week=select_week

    if selected_market == "全部市场":
        df_filtered = df
    elif selected_market != "全部市场" and selected_category == "全部品类":
        df_filtered = df[(df["子市场"] == selected_market)]
    elif selected_market != "全部市场" and selected_category != "全部品类":
        df_filtered = df[(df["子市场"] == selected_market) & (df["品类"] == selected_category)]
    elif selected_market != "全部市场" and selected_category != "全部品类" and selected_sku != "全部SKU":
        df_filtered = df[(df["子市场"] == selected_market) & (df["品类"] == selected_category) & (df["主料mrpsku"] == selected_sku)]
    if select_stockStatus == "全部状态":
        df_filtered = df_filtered
    elif select_stockStatus == "断货":
        df_filtered = df_filtered[df_filtered["状态"] == "断货"]
    elif select_stockStatus == "非断货":
        df_filtered = df_filtered[df_filtered["状态"] != "断货"]

    # df_filtered_noOutStock = df_filtered[df_filtered["状态"] != "断货"]
    单周预测偏差 = df_filtered["单周预测偏差率"].abs().mean()
    环比预测偏差 = df_filtered["环比预测偏差率"].abs().mean()
    m_col1, m_col2 = st.columns(2)
    m_col1.metric("预测偏差率", f"{单周预测偏差:.0%}", delta="目标:35%")
    m_col2.metric("预测偏差率(环比)", f"{环比预测偏差:.0%}", delta="目标:5%")
    col_left, col_right = st.columns(2)
    stage_cols = ["单周预测偏差率", "环比预测偏差率"]
    df_filtered[stage_cols] = df_filtered[stage_cols].abs()
    df_m = df_filtered[df_filtered["周数"] == select_week].groupby("子市场").agg({
        "单周预测偏差率": "mean",
        "环比预测偏差率": "mean",
        "主料mrpsku": "count"
    }).reset_index()
    df_m = df_m.rename(columns={"主料mrpsku": "SKU数量"})
    df_m = df_m.sort_values(by="SKU数量", ascending=True)

    with col_left:
        st.markdown("#### 历史预测偏差趋势")
        df_filtered_week = df_filtered.groupby("周数").agg({
            "单周预测偏差率": "mean",
            "环比预测偏差率": "mean"
        }).reset_index()
        fig_历史预测曲线=go.Figure()
        fig_历史预测曲线.add_trace(
            go.Scatter(
                x=df_filtered_week["周数"],
                y=df_filtered_week["单周预测偏差率"],
                mode='lines+markers+text',
                line=dict(color='#1f77b4', width=2),
                marker=dict(
                    color="#1f77b4", 
                    size=20, 
                    line=dict(width=5, color='white')
                ),
                text=df_filtered_week["单周预测偏差率"].apply(lambda x: f"{x:.0%}"),
                textposition="top left",
                name="单周偏差"
            )
        )
        fig_历史预测曲线.add_trace(
            go.Scatter(
                x=df_filtered_week["周数"],
                y=df_filtered_week["环比预测偏差率"],
                mode='lines+markers+text',
                line=dict(color='#ff7f0e', width=2),
                marker=dict(
                    color="#ff7f0e", 
                    size=20, 
                    line=dict(width=5, color='white')
                ),
                text=df_filtered_week["环比预测偏差率"].apply(lambda x: f"{x:.0%}"),
                textposition="bottom left",
                name="环比偏差",
                yaxis="y2"
            )
        )
        fig_历史预测曲线.add_hline(
            y=0.35,
            line_dash="dash",
            line_color="#1f77b4",  
            annotation_text="预测偏差: 35%",
            annotation_position="bottom right",
            opacity=0.7
            
        )
        fig_历史预测曲线.add_shape(
            type="line",
            xref="paper", yref="y2",  # x轴跨越整个画布，y轴绑定到右轴 y2
            x0=0, y0=0.05,
            x1=1, y1=0.05,
            line=dict(color="#ff7f0e", width=2, dash="dash"),
            opacity=0.7
        )
        fig_历史预测曲线.update_layout(
            # 左侧 Y 轴配置
            yaxis=dict(
                title="单周预测偏差率",
                range=[0, 1.0],  
                tickformat=".0%", 
                side="left"
            ),
            # 右侧 Y 轴配置
            yaxis2=dict(
                title="环比预测偏差率",
                range=[0, 0.2],  
                tickformat=".0%", 
                overlaying="y",   
                side="right"
            ),
            # 可选：优化整体布局，防止右侧 Y 轴标题被截断
            margin=dict(r=80),
            legend=dict(
                orientation="h",          # 水平排列
                yanchor="bottom",         # 垂直方向以底部为锚点
                y=1.02,                   # 放在绘图区顶部稍微偏上的位置
                xanchor="center",         # 水平方向以中心为锚点
                x=0.5,                    # 放在水平方向 50% 的位置（居中）
                bgcolor="rgba(255,255,255,0.8)" # 可选：设置半透明背景，防止遮挡标题或边框
            ),
            font=dict(family="Microsoft YaHei",size=10)
        )

        st.plotly_chart(fig_历史预测曲线,width="stretch",height=550)


    with col_right:
        TARGET_MAX = 0.35
        st.markdown("#### 各市场预测偏差全貌 (目标区间:35%)")

        # --- 2. 创建布局 ---
        fig = make_subplots(
            rows=1, cols=2, 
            shared_yaxes=True, 
            horizontal_spacing=0.03, 
            column_widths=[0.7, 0.3]
        )
        colors_bar = [
            "#ff9d4f" if (x > TARGET_MAX) else "#5da9c4" 
            for x in df_m["单周预测偏差率"]
        ]

        fig.add_trace(
            go.Bar(
                y=df_m["子市场"],
                x=df_m["单周预测偏差率"],
                orientation='h',
                marker_color=colors_bar,
                text=df_m["单周预测偏差率"].apply(lambda x: f"{x:.1%}"),
                textposition='outside',
                cliponaxis=False,
                name="单周偏差"
            ),
            row=1, col=1
        )

        for line_x in [TARGET_MAX]:
            fig.add_vline(
                x=line_x, 
                line_dash="dash", 
                line_color="red", 
                line_width=2,
                row=1, col=1,
                # 不透明度
                opacity=0.8
            )


        # --- 5. 右侧：环比预测偏差率 (折线图) ---
        colors_line = ["#5da9c4" if abs(x) <= 0.05 else "#f5222d" for x in df_m["环比预测偏差率"]]

        fig.add_trace(
            go.Scatter(
                y=df_m["子市场"],
                x=df_m["环比预测偏差率"],
                mode='lines+markers+text',
                line=dict(color='#bfbfbf', width=3),
                marker=dict(
                    color=colors_line, 
                    size=20, 
                    line=dict(width=5, color='white')
                ),
                text=df_m["环比预测偏差率"].apply(lambda x: f"{x:.0%}"),
                textposition="middle right",
                name="环比趋势"
            ),
            row=1, col=2
        )
        for line_x in [0.05, -0.05]:
            fig.add_vline(
                x=line_x, 
                line_dash="dash", 
                line_color="red", 
                line_width=2,
                row=1, col=2,
                # 不透明度
                opacity=0.8
            )

        # --- 6. 布局精修 ---
        # 处理 UK 等极端值：确保坐标轴能盖住 ±30%
        max_val = max(df_m["单周预测偏差率"].max(), 0.5)

        fig.update_layout(
            template="simple_white",
            showlegend=False,
            height=len(df_m) * 45 + 120,
            margin=dict(l=10, r=60, t=20, b=40),
            xaxis=dict(
                title="单周偏差(35%)",
                tickformat=".0%",
                range=[0, max_val + 0.3], # 留出文本空间
                zeroline=True, zerolinecolor="#8c8c8c"
            ),
            xaxis2=dict(
                title="环比偏差(5%)",
                tickformat=".0%",
                range=[0, 0.5], # 环比范围固定，方便观察斜率
                showgrid=False,
                zeroline=True
            ),
            font=dict(family="Microsoft YaHei",size=10),
        )
        

        # 样式细节：隐藏右图Y轴刻度，统一字体
        fig.update_yaxes(showgrid=False, row=1, col=1)
        fig.update_yaxes(showticklabels=False, row=1, col=2)
        fig.update_yaxes(tickfont=dict(size=13), row=1, col=1)

        st.plotly_chart(fig, width='stretch')
            
    df_cat = df_filtered[df_filtered["周数"] == select_week].groupby("品类")[stage_cols].mean().reset_index()
    df_cat = df_filtered.groupby("品类").agg({
        "单周预测偏差率": "mean",
        "环比预测偏差率": "mean",
        "主料mrpsku": "nunique" 
    }).reset_index()
    df_cat.columns = ["品类", "单周预测偏差率", "环比预测偏差率", "SKU个数"]
    # 策略：只给 [偏差率 > 30%] 且 [SKU数量在前 10%] 的重点品类打标签
    sku_threshold = df_cat["SKU个数"].quantile(0.9)
    def get_label(row):
        if (abs(row["单周预测偏差率"]) > 0.27 or abs(row["环比预测偏差率"]) > 0.3) and row["SKU个数"] >= sku_threshold:
            return row["品类"]
        # if abs(row["单周预测偏差率"]) > 1.0:
        #     return row["品类"]
        return ""

    df_cat["显示标签"] = df_cat.apply(get_label, axis=1)
    def get_color(row):
        if row["单周预测偏差率"] > 0.27:
            return "预测过高"
        if row["单周预测偏差率"] < -0.27:
            return "预测过低"
        return "正常"
    
    df_cat["偏差情况"] = df_cat.apply(get_color, axis=1) 
    detail_cat,detail_sku = st.columns([2,3])
    df_cat['单周预测偏差率'] = round(df_cat['单周预测偏差率']*100, 1)
    df_cat['环比预测偏差率'] = round(df_cat['环比预测偏差率']*100, 1)
    with detail_cat:
        st.markdown(
            f"""
            #### 预测情况下钻-<span style='color: #ff4b4b;'>{st.session_state.yuce_filter_market}</span> 品类明细表
            """, 
            unsafe_allow_html=True
        )

        def color_deviation(val,target):
            color = '#cf1322' if abs(val) > target else '#389e0d'
            return f'color: {color}; font-weight: bold' if abs(val) > target else f'color: {color}'
        df_cat["周数"] = select_week
        styled_df = df_cat[['周数','品类', '单周预测偏差率', '环比预测偏差率', 'SKU个数', '偏差情况']].sort_values("单周预测偏差率", ascending=False).style.map(
            color_deviation, subset=['单周预测偏差率'],target=35
        )
        styled_df = styled_df.map(
            color_deviation, subset=['环比预测偏差率'],target=5
        )

        
        st.dataframe(
            styled_df,
            column_config={
                "单周预测偏差率": st.column_config.NumberColumn(
                    "单周预测偏差率(目标值:35%)",
                    help="预测值与实际值的偏移程度",
                    format="%.1f%%",
                    width=80,
                    alignment="center"
                ),
                "环比预测偏差率": st.column_config.NumberColumn(
                    "环比预测偏差率(目标值:5%)",
                    help="本周对比上周偏差的变化",
                    format="%.1f%%",
                    width=80,
                    alignment="center"
                ),
                "SKU个数": st.column_config.NumberColumn(
                    "SKU个数",
                    format="%d",
                    width=50,
                    alignment="center"
                ),
                "偏差情况": st.column_config.TextColumn("偏差情况", width=80),
        },
         width="stretch", height=300, hide_index=True)
    with detail_sku:
        st.markdown(
            f"""
            #### 预测情况下钻-<span style='color: #ff4b4b;'>{st.session_state.yuce_filter_market}_{st.session_state.yuce_filter_category}</span> MRPSKU明细表
            """, 
            unsafe_allow_html=True
        )
        df_detail = df_filtered[df_filtered["周数"] == select_week][['周数','子市场', "channel_name",'品类', "主料mrpsku","状态", "当周实际值", "当周预测值" ,"单周预测偏差率", "环比预测偏差率"]].copy()
        df_detail=df_detail.sort_values("单周预测偏差率", ascending=False)
        df_detail['单周预测偏差率'] = round(df_detail['单周预测偏差率']*100, 1)
        df_detail['环比预测偏差率'] = round(df_detail['环比预测偏差率']*100, 1)
        def color_deviation(val,target):
            color = '#cf1322' if abs(val) > target else '#389e0d'
            return f'color: {color}; font-weight: bold' if abs(val) > target else f'color: {color}'
        styled_df = df_detail[['周数','子市场', "channel_name",'品类', "主料mrpsku","状态", "当周实际值", "当周预测值" ,"单周预测偏差率", "环比预测偏差率"]].sort_values("单周预测偏差率", ascending=False).style.map(
            color_deviation, subset=['单周预测偏差率'],target=35
        )
        styled_df = styled_df.map(
            color_deviation, subset=['环比预测偏差率'],target=5
        )
        st.dataframe(
            styled_df,
            column_config={
                "周数": st.column_config.TextColumn("周数", width=20),
                "子市场": st.column_config.TextColumn("子市场", width=20),
                "channel_name": st.column_config.TextColumn("渠道", width=20),
                "品类": st.column_config.TextColumn("品类", width=20),
                "主料mrpsku": st.column_config.TextColumn("MRPSKU", width=30),
                "状态": st.column_config.TextColumn("状态", width=20),
                "当周实际值": st.column_config.NumberColumn("实际周销", format="%d", width=20, alignment="center"),
                "当周预测值": st.column_config.NumberColumn("预测周销", format="%d", width=20, alignment="center"),
                # "上周预测值": st.column_config.NumberColumn("上周预测周销", format="%d", width=20, alignment="center"),
                "单周预测偏差率": st.column_config.NumberColumn(
                    "单周预测偏差率(目标值:35%)",
                    format="%.1f%%", width=50, alignment="center" 
                ),
                "环比预测偏差率": st.column_config.NumberColumn(
                    "环比预测偏差率(目标值:5%)",
                    format="%.1f%%", width=50, alignment="center" 
                )
            },
            hide_index=True,
            width="stretch",
            height=300
        )
# 干预指标区域
def ganyuSales_rate_area(df_ganyu,df_ganyu_bi, curr_filters):
    df = apply_filters(df_ganyu, curr_filters)
    if df is None or df.empty:
        st.warning("无数据")
        return
    df=df[df['是否有干预'] == '是']
    # st.subheader("📊 预测准确度监控看板")
    market_filter, category_filter, sku_filter,outStock_filter,week_filter = st.columns(5)
    with market_filter:
        market_list = ["全部市场"] + sorted(df["子市场"].unique().tolist())
        selected_market = st.selectbox("选择要查看的子市场", market_list,key="selectbox_market_ganyu")
        st.session_state.ganyu_filter_market=selected_market
    with category_filter:
        category_list = ["全部品类"] + sorted(df["品类"].unique().tolist())
        selected_category = st.selectbox("选择要查看的品类", category_list,key="selectbox_category_ganyu")
        st.session_state.ganyu_filter_category=selected_category
    with sku_filter:
        SKU_list = ["全部SKU"] + sorted(df["主料mrpsku"].unique().tolist())
        selected_sku = st.selectbox("选择要查看的SKU", SKU_list,key="selectbox_sku_ganyu")
        st.session_state.ganyu_filter_sku=selected_sku
    with outStock_filter:
        outStock_status_list = ["全部状态","断货","非断货"]
        select_stockStatus = st.selectbox("选择SKU的状态", outStock_status_list,key="selectbox_outStock_ganyu")
        st.session_state.ganyu_filter_outStockStatus=select_stockStatus
    with week_filter:
        select_week = st.date_input("选择要查看的周数", value=default_monday,key="date_input_week_ganyu")
        select_week = select_week.strftime("%Yw%V")
        st.session_state.ganyu_filter_week=select_week

    if selected_market == "全部市场":
        df_filtered = df
    elif selected_market != "全部市场" and selected_category == "全部品类":
        df_filtered = df[(df["子市场"] == selected_market)]
    elif selected_market != "全部市场" and selected_category != "全部品类":
        df_filtered = df[(df["子市场"] == selected_market) & (df["品类"] == selected_category)]
    elif selected_market != "全部市场" and selected_category != "全部品类" and selected_sku != "全部SKU":
        df_filtered = df[(df["子市场"] == selected_market) & (df["品类"] == selected_category) & (df["主料mrpsku"] == selected_sku)]
    
    if select_stockStatus == "全部状态":
        df_filtered = df_filtered
    elif select_stockStatus == "断货":
        df_filtered = df_filtered[df_filtered["状态"] == "断货"]
    elif select_stockStatus == "非断货":
        df_filtered = df_filtered[df_filtered["状态"] != "断货"]
    
    avg_bias = df_filtered["单周干预偏差率"].abs().mean()
    avg_bias_ratio = df_filtered["环比干预偏差率"].abs().mean()
    m_col1, m_col2, m_col3,m_col4 = st.columns([1,1,1,1])
    ganyu_intervention_rate = df_ganyu_bi["有干预样本数"].sum() / df_ganyu_bi["总样本数"].sum()
    m_col1.metric("干预SKU个数占比", f"{ganyu_intervention_rate:.0%}", delta="目标:15%")
    m_col2.metric("有效干预SKU个数占比", f"{(df_ganyu_bi['有干预样本数'].sum() - df_ganyu_bi['不应干预样本数'].sum()) / df_ganyu_bi['有干预样本数'].sum():.0%}",delta="有效干预规则:±30%")
    m_col3.metric("干预偏差率", f"{avg_bias:.0%}" , delta="目标:30%")
    m_col4.metric("干预偏差率(环比)", f"{avg_bias_ratio:.0%}", delta="目标:5%")

    col_left, col_right = st.columns([2,3])
    # 市场维度聚合
    stage_cols = ["单周干预偏差率", "环比干预偏差率"]
    df_market = df_filtered.groupby("子市场")[stage_cols].mean().reset_index()
    df_plot = df_market.melt(id_vars="子市场", value_vars=stage_cols, 
                            var_name="指标类型", value_name="偏差率")
    if st.session_state.ganyu_filter_market != "全部市场":
        df_ganyu_bi=df_ganyu_bi[df_ganyu_bi["子市场"] == st.session_state.ganyu_filter_market]
    df_inter = df_ganyu_bi.sort_values("有干预样本数", ascending=False)
    df_inter["有效干预数"] = df_inter["有干预样本数"] - df_inter["不应干预样本数"]
    with col_left:
        st.markdown("#### 历史干预SKU个数占比趋势")
        df_inter_week = df_inter.groupby("周数").agg({
            "总样本数": "sum",
            "有干预样本数": "sum",
        }).reset_index()
        df_inter_week['干预SKU个数占比'] = df_inter_week["有干预样本数"] / df_inter_week["总样本数"]
        fig_历史干预SKU曲线=go.Figure()
        fig_历史干预SKU曲线.add_trace(
            go.Scatter(
                x=df_inter_week["周数"],
                y=df_inter_week["干预SKU个数占比"],
                mode='lines+markers+text',
                line=dict(color='#1f77b4', width=2),
                marker=dict(
                    color="#1f77b4", 
                    size=20, 
                    line=dict(width=5, color='white')
                ),
                text=df_inter_week["干预SKU个数占比"].apply(lambda x: f"{x:.0%}"),
                textposition="top left",
                name="干预SKU个数占比"
            )
        )
        fig_历史干预SKU曲线.add_hline(
            y=0.15,
            line_dash="dash",
            line_color="#1f77b4",  
            annotation_text="预测偏差: 30%",
            annotation_position="bottom right",
            opacity=0.7
        )
        fig_历史干预SKU曲线.update_layout(
            # 左侧 Y 轴配置
            yaxis=dict(
                # title="单周干预偏差率",
                range=[0, 1.0],  
                tickformat=".0%", 
                side="left"
            ),
            margin=dict(r=80),
            legend=dict(
                orientation="h",          # 水平排列
                yanchor="bottom",         # 垂直方向以底部为锚点
                y=1.02,                   # 放在绘图区顶部稍微偏上的位置
                xanchor="center",         # 水平方向以中心为锚点
                x=0.5,                    # 放在水平方向 50% 的位置（居中）
                bgcolor="rgba(255,255,255,0.8)" # 可选：设置半透明背景，防止遮挡标题或边框
            ),
            font=dict(family="Microsoft YaHei",size=12)
        )

        st.plotly_chart(fig_历史干预SKU曲线,width="stretch",height=500)


    with col_right:
        st.markdown("#### 干预SKU个数占比")
        # if st.session_state.ganyu_filter_market != "全部市场":
        #     df_ganyu_bi=df_ganyu_bi[df_ganyu_bi["子市场"] == st.session_state.ganyu_filter_market]
        # df_inter = df_ganyu_bi.sort_values("有干预样本数", ascending=False)
        # df_inter["有效干预数"] = df_inter["有干预样本数"] - df_inter["不应干预样本数"]
        df_inter=df_inter[df_inter["周数"] == select_week]
        fig = make_subplots(
            rows=1, cols=2,
            column_widths=[0.35, 0.65], # 左侧占比稍小
            horizontal_spacing=0.12,
            subplot_titles=("市场干预SKU占比分布", "有效干预VS无效干预"),
            specs=[[{"type": "pie"}, {"type": "xy"}]]
        )
        labels = df_inter["子市场"].tolist()
        v_total = df_inter["有干预样本数"].tolist()
        v_effective = df_inter["有效干预数"].tolist()
        # --- 3. 左侧：精美的圆环图 ---
        business_colors = ['#1f77b4', '#4e79a7', '#76b7b2', '#54a24b', '#e15759', '#f28e2b', '#bab0ac', '#86bcb6', '#dbedff']
        fig.add_trace(go.Pie(
            labels=labels,
            values=v_total,
            hole=0.6,
            textinfo='label+percent', 
            textposition='outside',
            automargin=True,
            showlegend=False, 
            marker=dict(
                colors=business_colors, 
                line=dict(color='#FFFFFF', width=2),
            ),
            # marker=dict(line=dict(color='#FFFFFF', width=2)),
            name="市场占比",
            opacity=0.7,
        ), row=1, col=1)

        # --- 4. 右侧：簇状柱状图 ---
        # 柱子1：有干预样本数（背景色调）
        fig.add_trace(go.Bar(
            x=labels,
            y=v_total,
            name='有干预样本数',
            marker_color='rgba(55, 128, 191, 0.6)', # 浅蓝色
            offsetgroup=1,
            text=v_total,
            textposition='outside',
        ), row=1, col=2)

        # 柱子2：有效干预数（深色突出）
        fig.add_trace(go.Bar(
            x=labels,
            y=v_effective,
            name='有效干预数',
            marker_color='rgba(50, 171, 96, 1.0)', # 翠绿色
            offsetgroup=2,
            text=v_effective,
            textposition='outside',
        ), row=1, col=2)

        top_y1 = max(v_total) + 200
        for i, row in df_inter.iterrows():
            start_y = max(row["有干预样本数"], row["有效干预数"])  
            fig.add_shape(
                type="line",
                x0=row["子市场"], y0=start_y + 10,
                x1=row["子市场"], y1=top_y1 - 20,
                line=dict(color="gray", width=1, dash="dot"),
            )
            
            fig.add_trace(go.Scatter(
                x=[row["子市场"]],
                y=[top_y1],
                mode="markers+text",
                marker=dict(color="white", size=10, line=dict(color="#ff4d4f", width=2)),
                text=[f"{row['有干预样本数'] - row['有效干预数']:.0f}"],
                textposition="top center",
                textfont=dict(color="#ff4d4f", size=14, family="Arial Black"),
                showlegend=i == 0,
                hoverinfo='skip',
                name='无效干预样本数'
            ))
        for i, annotation in enumerate(fig['layout']['annotations']):
            annotation['font'] = dict(family="Microsoft YaHei", size=20, color="black")
        
        max_y = max(v_total) if v_total else 1000
        fig.update_yaxes(range=[0, max_y * 1.3], row=1, col=2)
        # --- 6. 整体布局美化 ---
        fig.update_layout(
            height=500,
            width=1100,
            template="plotly_white",
            barmode='group', # 簇状排列
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.2,
                xanchor="center",
                x=0.5
            ),
            margin=dict(t=40, b=100, l=50, r=50),
            font=dict(family="Microsoft YaHei", size=12)
        )

        # 优化坐标轴
        fig.update_yaxes(title_text="样本数量", row=1, col=2)
        fig.update_xaxes(tickangle=0, row=1, col=2)

        st.plotly_chart(fig, width='stretch')

    df_m = df_filtered[df_filtered["周数"] == select_week].groupby("子市场").agg({
        "单周干预偏差率": "mean",
        "环比干预偏差率": "mean",
        "主料mrpsku": "count"
    }).reset_index()
    df_m = df_m.rename(columns={"主料mrpsku": "SKU数量"})
    df_m = df_m.sort_values(by="SKU数量", ascending=True)
    col_bias_left, col_bias_right = st.columns([2,3])
    df_filtered_week = df_filtered.groupby("周数").agg({
        "单周干预偏差率": "mean",
        "环比干预偏差率": "mean"
    }).reset_index()

    with col_bias_left:
        st.markdown("#### 历史干预偏差趋势")
        fig_历史干预曲线=go.Figure()
        fig_历史干预曲线.add_trace(
            go.Scatter(
                x=df_filtered_week["周数"],
                y=df_filtered_week["单周干预偏差率"],
                mode='lines+markers+text',
                line=dict(color='#1f77b4', width=2),
                marker=dict(
                    color="#1f77b4", 
                    size=20, 
                    line=dict(width=5, color='white')
                ),
                text=df_filtered_week["单周干预偏差率"].apply(lambda x: f"{x:.0%}"),
                textposition="top left",
                name="单周偏差"
            )
        )
        fig_历史干预曲线.add_trace(
            go.Scatter(
                x=df_filtered_week["周数"],
                y=df_filtered_week["环比干预偏差率"],
                mode='lines+markers+text',
                line=dict(color='#ff7f0e', width=2),
                marker=dict(
                    color="#ff7f0e", 
                    size=20, 
                    line=dict(width=5, color='white')
                ),
                text=df_filtered_week["环比干预偏差率"].apply(lambda x: f"{x:.0%}"),
                textposition="bottom left",
                name="环比偏差",
                yaxis="y2"
            )
        )
        fig_历史干预曲线.add_hline(
            y=0.30,
            line_dash="dash",
            line_color="#1f77b4",  
            annotation_text="预测偏差: 30%",
            annotation_position="bottom right",
            opacity=0.7
            
        )
        fig_历史干预曲线.add_shape(
            type="line",
            xref="paper", yref="y2",
            x0=0, y0=0.05,
            x1=1, y1=0.05,
            line=dict(color="#ff7f0e", width=2, dash="dash"),
            opacity=0.7
        )
        fig_历史干预曲线.update_layout(
            # 左侧 Y 轴配置
            yaxis=dict(
                title="单周干预偏差率",
                range=[0, 1.5],  
                tickformat=".0%", 
                side="left"
            ),
            # 右侧 Y 轴配置
            yaxis2=dict(
                title="环比干预偏差率",
                range=[0, 0.2],  
                tickformat=".0%", 
                overlaying="y",   
                side="right"
            ),
            # 可选：优化整体布局，防止右侧 Y 轴标题被截断
            margin=dict(r=80),
            legend=dict(
                orientation="h",          # 水平排列
                yanchor="bottom",         # 垂直方向以底部为锚点
                y=1.02,                   # 放在绘图区顶部稍微偏上的位置
                xanchor="center",         # 水平方向以中心为锚点
                x=0.5,                    # 放在水平方向 50% 的位置（居中）
                bgcolor="rgba(255,255,255,0.8)" # 可选：设置半透明背景，防止遮挡标题或边框
            ),
            font=dict(family="Microsoft YaHei",size=12)
        )

        st.plotly_chart(fig_历史干预曲线,width="stretch",height=500)



    with col_bias_right:
        TARGET_MAX = 0.30
        st.markdown("#### 各市场干预偏差全貌 (目标区间: 30%)")

        # --- 2. 创建布局 ---
        fig = make_subplots(
            rows=1, cols=2, 
            shared_yaxes=True, 
            horizontal_spacing=0.03, 
            column_widths=[0.7, 0.3]
        )

        # --- 4. 左侧：单周干预偏差率 (条形图) ---
        # 颜色逻辑：出界的标红，在区间内的用青色
        colors_bar = [
            "#ff9d4f" if (x > TARGET_MAX) else "#5da9c4" 
            for x in df_m["单周干预偏差率"]
        ]

        fig.add_trace(
            go.Bar(
                y=df_m["子市场"],
                x=df_m["单周干预偏差率"],
                orientation='h',
                marker_color=colors_bar,
                text=df_m["单周干预偏差率"].apply(lambda x: f"{x:.1%}"),
                textposition='outside',
                cliponaxis=False,
                name="单周干预偏差"
            ),
            row=1, col=1
        )

        for line_x in [TARGET_MAX]:
            fig.add_vline(
                x=line_x, 
                line_dash="dash", 
                line_color="red", 
                line_width=2,
                row=1, col=1,
                # 不透明度
                opacity=0.8
            )


        # --- 5. 右侧：环比干预偏差率 (折线图) ---
        colors_line = ["#5da9c4" if abs(x) <= 0.05 else "#f5222d" for x in df_m["环比干预偏差率"]]

        fig.add_trace(
            go.Scatter(
                y=df_m["子市场"],
                x=df_m["环比干预偏差率"],
                mode='lines+markers+text',
                line=dict(color='#bfbfbf', width=3),
                marker=dict(
                    color=colors_line, 
                    size=15, 
                    line=dict(width=5, color='white')
                ),
                text=df_m["环比干预偏差率"].apply(lambda x: f"{x:.0%}"),
                textposition="middle right",
                name="环比趋势"
            ),
            row=1, col=2
        )
        for line_x in [0.05]:
            fig.add_vline(
                x=line_x, 
                line_dash="dash", 
                line_color="red", 
                line_width=2,
                row=1, col=2,
                # 不透明度
                opacity=0.8
            )

        max_val = max(df_m["单周干预偏差率"].max(), 0.5)
        min_val = min(df_m["单周干预偏差率"].min(), -0.5)

        fig.update_layout(
            template="simple_white",
            showlegend=False,
            height=len(df_m) * 45 + 120,
            margin=dict(l=10, r=60, t=20, b=40),
            xaxis=dict(
                title="干预偏差率",
                tickformat=".0%",
                range=[0, max_val+0.3], # 留出文本空间
                zeroline=True, zerolinecolor="#8c8c8c"
            ),
            xaxis2=dict(
                title="环比偏差率",
                tickformat=".0%",
                range=[0, 0.5], # 环比范围固定，方便观察斜率
                showgrid=False,
                zeroline=True,
                # 修改X轴区间范围
                # tickvals=[-1.1, -0.05, 0, 0.05, 1.1]
            ),
            font=dict(family="Microsoft YaHei"),
        )

        # 样式细节：隐藏右图Y轴刻度，统一字体
        fig.update_yaxes(showgrid=False, row=1, col=1)
        fig.update_yaxes(showticklabels=False, row=1, col=2)
        fig.update_yaxes(tickfont=dict(size=13), row=1, col=1)

        st.plotly_chart(fig, width='stretch')
    
    detail_cat,detail_sku_ganyu = st.columns([2,3])
    df_cat = df_filtered[df_filtered["周数"] == select_week].groupby("品类").agg({
        "单周干预偏差率": "mean",
        "环比干预偏差率": "mean",
        "主料mrpsku": "nunique" 
    }).reset_index()
    df_cat.columns = ["品类", "单周干预偏差率", "环比干预偏差率", "SKU个数"]
    # 策略：只给 [偏差率 > 30%] 且 [SKU数量在前 10%] 的重点品类打标签
    sku_threshold = df_cat["SKU个数"].quantile(0.9)
    def get_label(row):
        if (abs(row["单周干预偏差率"]) > 0.3 or abs(row["环比干预偏差率"]) > 0.05) and row["SKU个数"] >= sku_threshold:
            return row["品类"]
        return ""

    df_cat["显示标签"] = df_cat.apply(get_label, axis=1)
    def get_color(row):
        if row["单周干预偏差率"] > 0.3: 
            return "干预过高"
        if row["单周干预偏差率"] < -0.3:
            return "干预过低"
        return "正常"
    
    df_cat["偏差情况"] = df_cat.apply(get_color, axis=1)
    df_cat['单周干预偏差率'] = round(df_cat['单周干预偏差率']*100, 1)
    df_cat['环比干预偏差率'] = round(df_cat['环比干预偏差率']*100, 1)
    with detail_cat:
        st.markdown(
            f"""
            #### 干预情况下钻-<span style='color: #ff4b4b;'>{st.session_state.ganyu_filter_market}</span> 品类明细表
            """, 
            unsafe_allow_html=True
        )
        def color_deviation(val,target):
            color = '#cf1322' if abs(val) > target else '#389e0d'
            return f'color: {color}; font-weight: bold' if abs(val) > target else f'color: {color}'
        df_cat['周数']=select_week
        styled_df = df_cat[['周数','品类', '单周干预偏差率', '环比干预偏差率', 'SKU个数', '偏差情况']].sort_values("单周干预偏差率", ascending=False).style.map(
            color_deviation, subset=['单周干预偏差率'], target=30
        )
        styled_df = styled_df.map(
            color_deviation, subset=['环比干预偏差率'], target=5
        )


        st.dataframe(styled_df,
        column_config={
            "单周干预偏差率": st.column_config.NumberColumn(
                "单周干预偏差率(目标值:30%)",
                help="干预值与实际值的偏移程度",
                format="%.1f%%",
                width=80,
                alignment="center"
            ),
            "环比干预偏差率": st.column_config.NumberColumn(
                "环比干预偏差率(目标值:5%)",
                help="本周对比上周干预偏差的变化",
                format="%.1f%%",
                width=80,
                alignment="center"
            ),
            "SKU个数": st.column_config.NumberColumn("SKU个数", width=50, alignment="center"),
            "偏差情况": st.column_config.TextColumn("偏差情况", width=80),
            "品类": st.column_config.TextColumn("品类", width=50),
        },
         width="stretch", height=400, hide_index=True)
    
    with detail_sku_ganyu:
        st.markdown(
            f"""
            #### 干预情况下钻-<span style='color: #ff4b4b;'>{st.session_state.ganyu_filter_market}_{st.session_state.ganyu_filter_category}</span> MRPSKU明细表
            """, 
            unsafe_allow_html=True
        )
        inter_list = ["全部"] + sorted(df["是否应该干预"].unique().tolist())
        selected_inter = st.selectbox("是否应该干预", inter_list,key="selectbox_inter")
        if selected_inter != "全部":
            df_filtered = df_filtered[df_filtered["是否应该干预"] == selected_inter]
        df_detail = df_filtered[df_filtered["周数"] == select_week][['周数','子市场', "channel_name",'品类', "主料mrpsku","状态", "当周实际值", "当周干预值", "单周干预偏差率", "环比干预偏差率", "是否应该干预"]].copy()
        df_detail['单周干预偏差率'] = round(df_detail['单周干预偏差率']*100, 1)
        df_detail['环比干预偏差率'] = round(df_detail['环比干预偏差率']*100, 1)
        
        def color_deviation(val,target):
            color = '#cf1322' if abs(val) > target else '#389e0d'
            return f'color: {color}; font-weight: bold' if abs(val) > target else f'color: {color}'
        styled_df = df_detail[['周数','子市场', "channel_name",'品类', "主料mrpsku","状态", "当周实际值", "当周干预值", "单周干预偏差率", "环比干预偏差率", "是否应该干预"]].sort_values("单周干预偏差率", ascending=False).style.map(
            color_deviation, subset=['单周干预偏差率'],target=30
        )
        styled_df = styled_df.map(
            color_deviation, subset=['环比干预偏差率'], target=5
        )
        st.dataframe(
            styled_df,
            column_config={
                "周数": st.column_config.TextColumn("周数",width=50),
                "子市场": st.column_config.TextColumn("子市场",width=90),
                "channel_name": st.column_config.TextColumn("渠道",width=90),
                "品类": st.column_config.TextColumn("品类",width=80),
                "主料mrpsku": st.column_config.TextColumn("MRPSKU", width="medium"),
                "状态": st.column_config.TextColumn("状态", width=80),
                "当周实际值": st.column_config.NumberColumn("实际周销", format="%.0f", width=80, alignment="center"),
                "当周干预值": st.column_config.NumberColumn("干预周销", format="%.0f", width=80, alignment="center"),
                "单周干预偏差率": st.column_config.NumberColumn(
                    "单周干预偏差率(目标值:30%)",
                    help="干预值与实际值的偏移程度",
                    format="%.1f%%",
                    width=80,
                    alignment="center"
                ),
                "环比干预偏差率": st.column_config.NumberColumn(
                    "环比干预偏差率(目标值:5%)",
                    help="本周对比上周干预偏差的变化",
                    format="%.1f%%",
                    width=80,
                    alignment="center"
                ),
                "是否应该干预": st.column_config.TextColumn("是否应该干预", width=80)
            },
            hide_index=True,
            width="stretch",
            height=300
        )


def delivery_stock_area(df_fahuo, curr_filters, filter_market,df_country_stock):
    df = apply_filters(df_fahuo, curr_filters)
    if df is None or df.empty: return
    
    stage_cols = ["计划达成率", "配货达成率", "排单达成率", "出库达成率"]
    df_avg = df.groupby("子市场").agg({
        "计划达成率": "mean",
        "配货达成率": "mean",
        "排单达成率": "mean",
        "出库达成率": "mean",
        "主料mrpsku": "count"
    }).reset_index()
    df_avg = df_avg.rename(columns={"主料mrpsku": "SKU数量"})
    df_avg = df_avg.sort_values(by="SKU数量", ascending=True)
    # 计算全市场平均值
    df_avg_allmarket = df[stage_cols].mean().reset_index()
    df_avg_allmarket[0] = df_avg_allmarket[0].astype(float)
    
    col1, col2 = st.columns([2.5, 2])
    thresholds = {
        "计划达成率": 0.9, 
        "配货达成率": 0.9,
        "排单达成率": 1.0,
        "出库达成率": 1.0
    }
    def style_threshold_clean(col):
        if col.name in thresholds:
            target = thresholds[col.name]
            styles = []
            for v in col:
                if v >= target:
                    styles.append('color: #2E7D32;') 
                else:
                    styles.append('background-color: #FFF5F5; color: #C62828; font-weight: bold;')
            return styles
        return [''] * len(col)

    styled_df = (
        df_avg.style
        .apply(style_threshold_clean)
        .set_properties(**{
            'text-align': 'center',
        })
        .format("{:.1%}", subset=list(thresholds.keys()))
    )
    with col1:
        st.markdown("#### 📦 全市场发货过程指标全貌")
        c1,c2,c3,c4 = st.columns(4)
        # 如果指标值小于目标就为红色，反之为绿色
        c1.metric("计划达成率", f"{df_avg_allmarket.iloc[0,1]:.1%}",delta="目标: 90%",delta_color="green")
        c2.metric("配货达成率", f"{df_avg_allmarket.iloc[1,1]:.1%}",delta="目标: 90%",delta_color="green")
        c3.metric("排单达成率", f"{df_avg_allmarket.iloc[2,1]:.1%}",delta="目标: 100%",delta_color="green")
        c4.metric("出库达成率", f"{df_avg_allmarket.iloc[3,1]:.1%}",delta="目标: 100%",delta_color="green")
        st.dataframe(
            styled_df, 
            width='stretch', 
            column_config={
                "子市场": st.column_config.TextColumn("子市场", width="small"),
                **{
                    col: st.column_config.TextColumn(
                        label=f"{col}\n(目标:{thresholds[col]:.0%})", 
                        help=f"目标: {thresholds[col]:.1%}",
                        alignment="center"
                    ) for col in thresholds.keys()
                },
                "SKU数量": st.column_config.TextColumn("SKU数量", width="small",alignment="center"),
            },
            height=480,
            hide_index=True,
        )

    with col2:
        st.markdown("#### 📦 发货指标下钻分析-子市场")
        market_list =  ["全部市场"] + sorted(df['子市场'].unique().tolist()) 
        selected_market = st.selectbox("🎯 选择下钻的子市场", market_list)
        if selected_market == "全部市场":
            df_filtered=df.copy()
        else:
            df_filtered = df[df['子市场'] == selected_market]
        st.session_state.filter_market=selected_market
        stage_cols = ['计划发货量', '配货数量', '排单数量', '实际出库量']
        df_sum = df_filtered[stage_cols].sum().reset_index()
        df_sum.columns = ['执行环节', '数量']
        st.markdown(
            f"""
            ##### 子市场 - <span style='color: #ff4b4b;'>{st.session_state.filter_market}</span> 发货全链路执行情况
            """, 
            unsafe_allow_html=True
        )
        fig = px.bar(
            df_sum,
            x="数量",
            y="执行环节",
            orientation='h',
            text='数量',
            color='执行环节',
            category_orders={"执行环节": ['计划发货量', '配货数量', '排单数量',  '实际出库量']},
            # title=f"子市场 [{selected_market}] 发货全链路执行情况",
            color_discrete_sequence=px.colors.qualitative.Pastel
        )

        # 5. 美化调整
        fig.update_traces(
            textposition='inside', 
            # 【修改】使用 HTML 标签 <b> 让数值加粗，并设置 textfont 变大
            # texttemplate='<b>%{text}</b>', 
            texttemplate='<b>%{text:,}</b>', 
            textfont=dict(
                size=16,             # 设置条形图内部字体大小
                family="Microsoft YaHei",
                color="black",      # 根据背景调整颜色，通常内部用白色或黑色
            ),
            hovertemplate="环节: %{y}<br>数量: %{x}<extra></extra>"
        )

        fig.update_layout(
            xaxis_title="<b>数量</b>",  # 轴标题加粗
            yaxis_title="",
            title_text="",
            showlegend=False,
            height=450,
            plot_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=150),
            # 【新增】全局字体设置
            font=dict(family="Microsoft YaHei"),
            # 【新增】设置标题加粗
            # title_font=dict(size=20, family="Microsoft YaHei"),
        )

        # 【修改】坐标轴刻度字体加粗
        fig.update_xaxes(
            showgrid=True, 
            gridcolor='LightGrey',
            tickfont=dict(size=14, family="Microsoft YaHei", color="black"), # 刻度加粗
            title_font=dict(size=16, family="Microsoft YaHei")              # 轴标题加粗
        )
        
        fig.update_yaxes(
            tickfont=dict(size=15, family="Microsoft YaHei", color="black"), # 纵轴环节名称加粗变大
        )

        st.plotly_chart(fig, width='stretch')
    
    if curr_filters is None:
        if filter_market == "全部市场":
            df=df_fahuo.copy()
        else:
            df=df_fahuo[df_fahuo['子市场'] == filter_market].copy()
    else:
        df = apply_filters(df_fahuo, curr_filters)

    # 1. 指标与排序配置
    col_ctrl1, col_ctrl2 = st.columns([2.5,3.5])
    with col_ctrl1:
        # st.markdown("#### 📊 发货指标下钻分析-品类")
        st.markdown(
            f"""
            #### 📊 发货指标下钻分析 - <span style='color: #ff4b4b;'>{st.session_state.filter_market}</span> 品类
            """, 
            unsafe_allow_html=True
        )
        metric = st.selectbox(
            "🎯 选择分析指标", 
            ["计划达成率", "配货达成率", "排单达成率", "出库达成率"]
        )
        
        # 1. 数据处理
        if st.session_state.filter_market=="全部市场":
            df_category=df.copy()
        else:
            df_category = df[df['子市场'] == st.session_state.filter_market]
        df_cat_rate = df_category.groupby("品类")[metric].mean().reset_index()

        # 排序：按达成率降序排列，最好的在上面
        df_category_tab = df_cat_rate.sort_values(metric, ascending=True).reset_index(drop=True)
        
        thresholds = {
            "计划达成率": 0.9, "配货达成率": 0.9, "排单达成率": 1.0, "出库达成率": 1.0
        }
        target_val = thresholds[metric]
        
        # 2. 计算配色与差值
        df_category_tab['是否达标'] = df_category_tab[metric] >= target_val
        df_category_tab['主颜色'] = df_category_tab['是否达标'].apply(lambda x: '#4CAABF' if x else '#FF6B6B')
        df_category_tab['背景色'] = '#EAEAEA'

        # 【优化】动态高度计算：每行40px + 顶部40px，取消max(500)防止品类少时留白过多
        chart_real_height = (len(df_category_tab) * 40) + 40

        # 3. 创建子图
        fig = make_subplots(
            rows=1, cols=2, 
            shared_yaxes=True, 
            column_widths=[0.65, 0.35], # 稍微拓宽右侧空间显示差值
            horizontal_spacing=0.01
        )

        # 4. 添加目标背景条
        fig.add_trace(go.Bar(
            y=df_category_tab['品类'], 
            x=[target_val] * len(df_category_tab),
            orientation='h',
            marker_color=df_category_tab['背景色'],
            name='目标要求',
            width=0.7,
            hoverinfo='skip'
        ), row=1, col=1)

        # 5. 添加实际达成条
        fig.add_trace(go.Bar(
            y=df_category_tab['品类'], 
            x=df_category_tab[metric],
            orientation='h',
            marker_color=df_category_tab['主颜色'],
            name='实际达成',
            width=0.4,
            text=[f"<b>{x:.1%}</b>" for x in df_category_tab[metric]],
            textposition='inside',
            textfont=dict(size=12, color='white'),
        ), row=1, col=1)

        # 6. 【关键修改】处理右侧文本：显示实际值及与目标的差值
        status_text = []
        for v, d in zip(df_category_tab[metric], df_category_tab['是否达标']):
            diff = v - target_val # 计算差值
            color = "#4CAABF" if d else "#FF6B6B"
            icon = "✔" if d else "✘"
            # 格式化差值：+1.2% 或 -3.5%
            diff_str = f"{diff:+.1%}"
            status_text.append(f"<b>{v:.1%}</b> <span style='font-size:12px; color:{color}'>({diff_str}) {icon}</span>")

        fig.add_trace(go.Scatter(
            y=df_category_tab['品类'], 
            x=[0.01] * len(df_category_tab), 
            mode='markers+text',
            marker=dict(symbol='circle-open', size=18, color=df_category_tab['主颜色'], line_width=2),
            text=status_text,
            textposition='middle right',
            textfont=dict(size=13, color='#333333', family="Microsoft YaHei"),
            showlegend=False
        ), row=1, col=2)

        # 7. 【优化】布局减少留白
        fig.update_layout(
            title=None,  # 彻底移除标题
            showlegend=False,
            barmode='overlay',
            plot_bgcolor='white',
            height=chart_real_height,
            margin=dict(l=80, r=0, t=0, b=0), 
            font=dict(family="Microsoft YaHei"),
            xaxis=dict(visible=False, range=[0, max(df_category_tab[metric].max(), target_val) * 1.1]),
            xaxis2=dict(visible=False),
            yaxis=dict(
                autorange="min reversed",  # 保持升序排列
                tickmode="linear",
                showgrid=False,
                tickfont=dict(size=12, weight='bold', color='#333333')
            ),
            yaxis2=dict(domain=[0, 1]), 
            autosize=False,
        )
        # st.markdown("""
        #     <style>
        #     div[data-testid="stVerticalBlockBorderWrapper"]:has(div.st-key-my-plot-container) .stPlotlyChart {
        #         margin-top: -270px !important;      
        #         margin-bottom: 0 !important;
        #     }
        #     </style>
        # """, unsafe_allow_html=True)
        
        with st.container(height=min(500, chart_real_height), key="my-plot-container"):
            st.plotly_chart(fig, width='stretch', config={'displayModeBar': False})
    
    with col_ctrl2:
        st.markdown(
            f"""
            #### 🔍发货指标下钻分析 - <span style='color: #ff4b4b;'>{st.session_state.filter_market}</span> MRPSKU
            """, 
            unsafe_allow_html=True
        )

        # 1. 获取该市场下的所有品类，供用户选择
        if st.session_state.filter_market=="全部市场":
            df_market=df.copy()
        else:
            df_market = df[df['子市场'] == st.session_state.filter_market]
        # df_market = df.copy()
        category_list = sorted(df_market['品类'].unique().tolist())
        
        col1, col2 = st.columns([1, 3])
        with col1:
            selected_cat = st.selectbox("🎯 选定下钻的品类", ["全部品类"] + category_list)

        st.session_state.filter_cat=selected_cat
        # 2. 过滤数据
        if selected_cat == "全部品类":
            df_sku = df_market.copy()
        else:
            df_sku = df_market[df_market['品类'] == selected_cat].copy()

        # 3. 计算关键指标（可选：为了让用户知道在这个品类下有多少坑）
        total_skus = len(df_sku)
        problem_skus = len(df_sku[df_sku['实际出库量'] < df_sku['计划发货量']])
        
        with col2:
            # 用小组件显示概况
            m1, m2, m3 = st.columns(3)
            m1.metric("SKU 总数", total_skus)
            m2.metric("异常 SKU 数", problem_skus)
            m3.write("💡 *异常定义：实际出库量 < 计划发货量*")
        display_cols = [
            "子市场","主料mrpsku", "品类", "运输方式","计划发货量", "配货数量", "排单数量", "拣货量", "实际出库量"
        ]
        df_sku['缺口'] = df_sku['计划发货量'] - df_sku['实际出库量']
        df_sku_sorted = df_sku.sort_values("缺口", ascending=False)
        rate_cols = ["计划达成率", "配货达成率", "排单达成率", "拣货达成率", "出库达成率"]
        
        styled_sku = (
            df_sku_sorted[display_cols].style
        )

        # 6. 展示表格
        st.dataframe(
            styled_sku, 
            width='stretch', 
            height=500,
            column_config={
                "子市场": st.column_config.TextColumn("子市场", width=20),
                "品类": st.column_config.TextColumn("品类", width=20),
                "主料mrpsku": st.column_config.TextColumn("主料mrpsku", width=50),
                "运输方式": st.column_config.TextColumn("运输方式", width=20),
                "计划发货量": st.column_config.NumberColumn("计划发货量", format="%d", alignment="center"),
                "配货数量": st.column_config.NumberColumn("配货数量", format="%d", alignment="center"),
                "排单数量": st.column_config.NumberColumn("排单数量", format="%d", alignment="center"),
                "拣货量": st.column_config.NumberColumn("拣货量", format="%d", alignment="center"),
                "实际出库量": st.column_config.NumberColumn("实际出库量", format="%d", alignment="center"),
            },
            hide_index=True
        )
    st.subheader("国内在库情况明细表")
    if st.session_state.filter_market == "全部市场":
        df_filtered_country_stock = df_country_stock.copy()
    elif st.session_state.filter_market != "全部市场" and st.session_state.filter_cat == "全部品类":
        df_filtered_country_stock = df_country_stock[(df_country_stock["子市场"] == st.session_state.filter_market)]
    elif st.session_state.filter_market != "全部市场" and st.session_state.filter_cat != "全部品类":
        df_filtered_country_stock = df_country_stock[(df_country_stock["子市场"] == st.session_state.filter_market) & (df_country_stock["品类"] == st.session_state.filter_cat)]
    mrpsku_list = sorted(df_filtered_country_stock['主料mrpsku'].unique().tolist())
    selected_mrpsku = st.selectbox("🎯 选定下钻的SKU", ["全部SKU"] + mrpsku_list)
    if selected_mrpsku == "全部SKU":
        df_sku = df_filtered_country_stock.copy()
    else:
        df_sku = df_filtered_country_stock[df_filtered_country_stock['主料mrpsku'] == selected_mrpsku].copy()
    index_cols = [
        '子市场', '主料mrpsku', '货源地', '规格', '二级品类', '品类', 
        'PTSKU',  '是否有PO', '最近PO到货日期', '最近PO到货数量'
    ]
    # 日期列去掉时分秒
    df_sku['日期'] = df_sku['日期'].dt.strftime('%Y-%m-%d')
    df_sku['最近PO到货日期'] = df_sku['最近PO到货日期'].dt.strftime('%Y-%m-%d')
    value_cols = ['可用库存', '预占库存', '待上架库存', 'IQC数量']
    
    pivot_df = df_sku.pivot(
        index=index_cols, 
        columns='日期', 
        values=value_cols
    )
    # 3. 调整表头顺序（可选）
    # 默认 pivot 后，一级表头是“指标名称”，二级表头是“日期”
    # 如果你想要一级表头是“日期”，二级是“指标”，需要执行 swaplevel

    pivot_df = pivot_df.swaplevel(0, 1, axis=1).sort_index(axis=1)
    cols = pivot_df.columns
    new_order = ['可用库存', '预占库存','待上架库存', 'IQC数量']

    # 生成新的列顺序
    new_cols = []
    for date in pivot_df.columns.levels[0]:  # 遍历每个日期
        for metric in new_order:
            if (date, metric) in cols:
                new_cols.append((date, metric))

    # 重新排列列顺序
    pivot_df = pivot_df[new_cols]
    # 4. 在 Sreamlit 中展示
    # 整个表的字体为微软雅黑
    st.dataframe(
        pivot_df, 
        column_config={
            "可用库存": st.column_config.NumberColumn("可用库存", format="%d", alignment="center"),
            "预占库存": st.column_config.NumberColumn("预占库存", format="%d", alignment="center"),
            "待上架库存": st.column_config.NumberColumn("待上架库存", format="%d", alignment="center"),
            "IQC数量": st.column_config.NumberColumn("IQC数量", format="%d", alignment="center"),
            "最近PO到货数量": st.column_config.NumberColumn("最近PO到货数量", format="%d", alignment="center"),
        },
        width="stretch", 
        height=400
    )







# =========================================================
# 6、主体渲染
# =========================================================
if st.session_state.df_fahuo is not None:
    curr_filters = st.session_state.committed_filters
    # 库销比区域
    st.header("📉 海外库存周转", anchor="1")
    inventorySales_rate_area(st.session_state.df_stock_turnover, curr_filters)
    # plot_inventorySales_rate(st.session_state.df_kuxiao, curr_filters)
    st.divider()
    # 发货指标区域
    st.markdown("# 供")
    st.header("🚚 发货过程指标", anchor="2")
    delivery_stock_area(st.session_state.df_fahuo, curr_filters, st.session_state.filter_market, st.session_state.df_country_stock)
    st.divider()
    # 预测指标区域
    st.markdown("# 销")
    st.header("📈 预测指标", anchor="3")
    predictSales_rate_area(st.session_state.df_yuce, curr_filters)
    st.divider()
    # 干预指标区域
    st.header("🛠️ 干预指标", anchor="4")
    ganyuSales_rate_area(st.session_state.df_ganyu, st.session_state.df_ganyu_bi, curr_filters)
else:
    st.info("👋 请先在左侧侧边栏上传数据文件。")


