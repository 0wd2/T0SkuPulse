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
import warnings
import fastexcel

warnings.filterwarnings("ignore")
import io
import time
from xlsxwriter import Workbook

st.set_page_config(page_title="全层级指标监控", layout="wide", page_icon="📊")
if "df_海外周转" not in st.session_state: st.session_state.df_海外周转 = None
if "df_国内在库周转" not in st.session_state: st.session_state.df_国内在库周转 = None
if "df_断货率" not in st.session_state: st.session_state.df_断货率 = None
if "committed_filters" not in st.session_state:st.session_state.committed_filters = {"level": []}
if "filter_ver" not in st.session_state: st.session_state.filter_ver = 0

@st.cache_data(show_spinner="正在解析并格式化数据，请稍候...", ttl=3600)
def process_uploaded_files(uploaded_files):
    data_pool = {
        "df_海外周转": None, "df_国内在库周转": None,
        "df_断货率": None,
    }

    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        file_bytes = uploaded_file.getvalue()
        file_type = file_name.split('.')[-1].lower()

        try:
            if file_type == 'xlsx':
                # 使用 BytesIO 读取
                with io.BytesIO(file_bytes) as f:
                    # 获取所有 Sheet 名称
                    workbook = load_workbook(f, read_only=True)
                    sn = workbook.sheetnames

                    # 定义内部格式化工具
                    def format_df(df):
                        if df is not None and not df.empty:
                            if '日期' in df.columns:
                                df['日期'] = pd.to_datetime(df['日期'])
                            return df
                        return None

                    # 批量读取各 Sheet
                    if '海外周转汇总' in sn:
                        data_pool["df_海外周转"] = format_df(pl.read_excel(f, sheet_name='海外周转汇总').to_pandas())
                    if '国内在库周转' in sn:
                        data_pool["df_国内在库周转"] = format_df(pl.read_excel(f, sheet_name='国内在库周转').to_pandas())
                    if '断货率' in sn:
                        data_pool["df_断货率"] = format_df(pl.read_excel(f, sheet_name='断货率').to_pandas())

            elif file_type == 'parquet':
                with io.BytesIO(file_bytes) as f:
                    data_pool["stock_turnover"] = pd.read_parquet(f)

        except Exception as e:
            st.error(f"解析文件 {file_name} 时出错: {e}")

    return data_pool


# 3、侧边栏配置
with st.sidebar:
    st.header("⚙️ 配置参数")
    today = date.today()
    default_monday = today - timedelta(days=today.weekday()) - timedelta(days=7)
    t0_date_val = st.date_input("🗓️ 当前周周一", value=default_monday)
    st.session_state.t0_date = t0_date_val

    up_folder_btn = st.file_uploader("请选择上传文件：", type=['xlsx', 'parquet'], accept_multiple_files=True)

    if up_folder_btn:
        # 调用缓存函数处理文件
        with st.spinner("正在快速加载数据..."):
            processed_data = process_uploaded_files(up_folder_btn)

            # 将处理后的数据更新到 session_state（仅在有值时更新，避免覆盖已有的其他数据）
            for key, df in processed_data.items():
                if df is not None:
                    st.session_state[key] = df

            st.success("✅ 数据加载完成 (已缓存)")

    st.divider()

if st.session_state.df_海外周转 is not None:
    df_ref = st.session_state.df_海外周转
    ver = st.session_state.filter_ver
    c1, c2= st.columns([3,3], vertical_alignment="bottom")
    with c1:
        # ui_level = st.multiselect("🌍 层级", sorted(df_ref['层级'].unique().tolist()), key=f"m_{ver}")
        ui_level = st.selectbox("🌍 层级", sorted(df_ref['层级'].unique().tolist()), key=f"m_{ver}")
        st.session_state.committed_filters["level"] = ui_level

def apply_filters(df, filters):
    if df is None: return None
    if filters is None: return df
    df = df.copy()
    # if filters.get("level"): df = df[df["层级"].isin(filters["level"])]
    if filters.get("level"): df = df[df["层级"] == filters["level"]]
    return df

@st.fragment
def actual_turnover_area(df_海外周转, df_国内在库周转, df_断货率, filters):
    # df_country_turnover=apply_filters(df_country_turnover,filters)
    df_海外周转 = apply_filters(df_海外周转, filters)
    df_国内在库周转 = apply_filters(df_国内在库周转, filters)
    df_断货率 = apply_filters(df_断货率, filters)

    if df_海外周转 is None or df_海外周转.empty:
        st.warning("无数据")    
        return

    df_hw = df_海外周转.groupby("周数").agg(
         期初库存金额 = ("期初库存金额", "sum"),
         期末库存金额 = ("期末库存金额", "sum"),
         期初在途金额 = ("期初在途金额", "sum"),
         期末在途金额 = ("期末在途金额", "sum"),
         周销售出库金额 = ("周销售出库金额", "sum"),
    ).reset_index()
    df_hw['海外在库周转'] = ((df_hw['期末库存金额'] + df_hw['期初库存金额'])/2 / (df_hw['周销售出库金额']/7+0.001)).round(2)
    df_hw['海外在途周转'] = ((df_hw['期末在途金额'] + df_hw['期初在途金额'])/2 / (df_hw['周销售出库金额']/7+0.001)).round(2)
    # 合并结果
    df_nostock = df_断货率.groupby("周数").agg(
        SKU总数 = ("sku总数", "sum"),
        断货SKU数量 = ("断货sku数量", "sum"),
    ).reset_index()
    df_nostock['断货率'] = df_nostock['断货SKU数量']/df_nostock['SKU总数'].round(4)
    df_nostock['断货率'] = df_nostock['断货率']*100
    result_df = pd.merge(df_国内在库周转[['周数','层级','国内在库周转天数']], df_hw, on='周数', how='left')
    result_df = pd.merge(result_df, df_nostock, on='周数', how='left')
    result_df['海外周转天数'] = (((result_df['海外在库周转'] + result_df['海外在途周转']).round(1)))
    result_df = result_df.sort_values(by='周数')
    result_df = result_df.reset_index(drop=True)
    fig_历史周转 = make_subplots(
        rows=2, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.1,  # 上下两个图的间距
        row_heights=[0.3, 0.7]
    )
    fig_历史周转.add_trace(go.Bar(
        x=result_df['周数'],
        y=result_df['海外在库周转'],
        name='海外在库周转',
        text=result_df['海外在库周转'].apply(lambda x: f"{x:.1f}"),
        textposition="outside",
        marker=dict(
            color='#85C1E9'
        ),
        textfont=dict(size=18, color="black"),
        hovertemplate=(
            "周数: %{x}<br>"  # x[1] 对应 '周数new'
            "海外在库周转: %{y:.0f}<extra></extra>"  # extra 用于去掉默认 trace 名称
        ),
        hoverlabel=dict(
            font_size=12,
            font_family="Microsoft YaHei",
            font_color="black",
            bgcolor="white"
        )
    ),
        row=2,
        col=1
    )
    fig_历史周转.add_trace(go.Bar(
        x=result_df['周数'],
        y=result_df['海外在途周转'],
        name='海外在途周转',
        text=result_df['海外在途周转'].apply(lambda x: f"{x:.1f}"),
        textposition="outside",
        marker=dict(
            color='#f8c471'
        ),
        textfont=dict(size=18, color="black"),
        hovertemplate=(
            "周数: %{x}<br>"  # x[1] 对应 '周数new'
            "海外在途周转: %{y:.0f}<extra></extra>"  # extra 用于去掉默认 trace 名称
        ),
        hoverlabel=dict(
            font_size=12,
            font_family="Microsoft YaHei",
            font_color="black",
            bgcolor="white"
        )
    ),
        row=2,
        col=1
    )
    fig_历史周转.add_trace(go.Bar(
        x=result_df['周数'],
        y=result_df['国内在库周转天数'],
        name='国内在库周转',
        text=result_df['国内在库周转天数'].apply(lambda x: f"{x:.1f}"),
        textposition="outside",
        marker=dict(
            color='#bfc9ca'
        ),
        textfont=dict(size=18, color="black"),
        hovertemplate=(
            "周数: %{x}<br>"  # x[1] 对应 '周数new'
            "国内在库周转: %{y:.0f}<extra></extra>"  # extra 用于去掉默认 trace 名称
        ),
        hoverlabel=dict(
            font_size=12,
            font_family="Microsoft YaHei",
            font_color="black",
            bgcolor="white"
        )
    ),
        row=2,
        col=1
    )
    fig_历史周转.add_trace(go.Scatter(
        x=result_df['周数'],
        y=result_df['海外周转天数'],
        mode='lines+markers+text',
        name='海外周转天数',
        text=result_df['海外周转天数'].apply(lambda x: f"{x:.1f}"),
        textposition="top center",
        marker=dict(
            color='#2E8421',
            size=12
        ),
        textfont=dict(
            size=15,
            color="black",
        )
    ),
        row=2,
        col=1
    )
    fig_历史周转.add_hline(
        y=105,
        line_dash="dash",  # 设置为虚线
        line_color="#CC0033",  # 保持颜色一致
        annotation_text="海外周转目标: 105天",
        annotation_position="bottom right",  # 文字显示位置
        opacity=0.7  # 设置透明度，避免抢了主数据的视觉焦点
    )

    def get_trend_text_and_color(val):
        # if val > 0.01:
        if val > 1:
            return f"{val:.2f}%", "red"
        else:
            return f"{abs(val):.2f}%", "green"

    text_labels = result_df['断货率'].apply(lambda x: get_trend_text_and_color(x)[0])
    marker_colors = result_df['断货率'].apply(lambda x: get_trend_text_and_color(x)[1])

    fig_历史周转.add_trace(
        go.Scatter(
            x=result_df['周数'],
            y=result_df['断货率'],
            mode='markers+text',
            text=text_labels,
            textposition="top center",
            textfont=dict(color=marker_colors, size=16),
            marker=dict(
                color=marker_colors,
                size=15,
                line=dict(width=1, color='white')
            ),
            # 关键：用error_y实现棒棒糖的垂直杆
            error_y=dict(
                type='data',
                symmetric=False,
                array=[0] * len(result_df),  # 向上延伸量为0
                arrayminus=result_df['断货率'],  # 向下延伸到0点
                width=0,  # 不显示横向的小横杠
                thickness=1.5,
                color='rgba(100, 100, 100, 0.5)'  # 杆子的颜色（浅灰色半透明）
            ),
            showlegend=True,
            name='断货率',
            hovertemplate='周数: %{x}<br>断货率: %{y}%<extra></extra>'
        ),
        row=1, col=1
    )

    # --- 全局布局调整 ---
    fig_历史周转.update_layout(
        barmode='group',
        height=600,
        plot_bgcolor='white',  # 设置背景为白色更接近原图
        margin=dict(t=80, b=50, l=50, r=50),
        font=dict(family="Microsoft YaHei"),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5,
            title_text="",
            font=dict(size=16, color="black")
        )
    )
    max_y = max(result_df['海外周转天数']) + 30
    # --- 坐标轴设置 ---
    fig_历史周转.update_yaxes(
        title_text="周转天数",
        showgrid=True,
        gridcolor='lightgray',
        row=2, col=1,
        tickfont=dict(size=16, color="black"),
        range=[0, max_y]
    )
    max_out_of_stock = result_df['断货率'].max()
    upper_limit = max_out_of_stock * 1.5 if max_out_of_stock > 0 else 1 
    fig_历史周转.update_yaxes(
        # range=[0, 9],  # 稍微给上方文本留点空间
        range=[0, upper_limit],  # 稍微给上方文本留点空间
        showticklabels=False,  # 隐藏数值标签
        showgrid=False,  # 隐藏网格线
        zeroline=True,  # 显示 0 基准线
        zerolinecolor='gray',
        zerolinewidth=1,
        row=1, col=1
    )

    # 调整下方 X 轴
    fig_历史周转.update_xaxes(
        type='category',  # 确保周数按类别等距显示
        row=2, col=1,
        tickfont=dict(size=16, color="black")
    )

    st.plotly_chart(fig_历史周转, width='stretch')

@st.fragment
def delivery_stock_area(df_海外周转, df_断货率, curr_filters):
    df_海外周转 = apply_filters(df_海外周转, curr_filters)
    df_断货率 = apply_filters(df_断货率, curr_filters)
    if df_海外周转 is None or df_海外周转.empty: return
    if df_断货率 is None or df_断货率.empty: return

    df_hw = df_海外周转.groupby(["周数",'子市场']).agg(
         期初库存金额 = ("期初库存金额", "sum"),
         期末库存金额 = ("期末库存金额", "sum"),
         期初在途金额 = ("期初在途金额", "sum"),
         期末在途金额 = ("期末在途金额", "sum"),
         周销售出库金额 = ("周销售出库金额", "sum"),
    ).reset_index()
    df_hw['海外在库周转'] = ((df_hw['期末库存金额'] + df_hw['期初库存金额'])/2 / (df_hw['周销售出库金额']/7+0.001)).round(2)
    df_hw['海外在途周转'] = ((df_hw['期末在途金额'] + df_hw['期初在途金额'])/2 / (df_hw['周销售出库金额']/7+0.001)).round(2)
    # 合并结果
    df_nostock = df_断货率.groupby(["周数",'子市场']).agg(
        SKU总数 = ("sku总数", "sum"),
        断货SKU数量 = ("断货sku数量", "sum"),
    ).reset_index()
    df_nostock['断货率'] = df_nostock['断货SKU数量']/df_nostock['SKU总数'].round(4)
    # df_nostock['断货率'] = df_nostock['断货率']*100
    result_df = pd.merge(df_hw, df_nostock, on=['周数','子市场'], how='left')
    result_df['海外在库周转'] = np.where(result_df['海外在库周转']>1000, 0.001, result_df['海外在库周转'])
    result_df['海外在途周转'] = np.where(result_df['海外在途周转']>1000, 0.001, result_df['海外在途周转'])
    result_df['海外周转天数'] = (((result_df['海外在库周转'] + result_df['海外在途周转']).round(1)))

    result_df=result_df[result_df['海外周转天数']>0]
    result_df=result_df[~result_df['子市场'].isin(['LC-MX','LCM-MX','CP-JCW','MC-MX'])]
    result_df = result_df.sort_values(by=['周数','子市场'])
    result_df = result_df.reset_index(drop=True)
    six_weeks = result_df['周数'].sort_values().unique()[-5:]
    result_df = result_df[result_df['周数'].isin(six_weeks)]

    fig_子市场指标 = make_subplots(
            rows=2, cols=1,
            shared_xaxes=True,
            vertical_spacing=0.04,
            row_heights=[0, 1.0],
            specs=[[{"secondary_y": False}], [{"secondary_y": False}]]
        )

    result_df['周数new'] = result_df["周数"].str[2:]
    # 不同子市场添加目标天数
    result_df['目标天数'] = np.where(
        result_df['子市场'] == 'DE',
        129,
        np.where(
            result_df['子市场'] == 'APM',
            135,
            np.where(
                result_df['子市场'] == 'AP-CA',
                135,
                105
            )
        )
    )
    result_df['目标天数'] = result_df['目标天数'].round(0)
    result_df['子市场_new'] = (
            result_df['子市场'] + '<br>' +
            '(' + result_df['目标天数'].astype(str) + '天)'
    )

    fig_子市场指标.add_trace(
        go.Bar(
            x=[result_df['子市场_new'], result_df['周数new']],
            y=result_df['海外在库周转'],
            name="海外在库周转",
            marker_color='#85C1E9',
            text=[f"{v:.0f}" for v in result_df['海外在库周转']],
            textposition='inside',
            textfont=dict(size=11, color="black"),
            cliponaxis=False,
            showlegend=True,
            hovertemplate=(
                "子市场: %{x[0]}<br>"  # x[0] 对应 '子市场'
                "周数: %{x[1]}<br>"  # x[1] 对应 '周数new'
                "海外在库周转: %{y:.0f}<extra></extra>"  # extra 用于去掉默认 trace 名称
            ),
            hoverlabel=dict(
                font_size=10,
                font_family="Microsoft YaHei",
                font_color="black",
                bgcolor="white"
            )
        ),
        row=2, col=1, secondary_y=False,
    )
    fig_子市场指标.add_trace(
        go.Bar(
            x=[result_df['子市场_new'], result_df['周数new']],
            y=result_df['海外在途周转'],
            name="海外在途周转",
            marker_color='#f8c471',
            text=[f"{v:.0f}" for v in result_df['海外在途周转']],
            textposition='inside',
            textfont=dict(size=11, color="black"),
            cliponaxis=False,
            showlegend=True,
            hovertemplate=(
                "子市场: %{x[0]}<br>"  # x[0] 对应 '子市场'
                "周数: %{x[1]}<br>"  # x[1] 对应 '周数new'
                "海外在途周转: %{y:.0f}<extra></extra>"  # extra 用于去掉默认 trace 名称
            ),
            hoverlabel=dict(
                font_size=10,
                font_family="Microsoft YaHei",
                font_color="black",
                bgcolor="white"
            )
        ),
        row=2, col=1, secondary_y=False,
    )
    fig_子市场指标.add_trace(
        go.Scatter(
            x=[result_df['子市场_new'], result_df['周数new']],
            y=result_df['海外周转天数'],
            mode='text',
            text=[f"{v:.0f}" for v in result_df['海外周转天数']],
            textposition='top center',
            textfont=dict(
                size=14,
                color='#0461AB'
            ),
            showlegend=False,
            hoverinfo='skip'
        ),
        row=2, col=1, secondary_y=False
    )
    top_y1 = (result_df['海外在库周转'].fillna(0) + result_df['海外在途周转'].fillna(0)).max() * 1.1
    # top_y1 = 430
    y_constant_list = [top_y1] * len(result_df)

    # 断货率如何大于0字体就是红色否则是绿色
    def get_trend_text_and_color(val):
        if val > 0.01:
            return "red"
        else:
            return "green"

    fig_子市场指标.add_trace(
        go.Scatter(
            x=[result_df['子市场_new'], result_df['周数new']],
            y=y_constant_list,
            mode="lines+markers+text",
            line=dict(
                color="#D3D3D3",
                width=1.5,
                dash="solid"
            ),
            marker=dict(
                # color="#66CC99",
                color=["red" if v > 0.01 else "green" for v in result_df['断货率']],
                size=6,
                line=dict(
                    # color="#66CC99",
                    color=["red" if v > 0.01 else "green" for v in result_df['断货率']],
                    width=2
                )
            ),
            # 正确提取单行对应的断货率文本
            text=[f"{v * 100:.1f}%" if v == v else "" for v in result_df['断货率']],
            textposition="top center",
            textfont=dict(
                # color="#009966",
                color=["red" if v > 0.01 else "green" for v in result_df['断货率']],

                size=12,
                family="Microsoft YaHei"
            ),
            showlegend=True,
            hoverinfo='skip',
            name='断货率'
        ),
        row=2, col=1, secondary_y=False
    )
    market_counts = result_df['子市场'].value_counts(sort=False)
    unique_markets = result_df['子市场'].unique()
    shapes = []
    current_position = -0.5
    for market in unique_markets[:-1]:
        current_position += market_counts[market]
        shapes.append(
            dict(
                type="line",
                xref="x2",
                yref="paper",
                x0=current_position,
                x1=current_position,
                y0=0,
                y1=1,
                line=dict(
                    color="#CCCCCC",
                    width=2,
                    dash="dash"
                )
            )
        )

    # ==================== 样式与坐标轴更新 ====================
    fig_子市场指标.update_layout(
        barmode='stack',
        height=650,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.05,
            xanchor="center",
            x=0.5,
            font=dict(size=14, color="black")
        ),
        plot_bgcolor='white',
        margin=dict(t=40, b=10, l=10, r=10),
        font=dict(family="Microsoft YaHei"),
        shapes=shapes
    )

    fig_子市场指标.update_xaxes(visible=False, row=1, col=1)
    fig_子市场指标.update_yaxes(visible=False, range=[-0.05, 1.15], row=1, col=1, secondary_y=False)
    fig_子市场指标.update_yaxes(showgrid=False, zeroline=False, range=[0, (
                result_df['海外在库周转'].fillna(0) + result_df['海外在途周转'].fillna(0)).max() * 1.3], row=2,
                                col=1, secondary_y=False)

    # 下方柱状图 X 轴
    fig_子市场指标.update_xaxes(
        tickangle=60,
        showdividers=True,
        dividercolor="#999999",
        row=2, col=1,
        tickfont=dict(size=12, color="black")
    )

    st.plotly_chart(fig_子市场指标, width='stretch')



if st.session_state.df_海外周转 is not None:
    curr_filters = st.session_state.committed_filters
    # 历史实际周转区域
    st.header("📈 历史实际周转", anchor="0")
    actual_turnover_area(st.session_state.df_海外周转, st.session_state.df_国内在库周转, st.session_state.df_断货率, curr_filters)
    st.divider()
    delivery_stock_area(st.session_state.df_海外周转, st.session_state.df_断货率, curr_filters)
else:
    st.info("👋 请先在左侧侧边栏上传数据文件。")
