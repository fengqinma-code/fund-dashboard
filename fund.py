"""
Streamlit 基金运营自动化工具  📊  (v3)
--------------------------------------------------------------
更新记录 (2026-03-09)
1. 第一张表：新增「规模增长」列 = 本周规模 - 上上周规模
2. 底部新增图表：
   • 管理人与策略汇总（两周规模 & 规模增长）
   • 产品周度涨跌幅 Top-10
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import matplotlib.pyplot as plt
from datetime import datetime

# ---------- Matplotlib 中文 & 负号 ----------
plt.rcParams.update({
    "font.family": ["SimHei", "Microsoft YaHei", "Arial Unicode MS"],
    "axes.unicode_minus": False,
})

st.set_page_config(page_title="基金运营自动化工具", layout="wide")
st.title("📈 基金运营自动化工具")

# ----------------------------- 侧栏上传 -----------------------------
side = st.sidebar
side.header("📂 上传Excel文件")
up_this = side.file_uploader("① 上周最后交易日净值表", key="this", type=["xlsx", "xls"])
up_last = side.file_uploader("② 上上周最后交易日净值表", key="last", type=["xlsx", "xls"])
up_y0   = side.file_uploader("③ 年初净值表 (如 20260101)", key="y0",   type=["xlsx", "xls"])
up_txn  = side.file_uploader("④ 周度申赎表", key="txn",  type=["xlsx", "xls"])

# ----------------------------- 常量/工具 -----------------------------
REQ_NET = ["产品名称", "单位净值", "资产净值", "累计单位净值", "净值日期"]
REQ_TXN = ["产品名称", "客户名称", "交易类型", "确认金额"]

STRATEGY_MAP = {
    "中性策略": ["中性"],
    "CTA策略":  ["CTA", "安贤好风如水私募证券投资基金", "安贤权贤而期私募证券投资基金", "安贤方圆守正二号私募证券投资基金"],
    "多策略":   ["多策略", "安贤定制1号私募证券投资基金", "安贤花木长盛私募证券投资基金", "安贤致远1号私募证券投资基金"],
    "指增策略": ["指数增强"],
    "量化多头策略": ["量化多头"],
    "北交所策略": ["北证"],
    "IPM策略":  ["IPM", "安贤价值精选1号私募证券投资基金", "安贤麦穗对冲1号私募证券投资基金",
                 "安贤麦芽多头3号私募证券投资基金", "安贤麦芽多头4号私募证券投资基金",
                 "安贤麦芒灵活对冲2号私募证券投资基金"],
}

def read_xl(file, name):
    try:
        return pd.read_excel(file)
    except Exception as e:
        st.error(f"❌ 读取{name}失败: {e}")
        return None

def check_cols(df, req, name):
    miss = [c for c in req if c not in df.columns]
    if miss:
        st.error(f"{name} 缺少列: {', '.join(miss)}")
        return False
    return True

def get_date_str(df):
    dt = pd.to_datetime(df["净值日期"], errors="coerce").max()
    return dt.strftime("%Y%m%d") if pd.notna(dt) else "未知日期"

def is_master(name):           # 过滤子份额（末尾 A/B/C 等）
    return not name.rstrip().endswith(tuple("ABC"))

def prepare_net(df, date_str):
    df = df[df["产品名称"].apply(is_master)].copy()
    scale_col = "管理规模" if "管理规模" in df.columns else "资产净值"
    df.rename(columns={
        "单位净值":         f"单位净值 {date_str}",
        scale_col:          f"管理规模 截止{date_str}",
        "累计单位净值":     f"累计单位净值 {date_str}",
    }, inplace=True)
    return df[["产品名称",
               f"单位净值 {date_str}",
               f"管理规模 截止{date_str}",
               f"累计单位净值 {date_str}"]]

def detect_strategy(pname):
    for strat, keys in STRATEGY_MAP.items():
        if any(k in pname for k in keys):
            return strat
    return "其他"

def build_metrics(df_this, df_last, df_y0, d_this, d_last, d_y0):
    m = df_this.merge(df_last, on="产品名称")\
               .merge(df_y0, on="产品名称", how="left")
    m["周度涨跌幅(%)"] = (m[f"单位净值 {d_this}"] - m[f"单位净值 {d_last}"]) \
                         / m[f"单位净值 {d_last}"] * 100
    m["YTD (%)"] = (m[f"累计单位净值 {d_this}"] - m[f"累计单位净值 {d_y0}"]) \
                   / m[f"累计单位净值 {d_y0}"] * 100
    m["策略"] = m["产品名称"].apply(detect_strategy)
    return m

def agg_strategy(m, d_this, d_last):
    g = m.groupby("策略", as_index=False)[[f"管理规模 截止{d_this}", f"管理规模 截止{d_last}"]].sum()
    g = g[g["策略"] != "其他"].copy()
    g["规模增长"] = g[f"管理规模 截止{d_this}"] - g[f"管理规模 截止{d_last}"]
    g["周度规模增长率%"] = g["规模增长"] / g[f"管理规模 截止{d_last}"] * 100
    g.insert(0, "统计口径", g.pop("策略"))          # 把策略名放到统计口径列
    return g

def top3_txn(df):
    df["确认金额"] = pd.to_numeric(df["确认金额"], errors="coerce").fillna(0)
    pvt = df.pivot_table(index=["客户名称", "产品名称"],
                         columns="交易类型",
                         values="确认金额",
                         aggfunc="sum",
                         fill_value=0).reset_index()
    for col in ["申购", "赎回"]:
        if col not in pvt.columns:
            pvt[col] = 0
    return pvt.sort_values("申购", ascending=False).head(3)[
        ["客户名称", "产品名称", "申购", "赎回"]
    ]

# ----------------------------- 主流程 -----------------------------
if all([up_this, up_last, up_y0, up_txn]):
    df_this = read_xl(up_this, "上周净值")
    df_last = read_xl(up_last, "上上周净值")
    df_y0   = read_xl(up_y0,   "年初净值")
    df_txn  = read_xl(up_txn,  "申赎")

    if any(d is None for d in (df_this, df_last, df_y0, df_txn)):
        st.stop()

    # 列校验
    if not all([
        check_cols(df_this, REQ_NET, "上周净值"),
        check_cols(df_last, REQ_NET, "上上周净值"),
        check_cols(df_y0,   REQ_NET, "年初净值"),
        check_cols(df_txn,  REQ_TXN, "申赎")
    ]):
        st.stop()

    d_this = get_date_str(df_this)
    d_last = get_date_str(df_last)
    d_y0   = get_date_str(df_y0)

    df_this_p = prepare_net(df_this, d_this)
    df_last_p = prepare_net(df_last, d_last)
    df_y0_p   = prepare_net(df_y0,   d_y0)

    metrics = build_metrics(df_this_p, df_last_p, df_y0_p, d_this, d_last, d_y0)

    # ---------- 表一：管理人与策略汇总 ----------
    total_row = pd.DataFrame({
        "统计口径": ["管理人资管产品层面"],
        f"管理规模 截止{d_this}": [metrics[f"管理规模 截止{d_this}"].sum()],
        f"管理规模 截止{d_last}": [metrics[f"管理规模 截止{d_last}"].sum()],
    })
    total_row["规模增长"] = total_row[f"管理规模 截止{d_this}"] - total_row[f"管理规模 截止{d_last}"]
    total_row["周度规模增长率%"] = total_row["规模增长"] / total_row[f"管理规模 截止{d_last}"] * 100

    strat_sum = agg_strategy(metrics, d_this, d_last)

    overview = pd.concat([total_row, strat_sum], ignore_index=True)

    st.subheader("📊 WEEKLY FUND STATS")
    st.caption(f"统计区间: {d_last} – {d_this}")
    st.dataframe(overview, use_container_width=True)

    # ---------- 表二：产品明细 ----------
    cols_keep = ["产品名称",
                 f"单位净值 {d_this}", f"单位净值 {d_last}",
                 "周度涨跌幅(%)", "YTD (%)",
                 f"累计单位净值 {d_this}", f"累计单位净值 {d_y0}"]
    st.subheader("📋 产品明细")
    st.dataframe(metrics[cols_keep], height=430, use_container_width=True)

    # ---------- 表三：Top-3 申赎 ----------
    st.subheader("💰 SUBSCRIPTION & REDEMPTION STATS (Top-3)")
    top3 = top3_txn(df_txn)
    st.table(top3)

    # ---------- 周度汇总 KPI ----------
    buy_sub  = df_txn.loc[df_txn["交易类型"] == "申购", "确认金额"].sum()
    buy_recg = df_txn.loc[df_txn["交易类型"] == "认购", "确认金额"].sum()
    total_buy = buy_sub + buy_recg
    total_sell = df_txn.loc[df_txn["交易类型"] == "赎回", "确认金额"].sum()
    net_buy = total_buy - total_sell

    st.markdown("### 📑 周度申赎汇总")
    k1, k2, k3 = st.columns(3)
    k1.metric("总申购", f"{total_buy:,.2f}")
    k2.metric("总赎回", f"{total_sell:,.2f}")
    k3.metric("净申购", f"{net_buy:,.2f}")

    # ---------- 图表 1：表一柱状图 ----------
    st.markdown("#### 📉 管理人与策略规模对比")
    fig1, ax1 = plt.subplots(figsize=(8, 4))
    bar_w = 0.35
    x = np.arange(len(overview))
    ax1.bar(x - bar_w/2, overview[f"管理规模 截止{d_last}"] / 1e8, bar_w, label=d_last)
    ax1.bar(x + bar_w/2, overview[f"管理规模 截止{d_this}"] / 1e8, bar_w, label=d_this)
    ax1.set_xticks(x)
    ax1.set_xticklabels(overview["统计口径"], rotation=30, ha="right")
    ax1.set_ylabel("规模（亿元）")
    ax1.legend()
    ax1.set_ylim(0)
    st.pyplot(fig1)

    # ---------- 图表 2：产品周度涨跌 Top-10 ----------
    st.markdown("#### 🔥 周度涨跌幅 Top-10")
    top10 = metrics.nlargest(10, "周度涨跌幅(%)").sort_values("周度涨跌幅(%)", ascending=True)
    fig2, ax2 = plt.subplots(figsize=(8, 4))
    ax2.barh(top10["产品名称"], top10["周度涨跌幅(%)"])
    ax2.set_xlabel("周度涨跌幅 (%)")
    ax2.set_ylabel("")
    ax2.set_xlim(left=min(0, top10["周度涨跌幅(%)"].min()*1.1))
    st.pyplot(fig2)

    # ----- 图表 3：产品周度涨跌 Last-10 -----------------------------
    st.markdown("#### 🔥 周度涨跌幅 Last-5")
    bottom5 = (
        metrics
        .sort_values("周度涨跌幅(%)")
        .head(5)[["产品名称", "周度涨跌幅(%)"]]
        .reset_index(drop=True)
    )

    fig2, ax2 = plt.subplots(figsize=(6, 3.5))
    ax2.barh(bottom5["产品名称"], bottom5["周度涨跌幅(%)"], color="#d95f02")
    ax2.set_xlabel("周度涨跌幅 (%)")
    ax2.set_title("周度涨跌幅排名后 5 名")
    ax2.invert_yaxis()  # 名次 1-5 自上而下
    st.pyplot(fig2)
    # ---------- 下载 Excel ----------
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        overview.to_excel(w, sheet_name="管理人与策略汇总", index=False)
        metrics[cols_keep].to_excel(w, sheet_name="产品明细", index=False)
        top3.to_excel(w, sheet_name="申赎Top3", index=False)
        pd.DataFrame({"指标": ["总申购", "总赎回", "净申购"],
                      "金额": [total_buy, total_sell, net_buy]}).to_excel(
            w, sheet_name="周度申赎汇总", index=False
        )
    st.download_button("📥 下载Excel",
                       buf.getvalue(),
                       file_name=f"weekly_fund_{d_this}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("请在左侧依次上传：两周净值表 + 年初净值表 + 申赎表")