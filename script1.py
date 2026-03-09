"""
Streamlit 小程序：基金运营自动化工具（按附件列名适配）
---------------------------------------------------------
* 支持上传 3 个 Excel：上周净值、上上周净值、周度申赎
* 净值 Excel 列：`产品名称`、`单位净值`、`资产净值`（作为规模）；若有 `管理规模` 优先使用
* 申赎 Excel 列：`产品名称`、`交易类型`（申购/赎回）、`确认金额`
* 净申购 = Σ申购确认金额 − Σ赎回确认金额（按产品汇总）
* 其他逻辑同前
"""
import streamlit as st
import pandas as pd
import numpy as np
import io
import matplotlib.pyplot as plt

# -------- 中文字体兼容 -----------------
plt.rcParams["font.family"] = ["SimHei", "Microsoft YaHei", "Arial Unicode MS"]  # 优先黑体/雅黑
plt.rcParams["axes.unicode_minus"] = False  # 解决负号乱码
# --------------------------------------

st.set_page_config(page_title="基金运营自动化工具", layout="wide")
st.title("📈 基金运营自动化工具 · 附件适配版")

###############################################################################
# 上传
###############################################################################
st.sidebar.header("📂 上传 3 份 Excel")
up_this = st.sidebar.file_uploader("① 上周净值", type=["xlsx", "xls"], key="this")
up_last = st.sidebar.file_uploader("② 上上周净值", type=["xlsx", "xls"], key="last")
up_txn  = st.sidebar.file_uploader("③ 周度申赎", type=["xlsx", "xls"], key="txn")

###############################################################################
# 工具函数
###############################################################################

def read_xl(file, name):
    try:
        return pd.read_excel(file)
    except Exception as e:
        st.error(f"❌ 读取{name}失败: {e}")
        return None


REQ_NET = ["产品名称", "单位净值", "资产净值"]  # 资产净值≈规模
REQ_TXN = ["产品名称", "交易类型", "确认金额"]


def check_cols(df: pd.DataFrame, req: list[str], name: str):
    miss = [c for c in req if c not in df.columns]
    if miss:
        st.error(f"{name} 缺少列: {', '.join(miss)}")
        return False
    return True


def is_master_product(pname: str) -> bool:
    return not pname.rstrip().endswith(tuple("ABC"))


def prepare_net(df: pd.DataFrame, suffix: str):
    # 过滤子产品
    df = df[df["产品名称"].apply(is_master_product)].copy()
    # 统一列
    df.rename(columns={"单位净值": f"{suffix}净值"}, inplace=True)
    if "管理规模" in df.columns:
        df.rename(columns={"管理规模": f"{suffix}规模"}, inplace=True)
    else:
        df.rename(columns={"资产净值": f"{suffix}规模"}, inplace=True)
    return df[["产品名称", f"{suffix}净值", f"{suffix}规模"]]


def calc_metrics(df_this: pd.DataFrame, df_last: pd.DataFrame):
    merged = pd.merge(df_this, df_last, on="产品名称", how="inner")
    merged["周度涨跌幅(%)"] = (merged["本周净值"] - merged["上周净值"]) / merged["上周净值"] * 100
    merged["规模增长率(%)"] = (merged["本周规模"] - merged["上周规模"]) / merged["上周规模"] * 100
    return merged


def process_txn(df: pd.DataFrame):
    df["确认金额"] = pd.to_numeric(df["确认金额"], errors="coerce").fillna(0)
    grp = df.groupby(["产品名称", "交易类型"], as_index=False)["确认金额"].sum()
    # 透视到列
    pivot = grp.pivot(index="产品名称", columns="交易类型", values="确认金额").fillna(0)
    if "申购" not in pivot.columns:
        pivot["申购"] = 0
    if "赎回" not in pivot.columns:
        pivot["赎回"] = 0
    pivot["净申购"] = pivot["申购"] - pivot["赎回"]
    pivot.reset_index(inplace=True)
    top3 = pivot.sort_values("净申购", ascending=False).head(3)
    return top3, pivot


def bar(data, x, y, title):
    fig, ax = plt.subplots(figsize=(10, 5))
    _d = data.sort_values(y, ascending=False)
    ax.bar(_d[x], _d[y])
    ax.set_title(title)
    ax.set_xticklabels(_d[x], rotation=45, ha="right")
    return fig

###############################################################################
# 主流程
###############################################################################
if up_this and up_last and up_txn:
    df_this = read_xl(up_this, "上周净值")
    df_last = read_xl(up_last, "上上周净值")
    df_txn  = read_xl(up_txn,  "周度申赎")

    if all([df_this is not None, df_last is not None, df_txn is not None]):
        if check_cols(df_this, REQ_NET, "上周净值") and \
           check_cols(df_last, REQ_NET, "上上周净值") and \
           check_cols(df_txn,  REQ_TXN, "周度申赎"):

            df_this = prepare_net(df_this, "本周")
            df_last = prepare_net(df_last, "上周")
            metrics = calc_metrics(df_this, df_last)

            top3, txn_summary = process_txn(df_txn)

            # 展示
            st.subheader("🗃️ 周度净值指标")
            st.dataframe(metrics, use_container_width=True)

            st.subheader("🏅 净申购 Top 3")
            st.dataframe(top3, use_container_width=True)

            col1, col2 = st.columns(2)
            with col1:
                st.subheader("📊 本周规模对比")
                st.pyplot(bar(metrics, "产品名称", "本周规模", "本周规模"))
            with col2:
                st.subheader("📊 净申购排名")
                st.pyplot(bar(txn_summary, "产品名称", "净申购", "净申购"))

            # 下载
            buf_excel = io.BytesIO()
            with pd.ExcelWriter(buf_excel, engine="xlsxwriter") as w:
                metrics.to_excel(w, sheet_name="净值指标", index=False)
                top3.to_excel(w, sheet_name="净申购Top3", index=False)
            st.download_button("📥 下载 Excel 报表", buf_excel.getvalue(), file_name="fund_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("请在左侧依次上传 3 份 Excel 文件（上周净值 / 上上周净值 / 申赎数据）")
