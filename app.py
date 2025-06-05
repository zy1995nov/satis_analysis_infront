import streamlit as st

# Streamlit 规定：set_page_config 必须是首个 Streamlit 调用
st.set_page_config(page_title="满意度波动归因分析", page_icon="📊", layout="wide")

# ---------------------- 依赖 ----------------------
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import date
import tempfile
import traceback
import os

# ----------------- 尝试导入核心分析函数 -----------------
try:
    from satis_analysis import run_analysis
    IMPORT_ERR = None  # 导入成功
except Exception as e:
    run_analysis = None  # 保证后续引用不报错
    IMPORT_ERR = e      # 保存异常信息供前端提示

# ---------------------- 页面标题 ----------------------
st.title("📊 满意度波动归因分析工具 · GUI 版")

with st.expander("使用说明", expanded=False):
    st.markdown(
        """
        **操作流程**
        1. 左侧上传包含 `dt` 列的 Excel 数据文件
        2. 选择 **现期** & **基期** 日期区间
        3. 输入维度组合（如 `国内国际,队列 ; 国内国际,渠道`）
        4. 设定最小评价量门槛 → 点击 **开始分析**

        <small>注：第一次上传大文件后如出现卡顿，请耐心等待后台解析完成。</small>
        """,
        unsafe_allow_html=True,
    )

# ----------------- 左侧栏输入 -----------------
st.sidebar.header("⚙️ 参数设置")

uploaded_file = st.sidebar.file_uploader(
    "上传数据文件 (.xlsx)", type=["xlsx", "xls"], accept_multiple_files=False
)

# 如果导入核心脚本失败，给出错误提示
if IMPORT_ERR is not None:
    st.sidebar.error(f"❌ 无法导入 satis_analysis.py\n{IMPORT_ERR}")

col1, col2 = st.sidebar.columns(2)
cur_start = col1.date_input("现期起", value=date.today())
cur_end   = col2.date_input("现期止", value=date.today())

col3, col4 = st.sidebar.columns(2)
base_start = col3.date_input("基期起", value=date.today())
base_end   = col4.date_input("基期止", value=date.today())

dims_str = st.sidebar.text_input(
    "维度组合 (逗号分隔，同组合用 ; 分隔)", value="国内国际,队列"
)

min_eva = st.sidebar.number_input("最小评价量门槛", min_value=0, value=0, step=1)

start_btn = st.sidebar.button(
    "🚀 开始分析",
    disabled=(uploaded_file is None) or (run_analysis is None),
)

# ----------------- 主逻辑 -----------------
if start_btn:
    try:
        # ------ 将上传文件写入临时文件，供 run_analysis 使用 ------
        with st.spinner("读取上传数据 …"):
            tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            tmp_file.write(uploaded_file.read())
            tmp_file.flush()
            tmp_path = tmp_file.name
            tmp_file.close()  # ✅ Windows 上必须先关闭句柄才能后续删除

        # 解析维度组合
        core_dims = [seg.strip().split(",") for seg in dims_str.split(";") if seg.strip()]
        if not core_dims:
            st.error("❌ 维度组合不能为空！")
            st.stop()

        # 日期格式化
        cur_start_str  = cur_start.strftime("%Y-%m-%d")
        cur_end_str    = cur_end.strftime("%Y-%m-%d")
        base_start_str = base_start.strftime("%Y-%m-%d")
        base_end_str   = base_end.strftime("%Y-%m-%d")

        st.success("参数校验通过，开始后台分析 …")
        progress_bar = st.progress(0.01)

        # ------ 调用核心分析函数 ------
        results = run_analysis(
            data_path      = tmp_path,
            cur_start      = cur_start_str,
            cur_end        = cur_end_str,
            base_start     = base_start_str,
            base_end       = base_end_str,
            core_dims      = core_dims,
            min_eva        = min_eva,
            excel_out      = "analysis_result.xlsx",
            enable_method2 = True,
        )
        progress_bar.progress(1.0)

        # ------ 展示结果 ------
        if not results:
            st.warning("分析返回空结果，请检查数据或参数设置。")
        else:
            st.subheader("📂 分析结果预览")
            tab_titles = [" + ".join(dims) for dims in results.keys()]
            tabs = st.tabs(tab_titles)
            for tab, dims in zip(tabs, results.keys()):
                with tab:
                    st.dataframe(results[dims], use_container_width=True, height=520)

            # 打包 Excel 供下载
            bio = BytesIO()
            with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
                for dims, table in results.items():
                    sheet = (" + ".join(dims))[:31] or "Sheet1"
                    table.to_excel(writer, sheet_name=sheet, index=False)
            bio.seek(0)
            st.download_button(
                label="💾 下载完整结果 Excel",
                data=bio,
                file_name="analysis_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        # 删除临时文件
        os.unlink(tmp_path)

    except Exception:
        st.error("❌ 运行出错！\n" + traceback.format_exc())
