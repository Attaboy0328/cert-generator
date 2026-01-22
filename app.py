import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os
from docx import Document
from docxcompose.composer import Composer
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# --- 1. 注入 CSS：消除白色背景框，打造流光毛玻璃感 ---
def inject_custom_style():
    st.markdown("""
    <style>
    /* 全局动态流光背景 */
    .stApp {
        background: linear-gradient(-45deg, #007FFE, #60B2FE, #C0E5FE, #F0FFFE);
        background-size: 400% 400%;
        animation: gradient 15s ease infinite;
    }
    @keyframes gradient {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    /* 彻底消除白色背景框 */
    div[data-testid="stVerticalBlock"], 
    div[data-testid="stMarkdownContainer"], 
    div[data-testid="stForm"],
    div[data-testid="stHeader"],
    .st-emotion-cache-12w0qpk, 
    .st-emotion-cache-6qob1r {
        background-color: transparent !important;
        border: none !important;
        box-shadow: none !important;
    }

    /* 步骤标题美化：半透明毛玻璃效果 */
    h3 {
        background: rgba(255, 255, 255, 0.25) !important;
        backdrop-filter: blur(12px) !important;
        -webkit-backdrop-filter: blur(12px) !important;
        padding: 12px 20px !important;
        border-radius: 12px !important;
        color: #ffffff !important;
        border: 1px solid rgba(255, 255, 255, 0.3) !important;
        margin-top: 20px !important;
        margin-bottom: 10px !important;
    }

    /* 大标题样式 */
    h1 {
        color: #ffffff !important;
        text-shadow: 0px 4px 12px rgba(0,0,0,0.15);
        font-weight: 800 !important;
        text-align: center;
    }

    /* 按钮样式优化 */
    .stButton>button {
        background-color: rgba(255, 255, 255, 0.3) !important;
        color: white !important;
        border: 1px solid rgba(255, 255, 255, 0.5) !important;
        backdrop-filter: blur(5px);
        border-radius: 10px;
    }

    /* 数据编辑器背景调整 */
    div[data-testid="stDataEditor"] {
        background-color: rgba(255, 255, 255, 0.2) !important;
        border-radius: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

# 页面基本配置
st.set_page_config(page_title="证书智能制作工具", layout="centered")
inject_custom_style()

st.title("🎓 内审员证书智能制作工具")

# --- 第一步：录入模式 ---
st.markdown("### 第一步：选择录入模式")
mode = st.radio("选择方式：", ["Excel 文件上传", "网页表格填写 (支持粘贴)"], index=0, horizontal=True, label_visibility="collapsed")

DEFAULT_TEMPLATE = "内审员证书.docx"
data_to_process = []

# --- 第二步：准备数据 ---
st.markdown("### 第二步：填写或上传信息")

if mode == "网页表格填写 (支持粘贴)":
    st.info("💡 提示：点击左上角第一个单元格并按下 Ctrl+V 即可从 Excel 粘贴数据。")
    init_df = pd.DataFrame({
        "序号": [i for i in range(1, 101)],
        "证书编号": [None] * 100, "姓名": [None] * 100, "身份证号": [None] * 100, "培训日期": [None] * 100, "标准号": [None] * 100,
    })
    edited_df = st.data_editor(
        init_df, num_rows="fixed", use_container_width=True, hide_index=True, height=385,
        column_config={
            "序号": st.column_config.NumberColumn("序号", width=40, disabled=True),
            "证书编号": st.column_config.TextColumn("证书编号", width="small"),
            "姓名": st.column_config.TextColumn("姓名", width="small"),
            "身份证号": st.column_config.TextColumn("身份证号", width="medium"),
            "培训日期": st.column_config.TextColumn("培训日期", width="medium"),
            "标准号": st.column_config.TextColumn("标准号", width="large"),
        }
    )
    temp_df = edited_df.drop(columns=["序号"])
    data_to_process = temp_df.dropna(how='all').to_dict('records')
    data_to_process = [{k: str(v).strip() for k, v
