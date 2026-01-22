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
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: rgba(255, 255, 255, 0.5) !important;
        border: 1px solid #ffffff !important;
        transform: translateY
