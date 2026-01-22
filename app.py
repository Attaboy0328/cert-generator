import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os
from docx import Document
from docxcompose.composer import Composer
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# --- 1. 极致兼容的样式注入 ---
def inject_stable_style():
    st.markdown("""
    <style>
    /* 采用最稳定的 CSS 背景方式 */
    .stApp {
        background: linear-gradient(135deg, #007FFE 0%, #60B2FE 50%, #C0E5FE 100%);
        background-attachment: fixed;
    }

    /* 彻底消除白色背景块，解决显示生硬问题 */
    div[data-testid="stVerticalBlock"], 
    div[data-testid="stMarkdownContainer"], 
    div[data-testid="stHeader"],
    .st-emotion-cache-12w0qpk {
        background-color: transparent !important;
    }

    /* 步骤标题：半透明磨砂感 */
    h3 {
        background: rgba(255, 255, 255, 0.2) !important;
        backdrop-filter: blur(10px);
        padding: 10px 15px !important;
        border-radius: 10px !important;
        color: white !important;
        border: 1px solid rgba(255, 255, 255, 0.2);
    }

    h1 { color: white !important; text-align: center; }
    
    /* 让按钮更有质感 */
    .stButton>button {
        background-color: rgba(255, 255, 255, 0.3) !important;
        color: white !important;
        border: 1px solid white !important;
    }
    </style>
    """, unsafe_allow_html=True)

# 基础配置
st.set_page_config(page_title="证书智能制作工具", layout="centered")
inject_stable_style()

st.title("🎓 内审员证书智能制作工具")

# --- 第一步：录入模式 ---
st.markdown("### 第一步：选择录入模式")
mode = st.radio("选择方式：", ["Excel 文件上传", "网页表格填写"], index=0, horizontal=True)

DEFAULT_TEMPLATE = "内审员证书.docx"
data_to_process = []

# --- 第二步：准备数据 ---
st.markdown("### 第二步：填写或上传信息")

if mode == "网页表格填写":
    st.info("💡 提示：点击左上角第一个单元格并按下 Ctrl+V 即可粘贴数据。")
    init_df = pd.DataFrame({
        "序号": [i for i in range(1, 101)],
        "证书编号": [None] * 100, "姓名": [None] * 100, "身份证号": [None] * 100, "培训日期": [None] * 100, "标准号": [None] * 100,
    })
    edited_df = st.data_editor(init_df, num_rows="fixed", use_container_width=True, hide_index=True)
    
    # 清洗逻辑：去除空行
    temp_df = edited_df.drop(columns=["序号"]).dropna(how='all')
    data_to_process = []
    for _, row in temp_df.iterrows():
        clean_row = {k: str(v).strip() for k, v in row.items() if pd.notna(v) and str(v) != 'None' and str(v) != ''}
        if clean_row.get('姓名'): # 至少要有名字
            data_to_process.append(clean_row)

else:
    col1, col2 = st.columns([2, 3])
    with col1:
        # 构造带示例模板
        df_ex = pd.DataFrame({"证书编号": ["T-2026-001(示例)"], "姓名": ["张三(示例)"], "身份证号": ["440683199001010001"], "培训日期": ["2026年1月"], "标准号": ["ISO9001"]})
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_ex.to_excel(writer, index=False)
            ws = writer.sheets['Sheet1']
            for cell in ws[2]: cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        st.download_button("📥 下载信息模板", data=buffer.getvalue(), file_name="学员信息模板.xlsx")
    
    with col2:
        uploaded_data = st.file_uploader("上传已填写的 Excel", type=["xlsx", "csv"], label_visibility="collapsed")

    if uploaded_data:
        df = pd.read_csv(uploaded_data, dtype=str).fillna("") if uploaded_data.name.endswith('.csv') else pd.read_excel(uploaded_data, dtype=str).fillna("")
        data_to_process = [row for row in df.to_dict('records') if "示例" not in str(row.get('姓名', ''))]
        if data_to_process: st.success(f"✅ 已加载 {len(data_to_process)} 条数据")

# --- 第三步：生成 ---
st.markdown("### 第三步：确认模板并生成")

if os.path.exists(DEFAULT_TEMPLATE):
    t_opt = st.radio("模板：", ["使用内置模板", "上传本地模板"], horizontal=True)
    t_path = DEFAULT_TEMPLATE if t_opt == "使用内置模板" else st.file_uploader("上传 Word 模板", type=["docx"])
else:
    t_path = st.file_uploader("请上传 Word 模板", type=["docx"])

if t_path and data_to_process:
    if st.button("🚀 开始批量制作汇总文档", use_container_width=True):
        try:
            master_doc, prog, count = None, st.progress(0), 0
            for i, row in enumerate(data_to_process):
                name_val = str(row.get('姓名', '')).strip()
                if not name_val: continue
                count += 1
                doc = DocxTemplate(t_path)
                doc.render({
                    'number': row.get('证书编号',''), 'name': name_val, 
                    'id_card': row.get('身份证号',''), 'date': row.get('培训日期',''), 
                    'standards': row.get('标准号','')
                })
                t_io = io.BytesIO(); doc.save(t_io); t_io.seek(0)
                cur = Document(t_io)
                if master_doc is None:
                    master_doc = cur
                    composer = Composer(master_doc)
                else:
                    master_doc.add_page_break(); composer.append(cur)
                prog.progress((i + 1) / len(data_to_process))
            
            if master_doc:
                out = io.BytesIO(); master_doc.save(out); out.seek(0)
                st.balloons()
                st.download_button(f"🎁 下载汇总文档({count}份)", out.getvalue(), "证书汇总.docx", use_container_width=True)
        except Exception as e:
            st.error(f"制作失败：{e}")
else:
    st.info("请先完成前两步操作...")
