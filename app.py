import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os
from docx import Document
from docxcompose.composer import Composer
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# --- 1. 注入 CSS：打造全透明流光感 ---
def inject_custom_style():
    st.markdown("""
    <style>
    /* 全局动态流纹背景 */
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

    /* 消除 Streamlit 默认的白色背景块 */
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

    /* 步骤标题：毛玻璃质感 */
    h3 {
        background: rgba(255, 255, 255, 0.25) !important;
        backdrop-filter: blur(12px) !important;
        padding: 12px 20px !important;
        border-radius: 12px !important;
        color: #ffffff !important;
        border: 1px solid rgba(255, 255, 255, 0.3) !important;
        margin: 15px 0 !important;
    }

    h1 {
        color: #ffffff !important;
        text-shadow: 0px 4px 12px rgba(0,0,0,0.15);
        text-align: center;
    }

    /* 组件透明度调整 */
    .stButton>button {
        background-color: rgba(255, 255, 255, 0.3) !important;
        color: white !important;
        border: 1px solid rgba(255, 255, 255, 0.5) !important;
        border-radius: 10px;
    }

    div[data-testid="stDataEditor"] {
        background-color: rgba(255, 255, 255, 0.2) !important;
    }
    </style>
    """, unsafe_allow_html=True)

# 基础配置
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
    st.info("💡 提示：点击左上角单元格并按下 Ctrl+V 即可从 Excel 粘贴数据。")
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
    # 处理表格数据，修复了之前的语法错误
    temp_df = edited_df.drop(columns=["序号"])
    valid_rows = temp_df.dropna(how='all')
    data_to_process = []
    for _, row in valid_rows.iterrows():
        clean_row = {k: str(v).strip() for k, v in row.items() if pd.notna(v) and str(v).lower() != 'none'}
        if clean_row:
            data_to_process.append(clean_row)

else:
    col1, col2 = st.columns([2, 3])
    with col1:
        # 下载带标黄示例模板
        df_ex = pd.DataFrame({"证书编号": ["T-2026-001 (示例)"], "姓名": ["张三 (示例)"], "身份证号": ["440683199001010001"], "培训日期": ["2026年1月23日"], "标准号": ["ISO9001:2015"]})
        template_buffer = io.BytesIO()
        with pd.ExcelWriter(template_buffer, engine='openpyxl') as writer:
            df_ex.to_excel(writer, index=False)
            ws = writer.sheets['Sheet1']
            for i in range(1, 6): ws.column_dimensions[get_column_letter(i)].width = 20
            for cell in ws[2]: cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        st.download_button(label="📥 下载标准模板 (含标黄示例)", data=template_buffer.getvalue(), file_name="学员信息上传模板.xlsx")
        st.caption("注：系统会自动跳过黄色示例行。")
    
    with col2:
        uploaded_data = st.file_uploader("上传文件", type=["xlsx", "csv"], label_visibility="collapsed")

    if uploaded_data:
        df = pd.read_csv(uploaded_data, dtype=str).fillna("") if uploaded_data.name.endswith('.csv') else pd.read_excel(uploaded_data, dtype=str).fillna("")
        data_to_process = [row for row in df.to_dict('records') if "示例" not in str(row.get('姓名', ''))]
        if data_to_process: st.success(f"✅ 已成功加载 {len(data_to_process)} 条有效数据")

# --- 第三步：生成 ---
st.markdown("### 第三步：模板确认与生成")

if os.path.exists(DEFAULT_TEMPLATE):
    template_option = st.radio("模板：", ["使用内置模板", "上传本地模板"], horizontal=True, label_visibility="collapsed")
    template_path = DEFAULT_TEMPLATE if template_option == "使用内置模板" else st.file_uploader("上传自定义 Word", type=["docx"])
else:
    template_path = st.file_uploader("上传 Word 模板", type=["docx"])

if template_path and data_to_process:
    if st.button("🚀 开始批量制作合并文档", use_container_width=True):
        try:
            master_doc, progress_bar, valid_count = None, st.progress(0), 0
            for i, row in enumerate(data_to_process):
                name_val = str(row.get('姓名', '')).replace('nan', '').strip()
                if not name_val: continue
                valid_count += 1
                doc = DocxTemplate(template_path)
                doc.render({
                    'number': str(row.get('证书编号','')).replace('nan',''),
                    'name': name_val,
                    'id_card': str(row.get('身份证号','')).replace('nan',''),
                    'date': str(row.get('培训日期','')).replace('nan',''),
                    'standards': str(row.get('标准号','')).replace('nan','')
                })
                t_io = io.BytesIO(); doc.save(t_io); t_io.seek(0)
                cur_doc = Document(t_io)
                if master_doc is None:
                    master_doc = cur_doc
                    composer = Composer(master_doc)
                else:
                    master_doc.add_page_break(); composer.append(cur_doc)
                progress_bar.progress((i + 1) / len(data_to_process))

            if master_doc:
                out_io = io.BytesIO(); master_doc.save(out_io); out_io.seek(0)
                st.balloons()
                st.download_button(label=f"🎁 制作完成({valid_count}份)！点击下载汇总文档", data=out_io.getvalue(), file_name="证书汇总导出.docx", use_container_width=True)
        except Exception as e:
            st.error(f"制作失败：{e}")
else:
    st.info("等待录入数据并确认模板...")
