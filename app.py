import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os
from docx import Document
from docxcompose.composer import Composer
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# --- 1. 注入动态丝绸流光背景 (CSS 实现) ---
def inject_silk_bg():
    st.markdown("""
    <style>
    /* 定义背景容器 */
    .stApp {
        background: linear-gradient(-45deg, #007FFE, #60B2FE, #C0E5FE, #F0FFFE);
        background-size: 400% 400%;
        animation: gradient 10s ease infinite;
        height: 100vh;
    }

    /* 丝绸流动动画 */
    @keyframes gradient {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    /* 让白色的组件容器带一点透明毛玻璃感，更高级 */
    .stMarkdown, .stButton, .stDataFrame, .stRadio, .stFileUploader {
        background-color: rgba(255, 255, 255, 0.6) !important;
        padding: 10px;
        border-radius: 10px;
    }
    
    /* 调整标题颜色 */
    h1 {
        color: #004080 !important;
        text-shadow: 2px 2px 4px rgba(255,255,255,0.5);
    }
    </style>
    """, unsafe_allow_html=True)

# 页面配置
st.set_page_config(page_title="证书智能制作工具", layout="centered")
inject_silk_bg()

st.title("🎓 内审员证书智能制作工具")

# --- 第二步：录入模式 ---
st.markdown("### 第一步：选择录入模式")
mode = st.radio("选择方式：", ["Excel 文件上传", "网页表格填写 (支持粘贴)"], index=0, horizontal=True)

DEFAULT_TEMPLATE = "内审员证书.docx"
data_to_process = []

# --- 第三步：准备数据 ---
st.markdown("---")
st.markdown("### 第二步：填写或上传信息")

if mode == "网页表格填写 (支持粘贴)":
    st.info("💡 提示：点击左上角第一个单元格并按下 Ctrl+V 即可粘贴 Excel 数据。")
    init_df = pd.DataFrame({
        "序号": [i for i in range(1, 101)],
        "证书编号": [None] * 100, "姓名": [None] * 100, "身份证号": [None] * 100, "培训日期": [None] * 100, "标准号": [None] * 100,
    })
    edited_df = st.data_editor(
        init_df, num_rows="fixed", use_container_width=True, hide_index=True, height=380,
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
    data_to_process = [{k: str(v).strip() for k, v in row.items() if v is not None} for row in data_to_process if any(row.values())]

else:
    col1, col2 = st.columns([2, 3])
    with col1:
        # 创建带标黄和自适应列宽的模板
        example_data = {
            "证书编号": ["T-2025-001 (示例)"],
            "姓名": ["张三 (示例)"],
            "身份证号": ["440683199001010001"],
            "培训日期": ["2025年9月3-5日"],
            "标准号": ["ISO9001:2015、ISO22000:2018"]
        }
        df_ex = pd.DataFrame(example_data)
        template_buffer = io.BytesIO()
        with pd.ExcelWriter(template_buffer, engine='openpyxl') as writer:
            df_ex.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            for i, col in enumerate(df_ex.columns):
                column_letter = get_column_letter(i + 1)
                max_length = max(df_ex[col].astype(str).map(len).max(), len(col)) + 5
                worksheet.column_dimensions[column_letter].width = max_length
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            for cell in worksheet[2]: cell.fill = yellow_fill

        st.download_button(label="📥 下载标准模板 (含标黄示例)", data=template_buffer.getvalue(), file_name="学员信息上传模板.xlsx")
        st.caption("注：系统会自动识别并跳过黄色示例行。")
    
    with col2:
        uploaded_data = st.file_uploader("上传文件", type=["xlsx", "csv"], label_visibility="collapsed")

    if uploaded_data:
        df = pd.read_csv(uploaded_data, dtype=str).fillna("") if uploaded_data.name.endswith('.csv') else pd.read_excel(uploaded_data, dtype=str).fillna("")
        data_to_process = [row for row in df.to_dict('records') if "示例" not in str(row.get('姓名', ''))]
        if data_to_process: st.success(f"✅ 已加载 {len(data_to_process)} 条数据")

# --- 第四步：生成 ---
st.markdown("---")
st.markdown("### 第三步：模板确认与生成")

if os.path.exists(DEFAULT_TEMPLATE):
    template_option = st.radio("模板：", ["使用内置模板", "上传本地模板"], horizontal=True)
    template_path = DEFAULT_TEMPLATE if template_option == "使用内置模板" else st.file_uploader("上传 Word", type=["docx"])
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
                doc.render({'number': str(row.get('证书编号','')), 'name': name_val, 'id_card': str(row.get('身份证号','')), 'date': str(row.get('培训日期','')), 'standards': str(row.get('标准号',''))})
                
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
                st.download_button(label=f"🎁 点击下载汇总文档({valid_count}份)", data=out_io.getvalue(), file_name="证书汇总.docx", use_container_width=True)
        except Exception as e:
            st.error(f"制作失败：{e}")
