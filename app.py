import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os
from docx import Document
from docxcompose.composer import Composer

# 页面配置
st.set_page_config(page_title="证书智能制作工具", layout="centered")

st.title("🎓 内审员证书智能制作工具")

# --- 第一步：选择模式 ---
st.markdown("### 第一步：选择录入模式")
mode = st.radio("选择方式：", ["网页表格填写 (支持粘贴)", "Excel 文件上传"], horizontal=True)

DEFAULT_TEMPLATE = "内审员证书.docx"
data_to_process = []

# --- 第二步：准备数据 ---
st.markdown("---")
st.markdown("### 第二步：填写或上传信息")

if mode == "网页表格填写 (支持粘贴)":
    st.info("💡 提示：您可以直接点击单元格输入，或从 Excel 复制数据后点击左上角第一个单元格粘贴。")
    
    # 创建一个初始的空 DataFrame，设置 100 行
    init_df = pd.DataFrame(
        columns=["证书编号", "姓名", "身份证号", "培训日期", "标准号"],
        index=range(100)
    )
    
    # 使用数据编辑器
    edited_df = st.data_editor(
        init_df,
        num_rows="fixed", # 固定 100 行
        use_container_width=True,
        hide_index=False,
        column_config={
            "证书编号": st.column_config.TextColumn("证书编号", width="medium"),
            "姓名": st.column_config.TextColumn("姓名", width="small"),
            "身份证号": st.column_config.TextColumn("身份证号", width="medium"),
            "培训日期": st.column_config.TextColumn("培训日期", width="medium"),
            "标准号": st.column_config.TextColumn("标准号", width="large"),
        }
    )
    
    # 过滤掉全空的行
    data_to_process = edited_df.dropna(how='all').to_dict('records')
    # 进一步过滤：至少要有姓名和编号
    data_to_process = [row for row in data_to_process if str(row.get('姓名', '')).strip() != 'None' and str(row.get('姓名', '')).strip() != '']

else:
    uploaded_data = st.file_uploader("上传学员信息 Excel 文件", type=["xlsx", "csv"])
    if uploaded_data:
        if uploaded_data.name.endswith('.csv'):
            df = pd.read_csv(uploaded_data, dtype=str).fillna("")
        else:
            df = pd.read_excel(uploaded_data, dtype=str).fillna("")
        data_to_process = df.to_dict('records')
        st.success(f"✅ 已加载 {len(data_to_process)} 条表格数据")

# --- 第三步：生成设置 ---
st.markdown("---")
st.markdown("### 第三步：模板确认与生成")

# 模板选择逻辑
if os.path.exists(DEFAULT_TEMPLATE):
    template_option = st.radio("模板选择：", ["使用仓库内置模板", "上传本地新模板"], horizontal=True)
    if template_option == "使用仓库内置模板":
        template_path = DEFAULT_TEMPLATE
        st.caption(f"📍 当前已加载默认模板: {DEFAULT_TEMPLATE}")
    else:
        template_path = st.file_uploader("请上传自定义 Word 模板", type=["docx"])
else:
    st.warning("⚠️ 仓库未发现默认模板，请手动上传。")
    template_path = st.file_uploader("请上传 Word 模板", type=["docx"])

# --- 执行生成 ---
if template_path and data_to_process:
    if st.button("🚀 开始批量制作合并文档", use_container_width=True):
        try:
            master_doc = None
            progress_bar = st.progress(0)
            
            for i, row in enumerate(data_to_process):
                # 填充内容
                doc = DocxTemplate(template_path)
                context = {
                    'number': str(row.get('证书编号', '')).strip(),
                    'name': str(row.get('姓名', '')).strip(),
                    'id_card': str(row.get('身份证号', '')).strip(),
                    'date': str(row.get('培训日期', '')).strip(),
                    'standards': str(row.get('标准号', '')).strip()
                }
                doc.render(context)
                
                # 存入内存
                temp_io = io.BytesIO()
                doc.save(temp_io)
                temp_io.seek(0)
                
                # 文档合并
                current_doc = Document(temp_io)
                if master_doc is None:
                    master_doc = current_doc
                    composer = Composer(master_doc)
                else:
                    master_doc.add_page_break()
                    composer.append(current_doc)
                
                progress_bar.progress((i + 1) / len(data_to_process))

            # 下载
            output_io = io.BytesIO()
            master_doc.save(output_io)
            output_io.seek(0)
            
            st.balloons()
            st.download_button(
                label="🎁 制作完成！点击下载汇总文档 (.docx)",
                data=output_io.getvalue(),
                file_name="内审员证书汇总导出.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"制作失败，请检查数据格式或模板：{e}")
else:
    st.info("等待录入数据并确认模板...")
