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

# --- 第一步：选择录入模式 ---
st.markdown("### 第一步：选择录入模式")
mode = st.radio(
    "选择方式：", 
    ["Excel 文件上传", "网页表格填写 (支持粘贴)"], 
    index=0, 
    horizontal=True
)

DEFAULT_TEMPLATE = "内审员证书.docx"
data_to_process = []

# --- 第二步：准备数据 ---
st.markdown("---")
st.markdown("### 第二步：填写或上传信息")

if mode == "网页表格填写 (支持粘贴)":
    st.info("💡 提示：点击左上角第一个单元格（证书编号下方）并按下 Ctrl+V 即可粘贴 Excel 数据。")
    
    init_df = pd.DataFrame(
        {
            "序号": [i for i in range(1, 101)],
            "证书编号": [None] * 100,
            "姓名": [None] * 100,
            "身份证号": [None] * 100,
            "培训日期": [None] * 100,
            "标准号": [None] * 100,
        }
    )
    
    edited_df = st.data_editor(
        init_df,
        num_rows="fixed", 
        use_container_width=True,
        hide_index=True, 
        height=380,      
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
    data_to_process = [
        {k: str(v).strip() for k, v in row.items() if v is not None} 
        for row in data_to_process if any(row.values())
    ]

else:
    # --- 第三步优化：带案例的模板下载 ---
    col1, col2 = st.columns([2, 3])
    with col1:
        # 创建带案例的示例数据
        example_data = {
            "证书编号": ["T-2025-001 (示例)"],
            "姓名": ["张三 (示例)"],
            "身份证号": ["440683199001010001"],
            "培训日期": ["2025年9月3-5日"],
            "标准号": ["ISO9001:2015、ISO22000:2018"]
        }
        template_df = pd.DataFrame(example_data)
        template_buffer = io.BytesIO()
        with pd.ExcelWriter(template_buffer, engine='openpyxl') as writer:
            template_df.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 下载带案例的 Excel 模板",
            data=template_buffer.getvalue(),
            file_name="学员信息上传模板(含案例).xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    
    with col2:
        uploaded_data = st.file_uploader("上传学员信息文件", type=["xlsx", "csv"], label_visibility="collapsed")

    if uploaded_data:
        if uploaded_data.name.endswith('.csv'):
            df = pd.read_csv(uploaded_data, dtype=str).fillna("")
        else:
            df = pd.read_excel(uploaded_data, dtype=str).fillna("")
        
        # 核心逻辑：自动过滤掉带“(示例)”字样的行
        full_data = df.to_dict('records')
        data_to_process = [
            row for row in full_data 
            if "(示例)" not in str(row.get('姓名', '')) and "(示例)" not in str(row.get('证书编号', ''))
        ]
        
        if len(data_to_process) > 0:
            st.success(f"✅ 已成功加载 {len(data_to_process)} 条有效数据 (已自动排除示例行)")

# --- 第三步：模板确认与生成 ---
st.markdown("---")
st.markdown("### 第三步：模板确认与生成")

if os.path.exists(DEFAULT_TEMPLATE):
    template_option = st.radio("证书 Word 模板：", ["使用内置模板", "上传本地新模板"], horizontal=True)
    if template_option == "使用内置模板":
        template_path = DEFAULT_TEMPLATE
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
            
            valid_count = 0
            for i, row in enumerate(data_to_process):
                # 清洗数据
                name_val = str(row.get('姓名', '')).replace('nan', '').strip()
                if not name_val or name_val == "":
                    continue
                
                valid_count += 1
                doc = DocxTemplate(template_path)
                context = {
                    'number': str(row.get('证书编号', '')).replace('nan', '').strip(),
                    'name': name_val,
                    'id_card': str(row.get('身份证号', '')).replace('nan', '').strip(),
                    'date': str(row.get('培训日期', '')).replace('nan', '').strip(),
                    'standards': str(row.get('标准号', '')).replace('nan', '').strip()
                }
                doc.render(context)
                
                temp_io = io.BytesIO()
                doc.save(temp_io)
                temp_io.seek(0)
                
                current_doc = Document(temp_io)
                if master_doc is None:
                    master_doc = current_doc
                    composer = Composer(master_doc)
                else:
                    master_doc.add_page_break()
                    composer.append(current_doc)
                
                progress_bar.progress((i + 1) / len(data_to_process))

            if master_doc and valid_count > 0:
                output_io = io.BytesIO()
                master_doc.save(output_io)
                output_io.seek(0)
                
                st.balloons()
                st.download_button(
                    label=f"🎁 制作完成({valid_count}份)！点击下载汇总文档",
                    data=output_io.getvalue(),
                    file_name="证书汇总导出.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
        except Exception as e:
            st.error(f"制作失败：{e}")
else:
    st.info("等待录入数据并确认模板...")
