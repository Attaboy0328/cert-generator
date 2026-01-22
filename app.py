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

# --- 第一步：选择录入模式 (默认设置为 Excel 上传) ---
st.markdown("### 第一步：选择录入模式")
mode = st.radio(
    "选择方式：", 
    ["Excel 文件上传", "网页表格填写 (支持粘贴)"], 
    index=0, # 默认索引为 0，即 Excel 文件上传
    horizontal=True
)

DEFAULT_TEMPLATE = "内审员证书.docx"
data_to_process = []

# --- 第二步：准备数据 ---
st.markdown("---")
st.markdown("### 第二步：填写或上传信息")

if mode == "网页表格填写 (支持粘贴)":
    st.info("💡 提示：点击左上角第一个单元格并按下 Ctrl+V 即可粘贴 Excel 数据。")
    
    # 创建 100 行初始数据，并设置序号从 1 开始
    # 我们用一个专门的列来存序号，方便显示
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
    
    # 使用数据编辑器
    # height=380 左右通常能完整显示表头 + 9行数据
    edited_df = st.data_editor(
        init_df,
        num_rows="fixed", 
        use_container_width=True,
        hide_index=True, # 隐藏 pandas 原生的 0 开始的索引
        height=380,      # 锁定高度，前9行左右可见，之后滚动
        column_config={
            "序号": st.column_config.NumberColumn("序号", width=40, disabled=True),
            "证书编号": st.column_config.TextColumn("证书编号", width="small"),
            "姓名": st.column_config.TextColumn("姓名", width="small"),
            "身份证号": st.column_config.TextColumn("身份证号", width="medium"),
            "培训日期": st.column_config.TextColumn("培训日期", width="medium"),
            "标准号": st.column_config.TextColumn("标准号", width="large"),
        }
    )
    
    # 提取有效数据：过滤掉所有业务字段都为空的行
    temp_df = edited_df.drop(columns=["序号"])
    data_to_process = temp_df.dropna(how='all').to_dict('records')
    # 进一步清洗：去除 None 和 空字符串
    data_to_process = [
        {k: str(v).strip() for k, v in row.items() if v is not None} 
        for row in data_to_process if any(row.values())
    ]

else:
    uploaded_data = st.file_uploader("上传学员信息 Excel 文件", type=["xlsx", "csv"])
    if uploaded_data:
        if uploaded_data.name.endswith('.csv'):
            df = pd.read_csv(uploaded_data, dtype=str).fillna("")
        else:
            df = pd.read_excel(uploaded_data, dtype=str).fillna("")
        data_to_process = df.to_dict('records')
        st.success(f"✅ 已加载 {len(data_to_process)} 条表格数据")

# --- 第三步：模板确认与生成 ---
st.markdown("---")
st.markdown("### 第三步：模板确认与生成")

if os.path.exists(DEFAULT_TEMPLATE):
    template_option = st.radio("模板选择：", ["使用内置模板", "上传本地新模板"], horizontal=True)
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
            
            # 清洗最终要填入模板的数据，确保没有 "None" 字符串
            valid_count = 0
            for i, row in enumerate(data_to_process):
                # 检查是否是真的有数据（比如至少有姓名）
                if not row.get('姓名') or row.get('姓名') == 'nan':
                    continue
                
                valid_count += 1
                doc = DocxTemplate(template_path)
                context = {
                    'number': str(row.get('证书编号', '')).replace('nan', '').strip(),
                    'name': str(row.get('姓名', '')).replace('nan', '').strip(),
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
            else:
                st.warning("未检测到有效数据，请检查表格内容。")
        except Exception as e:
            # 捕获模板错误并给出友好提示
            error_msg = str(e)
            if "expected token" in error_msg:
                st.error("❌ 制作失败：检测到 Word 模板语法错误。")
                st.info("💡 解决方案：请检查模板中的 {{变量名}} 是否写成了具体数字。模板中只能写英文变量名，如 {{ name }}。")
            else:
                st.error(f"❌ 制作失败：{error_msg}")
else:
    st.info("等待录入数据并确认模板...")
