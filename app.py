import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os
from docx import Document
from docxcompose.composer import Composer

# 页面配置
st.set_page_config(page_title="证书智能制作工具", layout="wide")

st.title("🎓 内审员证书智能制作工具")

# --- 第一步：模式选择 ---
st.markdown("### 第一步：选择制作模式")
mode = st.radio("选择模式：", ["手动填写 (支持从 Excel 复制粘贴)", "Excel 文件上传"], horizontal=True)

DEFAULT_TEMPLATE = "内审员证书.docx"
data_to_process = []

# --- 第二步：数据准备 ---
st.markdown("---")

if mode == "手动填写 (支持从 Excel 复制粘贴)":
    col_input, col_preview = st.columns([1, 1])
    
    with col_input:
        st.markdown("### ✍️ 数据录入")
        st.info("💡 技巧：您可以直接从 Excel 选中多行多列并复制，然后粘贴到下方。最多支持 100 份。")
        raw_text = st.text_area(
            "粘贴区域 (格式：姓名 证书编号 身份证号 培训日期 标准号)", 
            placeholder="张三\tT-2025-01\t4406...\t2025年9月\tISO9001",
            height=300
        )
        
        if raw_text:
            lines = raw_text.strip().split('\n')[:100]
            for line in lines:
                # 处理 Excel 的 Tab 分隔符
                parts = line.split('\t')
                if len(parts) >= 2:
                    data_to_process.append({
                        '姓名': parts[0].strip(),
                        '证书编号': parts[1].strip() if len(parts) > 1 else "",
                        '身份证号': parts[2].strip() if len(parts) > 2 else "",
                        '培训日期': parts[3].strip() if len(parts) > 3 else "",
                        '标准号': parts[4].strip() if len(parts) > 4 else ""
                    })
            
            if data_to_process:
                st.success(f"✅ 已识别 {len(data_to_process)} 条数据")

    with col_preview:
        st.markdown("### 👁️ 实时内容预览 (第一份)")
        if data_to_process:
            p = data_to_process[0]
            # 使用 Markdown 模拟一个简单的证书预览样式
            st.markdown(f"""
            <div style="border: 2px solid #555; padding: 20px; border-radius: 10px; background-color: #f9f9f9; color: #333; font-family: sans-serif;">
                <h4 style="text-align: center; color: #d32f2f;">内审员证书预览</h4>
                <hr>
                <p><b>证书编号：</b>{p['证书编号']}</p>
                <p><b>姓名：</b>{p['姓名']}</p>
                <p><b>身份证号：</b>{p['身份证号']}</p>
                <p><b>培训日期：</b>{p['培训日期']}</p>
                <p><b>标准号：</b><br>{p['标准号']}</p>
                <hr>
                <p style="font-size: 0.8em; color: #888;">* 实际生成的排版将严格遵循 Word 模板格式</p>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.warning("暂无数据，请在左侧输入或粘贴内容。")

else:
    st.markdown("### 📂 批量文件上传")
    uploaded_data = st.file_uploader("上传学员信息 (Excel/CSV)", type=["xlsx", "csv"])
    if uploaded_data:
        if uploaded_data.name.endswith('.csv'):
            df = pd.read_csv(uploaded_data, dtype=str).fillna("")
        else:
            df = pd.read_excel(uploaded_data, dtype=str).fillna("")
        data_to_process = df.to_dict('records')
        st.success(f"✅ 已加载 {len(data_to_process)} 条表格数据")

# --- 第三步：一键生成 ---
st.markdown("---")
st.markdown("### 第三步：生成与下载")

# 检查默认模板
if os.path.exists(DEFAULT_TEMPLATE):
    template_path = DEFAULT_TEMPLATE
    st.caption(f"📍 当前使用默认模板: {DEFAULT_TEMPLATE}")
else:
    template_path = st.file_uploader("⚠️ 未发现默认模板，请手动上传 Word 模板", type=["docx"])

if template_path and data_to_process:
    if st.button("🚀 开始批量制作合并文档", use_container_width=True):
        try:
            master_doc = None
            progress_bar = st.progress(0)
            
            for i, row in enumerate(data_to_process):
                # 填充
                doc = DocxTemplate(template_path)
                context = {
                    'number': str(row.get('证书编号', '')),
                    'name': str(row.get('姓名', '')),
                    'id_card': str(row.get('身份证号', '')),
                    'date': str(row.get('培训日期', '')),
                    'standards': str(row.get('标准号', ''))
                }
                doc.render(context)
                
                # 存入内存
                temp_io = io.BytesIO()
                doc.save(temp_io)
                temp_io.seek(0)
                
                # 合并
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
                label="🎁 点击下载汇总文档 (.docx)",
                data=output_io.getvalue(),
                file_name="内审员证书汇总导出.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"制作失败：{e}")
else:
    st.info("待处理数据为空，请先完成录入或上传。")
