import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os
from docx import Document
from docxcompose.composer import Composer

# 页面配置
st.set_page_config(page_title="证书合并生成器", layout="centered")

st.title("🎓 内审员证书一键生成工具")

# --- 核心修改：模板选择逻辑 ---
DEFAULT_TEMPLATE = "内审员证书.docx" # 这里填写你上传到 GitHub 的模板文件名

st.sidebar.header("设置")
use_default = False
if os.path.exists(DEFAULT_TEMPLATE):
    use_default = st.sidebar.checkbox("使用仓库默认模板", value=True)
    if use_default:
        st.sidebar.success(f"已加载默认模板: {DEFAULT_TEMPLATE}")
else:
    st.sidebar.warning("仓库中未发现默认模板，请手动上传")

# 1. 文件上传
if not use_default:
    uploaded_template = st.file_uploader("第一步：上传证书 Word 模板", type=["docx"])
else:
    uploaded_template = DEFAULT_TEMPLATE

uploaded_data = st.file_uploader("第二步：上传学员信息 Excel 或 CSV", type=["xlsx", "csv"])

if (uploaded_template) and uploaded_data:
    try:
        # 读取数据逻辑
        if uploaded_data.name.endswith('.csv'):
            df = pd.read_csv(uploaded_data)
        else:
            df = pd.read_excel(uploaded_data)
        
        st.success(f"✅ 成功读取到 {len(df)} 条学员数据！")

        # 2. 生成按钮
        if st.button("第三步：一键生成合并版 Word"):
            progress_bar = st.progress(0)
            master_doc = None
            
            # 模板来源判断
            # 如果是上传的文件，需要通过 io.BytesIO 读取；如果是默认文件，直接传路径
            source_template = uploaded_template if use_default else uploaded_template
            
            for index, row in df.iterrows():
                # 填充单份证书
                doc = DocxTemplate(source_template)
                context = {
                    'number': str(row['证书编号']),
                    'name': str(row['姓名']),
                    'id_card': str(row['身份证号']),
                    'date': str(row['培训日期']),
                    'standards': str(row['标准号'])
                }
                doc.render(context)
                
                # 将单份存入临时内存
                temp_io = io.BytesIO()
                doc.save(temp_io)
                temp_io.seek(0)
                
                # 合并逻辑
                current_doc = Document(temp_io)
                if master_doc is None:
                    master_doc = current_doc
                    composer = Composer(master_doc)
                else:
                    master_doc.add_page_break()
                    composer.append(current_doc)
                
                progress_bar.progress((index + 1) / len(df))

            # 3. 提供下载
            if master_doc:
                output_io = io.BytesIO()
                master_doc.save(output_io)
                output_io.seek(0)
                
                st.balloons()
                st.download_button(
                    label="🎉 点击下载【全员合并版证书】.docx",
                    data=output_io.getvalue(),
                    file_name="全员内审员证书汇总.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
    except Exception as e:
        st.error(f"发生错误：{e}")
