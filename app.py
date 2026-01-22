import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile
import os
import subprocess
from docx import Document
from docxcompose.composer import Composer

st.set_page_config(page_title="内审员证书批量生成器", layout="centered")

st.title("🎓 内审员证书批量制作工具")
st.info("功能更新：现在下载的压缩包内会额外包含一个【全员合并版.pdf】")

uploaded_template = st.file_uploader("第一步：上传 Word 证书模板", type=["docx"])
uploaded_data = st.file_uploader("第二步：上传学员信息 Excel 或 CSV", type=["xlsx", "csv"])

if uploaded_template and uploaded_data:
    try:
        if uploaded_data.name.endswith('.csv'):
            df = pd.read_csv(uploaded_data)
        else:
            df = pd.read_excel(uploaded_data)
        
        st.success(f"✅ 成功读取到 {len(df)} 条学员数据！")

        if st.button("第三步：一键生成并导出"):
            zip_buffer = io.BytesIO()
            master_doc = None  # 用于存放合并的大文档
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for index, row in df.iterrows():
                    # 1. 生成单份 Word
                    doc = DocxTemplate(uploaded_template)
                    context = {
                        'number': str(row['证书编号']),
                        'name': str(row['姓名']),
                        'id_card': str(row['身份证号']),
                        'date': str(row['培训日期']),
                        'standards': str(row['标准号'])
                    }
                    doc.render(context)
                    
                    # 保存单份到内存
                    single_docx_io = io.BytesIO()
                    doc.save(single_docx_io)
                    single_docx_io.seek(0)
                    
                    # 放入压缩包
                    word_name = f"{row['姓名']}_内审员证书.docx"
                    zip_file.writestr(word_name, single_docx_io.getvalue())
                    
                    # 2. 合并逻辑
                    current_doc = Document(single_docx_io)
                    if master_doc is None:
                        master_doc = current_doc
                        composer = Composer(master_doc)
                    else:
                        # 在合并前增加一个分页符
                        master_doc.add_page_break()
                        composer.append(current_doc)

                # 3. 处理合并后的 PDF
                if master_doc:
                    st.write("正在准备合并版 PDF，请稍候...")
                    # 先存为临时 Word 文件
                    temp_word = "all_certs.docx"
                    master_doc.save(temp_word)
                    
                    # 调用服务器的 LibreOffice 进行转换
                    try:
                        subprocess.run([
                            'libreoffice', '--headless', '--convert-to', 'pdf', temp_word
                        ], check=True)
                        
                        pdf_file_name = "all_certs.pdf"
                        if os.path.exists(pdf_file_name):
                            with open(pdf_file_name, "rb") as f:
                                zip_file.writestr("【重要】全员证书合并版.pdf", f.read())
                            os.remove(pdf_file_name) # 清理
                        os.remove(temp_word) # 清理
                    except Exception as e:
                        st.warning(f"PDF 合并失败（可能服务器环境限制）：{e}")

            st.download_button(
                label="🚀 点击下载生成的压缩包 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="内审员证书批量结果.zip",
                mime="application/x-zip-compressed"
            )
    except Exception as e:
        st.error(f"发生错误：{e}")
