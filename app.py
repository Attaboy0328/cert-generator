import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile
import os
import subprocess
from docx import Document
from docxcompose.composer import Composer

st.set_page_config(page_title="内审员证书批量工具", layout="centered")

st.title("🎓 内审员证书极速生成器")
st.markdown("""
**优化说明：**
- 批量生成所有 Word 文档。
- 自动合并为一个大文件并转换。
- **只进行一次 PDF 转换，速度大幅提升！**
""")

uploaded_template = st.file_uploader("1. 上传 Word 模板", type=["docx"])
uploaded_data = st.file_uploader("2. 上传数据 (Excel/CSV)", type=["xlsx", "csv"])

if uploaded_template and uploaded_data:
    try:
        df = pd.read_csv(uploaded_data) if uploaded_data.name.endswith('.csv') else pd.read_excel(uploaded_data)
        st.success(f"已读取 {len(df)} 人信息")

        if st.button("🚀 开始批量制作"):
            progress_bar = st.progress(0)
            zip_buffer = io.BytesIO()
            
            # 用于合并的主文档
            master_doc = None
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                for index, row in df.iterrows():
                    # 填充 Word
                    doc = DocxTemplate(uploaded_template)
                    context = {
                        'number': str(row['证书编号']),
                        'name': str(row['姓名']),
                        'id_card': str(row['身份证号']),
                        'date': str(row['培训日期']),
                        'standards': str(row['标准号'])
                    }
                    doc.render(context)
                    
                    # 存入内存
                    word_io = io.BytesIO()
                    doc.save(word_io)
                    word_io.seek(0)
                    
                    # 添加到压缩包
                    zip_file.writestr(f"{row['姓名']}_证书.docx", word_io.getvalue())
                    
                    # --- 合并逻辑 ---
                    current_doc = Document(word_io)
                    if master_doc is None:
                        master_doc = current_doc
                        composer = Composer(master_doc)
                    else:
                        master_doc.add_page_break() # 每个人的证书占一页
                        composer.append(current_doc)
                    
                    progress_bar.progress((index + 1) / len(df))

                # --- 核心优化：全员转换 PDF ---
                st.write("正在执行全员 PDF 转换，请稍后（仅需几秒）...")
                merged_word_path = "all_in_one.docx"
                master_doc.save(merged_word_path)
                
                # 调用 LibreOffice 执行单次转换
                subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', merged_word_path], check=True)
                
                # 将合并后的 Word 和 PDF 都存入压缩包
                if os.path.exists("all_in_one.pdf"):
                    with open("all_in_one.pdf", "rb") as f:
                        zip_file.writestr("【全员汇总】所有证书合并版.pdf", f.read())
                    os.remove("all_in_one.pdf")
                
                with open(merged_word_path, "rb") as f:
                    zip_file.writestr("【全员汇总】所有证书合并版.docx", f.read())
                os.remove(merged_word_path)

            st.balloons()
            st.download_button(
                label="🎁 下载全部结果 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="批量证书导出.zip",
                mime="application/x-zip-compressed"
            )
    except Exception as e:
        st.error(f"处理出错：{e}")
