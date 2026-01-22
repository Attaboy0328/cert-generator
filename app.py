import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile

st.set_page_config(page_title="内审员证书自动生成器", layout="centered")

st.title("🎓 内审员证书批量制作工具")
st.write("上传模板和数据，一键生成所有证书 Word 文档（打包下载）")

# 1. 上传文件
uploaded_template = st.file_uploader("第一步：上传 Word 证书模板", type=["docx"])
uploaded_data = st.file_uploader("第二步：上传学员信息 Excel/CSV", type=["xlsx", "csv"])

if uploaded_template and uploaded_data:
    # 读取数据
    if uploaded_data.name.endswith('.csv'):
        df = pd.read_csv(uploaded_data)
    else:
        df = pd.read_excel(uploaded_data)
    
    st.success(f"成功读取到 {len(df)} 条学员数据！")
    
    # 2. 点击生成
    if st.button("第三步：开始生成并打包"):
        # 创建一个内存中的 ZIP 文件
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for index, row in df.iterrows():
                # 渲染 Word
                doc = DocxTemplate(uploaded_template)
                # 这里的 key 对应 Word 模板里的 {{变量名}}
                context = {
                    [cite_start]'number': row['证书编号'],  # [cite: 1]
                    [cite_start]'name': row['姓名'],      # [cite: 2]
                    [cite_start]'id_card': row['身份证号'], # [cite: 3]
                    [cite_start]'date': row['培训日期'],    # [cite: 4]
                    [cite_start]'standards': row['标准号']  # [cite: 4]
                }
                doc.render(context)
                
                # 将生成的 Word 存入内存
                out_docx = io.BytesIO()
                doc.save(out_docx)
                out_docx.seek(0)
                
                # 添加到 ZIP 压缩包
                file_name = f"{row['姓名']}_内审员证书.docx"
                zip_file.writestr(file_name, out_docx.getvalue())
        
        # 3. 提供下载
        st.download_button(
            label="🎉 点击下载所有证书 (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="批量证书生成结果.zip",
            mime="application/x-zip-compressed"
        )