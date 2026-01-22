import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile

# 页面配置
st.set_page_config(page_title="内审员证书生成器", layout="centered")

st.title("🎓 内审员证书批量制作工具")
st.info("请确保 Excel 的表头包含：证书编号、姓名、身份证号、培训日期、标准号")

# 1. 上传文件
uploaded_template = st.file_uploader("第一步：上传 Word 证书模板", type=["docx"])
uploaded_data = st.file_uploader("第二步：上传学员信息 Excel 或 CSV", type=["xlsx", "csv"])

if uploaded_template and uploaded_data:
    # 读取数据逻辑
    try:
        if uploaded_data.name.endswith('.csv'):
            df = pd.read_csv(uploaded_data)
        else:
            df = pd.read_excel(uploaded_data)
        
        st.success(f"✅ 成功读取到 {len(df)} 条学员数据！")

        # 2. 生成按钮
        if st.button("第三步：开始生成并打包下载"):
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for index, row in df.iterrows():
                    # 每次循环都重新读取模板
                    doc = DocxTemplate(uploaded_template)
                    
                    # 填充内容（对应 Word 模板中的 {{变量名}}）
                    context = {
                        'number': str(row['证书编号']),
                        'name': str(row['姓名']),
                        'id_card': str(row['身份证号']),
                        'date': str(row['培训日期']),
                        'standards': str(row['标准号'])
                    }
                    
                    doc.render(context)
                    
                    # 保存到内存
                    out_docx = io.BytesIO()
                    doc.save(out_docx)
                    out_docx.seek(0)
                    
                    # 放入压缩包
                    file_name = f"{row['姓名']}_内审员证书.docx"
                    zip_file.writestr(file_name, out_docx.getvalue())
            
            # 3. 下载按钮
            st.download_button(
                label="🚀 点击下载生成的压缩包 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="批量证书导出.zip",
                mime="application/x-zip-compressed"
            )
    except Exception as e:
        st.error(f"发生错误：{e}")
