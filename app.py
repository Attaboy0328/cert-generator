import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile
import os
import subprocess
import shutil
from docx import Document
from docxcompose.composer import Composer

# --- 1. 字体安装逻辑 (解决PDF乱码) ---
def install_custom_fonts():
    # 定义服务器存放字体的路径
    target_font_dir = os.path.expanduser("~/.local/share/fonts")
    if not os.path.exists(target_font_dir):
        os.makedirs(target_font_dir)
    
    # 查找仓库里上传的字体文件 (ttf, ttc, otf)
    font_files = [f for f in os.listdir('.') if f.lower().endswith(('.ttf', '.ttc', '.otf'))]
    
    if font_files:
        for font in font_files:
            target_path = os.path.join(target_font_dir, font)
            if not os.path.exists(target_path):
                shutil.copy(font, target_path)
        
        # 刷新 Linux 字体缓存
        try:
            subprocess.run(["fc-cache", "-fv"], check=True)
            return True
        except:
            return False
    return False

# 尝试安装字体
font_installed = install_custom_fonts()

# --- 2. 页面设置 ---
st.set_page_config(page_title="内审员证书极速生成器", layout="centered")

st.title("🎓 内审员证书一键生成工具")
st.markdown("""
### 使用说明：
1. **Word 模板**：请确保包含 `{{number}}`, `{{name}}`, `{{id_card}}`, `{{date}}`, `{{standards}}` 占位符。
2. **字体解决**：若 PDF 格式不对，请将 `.ttf` 字体文件上传至 GitHub 仓库根目录。
""")

if font_installed:
    st.caption("✅ 已加载自定义字体，PDF 转换质量已优化")
else:
    st.caption("⚠️ 未检测到自定义字体文件，PDF 可能出现排版偏移")

# --- 3. 文件上传 ---
uploaded_template = st.file_uploader("1. 上传证书 Word 模板", type=["docx"])
uploaded_data = st.file_uploader("2. 上传学员信息 (Excel/CSV)", type=["xlsx", "csv"])

if uploaded_template and uploaded_data:
    try:
        # 读取数据
        if uploaded_data.name.endswith('.csv'):
            df = pd.read_csv(uploaded_data)
        else:
            df = pd.read_excel(uploaded_data)
            
        st.success(f"已成功识别 {len(df)} 位学员信息")

        # --- 4. 核心生成逻辑 ---
        if st.button("🚀 开始极速制作 (Word + 合并版PDF)"):
            progress_bar = st.progress(0)
            zip_buffer = io.BytesIO()
            master_doc = None  # 用于合并的主文档
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                for index, row in df.iterrows():
                    # 4.1 生成单份 Word
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
                    
                    # 放入压缩包
                    zip_file.writestr(f"{row['姓名']}_证书.docx", word_io.getvalue())
                    
                    # 4.2 准备合并
                    current_doc = Document(word_io)
                    if master_doc is None:
                        master_doc = current_doc
                        composer = Composer(master_doc)
                    else:
                        master_doc.add_page_break() # 每个人占一页
                        composer.append(current_doc)
                    
                    progress_bar.progress((index + 1) / len(df))

                # 4.3 执行单次 PDF 转换（大幅提速）
                st.write("正在执行全员 PDF 转换，请稍候...")
                temp_word_name = "temp_all_certs.docx"
                master_doc.save(temp_word_name)
                
                # 调用服务器 LibreOffice
                try:
                    subprocess.run([
                        'libreoffice', '--headless', '--convert-to', 'pdf', temp_word_name
                    ], check=True)
                    
                    pdf_name = "temp_all_certs.pdf"
                    if os.path.exists(pdf_name):
                        with open(pdf_name, "rb") as f:
                            zip_file.writestr("【全员汇总】所有证书合并版.pdf", f.read())
                        os.remove(pdf_name)
                    
                    with open(temp_word_name, "rb") as f:
                        zip_file.writestr("【全员汇总】所有证书合并版.docx", f.read())
                    os.remove(temp_word_name)
                except Exception as pdf_err:
                    st.error(f"PDF 转换失败，原因：{pdf_err}")

            st.balloons()
            st.download_button(
                label="🎁 点击下载全部证书结果 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="内审员证书批量制作结果.zip",
                mime="application/x-zip-compressed"
            )

    except Exception as e:
        st.error(f"处理过程中发生错误：{e}")
