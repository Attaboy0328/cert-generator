import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os
from docx import Document
from docxcompose.composer import Composer

# 页面配置
st.set_page_config(page_title="内审员证书批量生成", layout="centered")

# 标题
st.title("🎓 内审员证书一键生成工具")
st.write("只需三步，快速批量制作合并版证书 Word 文档。")

# --- 第一步：准备工作 (下载样例) ---
st.markdown("### 第一步：准备数据")
col1, col2 = st.columns([1, 1])

with col1:
    # 动态生成 Excel 样例
    example_data = {
        "姓名": ["张三", "李四"],
        "证书编号": ["T-2025-25-001", "T-2025-25-002"],
        "身份证号": ["'440683198811060001", "'440683198811060002"],
        "培训日期": ["2025年9月3-5日", "2025年9月3-5日"],
        "标准号": ["ISO9001", "ISO22000"]
    }
    df_sample = pd.DataFrame(example_data)
    output_sample = io.BytesIO()
    with pd.ExcelWriter(output_sample, engine='openpyxl') as writer:
        df_sample.to_excel(writer, index=False)
    
    st.download_button(
        label="📥 下载 Excel 数据填写样例",
        data=output_sample.getvalue(),
        file_name="证书数据样例.xlsx",
        help="点击下载标准格式表格，填好后再上传。"
    )

# --- 第二步：选择模板与数据 ---
st.markdown("---")
st.markdown("### 第二步：选择模板与上传数据")

DEFAULT_TEMPLATE = "内审员证书.docx"
uploaded_template = None

# 默认模板逻辑
if os.path.exists(DEFAULT_TEMPLATE):
    mode = st.radio("模板选择：", ["使用默认模板", "上传新模板"], horizontal=True)
    if mode == "使用默认模板":
        uploaded_template = DEFAULT_TEMPLATE
        st.success(f"✅ 已加载默认模板: {DEFAULT_TEMPLATE}")
    else:
        uploaded_template = st.file_uploader("请上传自定义 Word 模板", type=["docx"])
else:
    st.warning("⚠️ 仓库未发现默认模板，请手动上传。")
    uploaded_template = st.file_uploader("上传证书 Word 模板", type=["docx"])

# 上传 Excel 数据
uploaded_data = st.file_uploader("请上传填好的学员信息 (Excel/CSV)", type=["xlsx", "csv"])

# --- 第三步：生成与下载 ---
st.markdown("---")
st.markdown("### 第三步：开始批量制作")

if uploaded_template and uploaded_data:
    try:
        # 强制将所有数据读为字符串，彻底规避 'got integer' 报错
        if uploaded_data.name.endswith('.csv'):
            df = pd.read_csv(uploaded_data, dtype=str).fillna("")
        else:
            df = pd.read_excel(uploaded_data, dtype=str).fillna("")
        
        st.write(f"📊 已检测到 **{len(df)}** 位学员信息，点击下方按钮开始合并。")

        if st.button("🚀 生成全员合并版 Word", use_container_width=True):
            progress_bar = st.progress(0)
            master_doc = None
            
            for index, row in df.iterrows():
                # 渲染每一份证书
                doc = DocxTemplate(uploaded_template)
                
                # context 中的 key 必须对应 Word 里的 {{变量名}}
                context = {
                    'number': str(row.get('证书编号', '')),
                    'name': str(row.get('姓名', '')),
                    'id_card': str(row.get('身份证号', '')),
                    'date': str(row.get('培训日期', '')),
                    'standards': str(row.get('标准号', ''))
                }
                
                # 渲染
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
                    master_doc.add_page_break() # 分页
                    composer.append(current_doc)
                
                progress_bar.progress((index + 1) / len(df))

            # 导出下载
            if master_doc:
                output_io = io.BytesIO()
                master_doc.save(output_io)
                output_io.seek(0)
                
                st.balloons()
                st.download_button(
                    label="🎉 制作完成！点击下载结果文档",
                    data=output_io.getvalue(),
                    file_name="全员内审员证书汇总.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
    except Exception as e:
        st.error(f"❌ 运行出错：{e}")
        st.info("💡 温馨提示：请检查 Excel 表头是否完全对应：姓名、证书编号、身份证号、培训日期、标准号")
else:
    st.info("请先上传数据文件以启用生成按钮。")
