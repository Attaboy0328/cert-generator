import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os
from docx import Document
from docxcompose.composer import Composer
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import streamlit.components.v1 as components

# --- 1. 页面基础配置 ---
st.set_page_config(page_title="内审员证书智能化工具", layout="centered")

# --- 2. 深度界面定制 (CSS & JS) ---
def apply_custom_interface():
    # CSS 注入：隐藏按钮、标题居中、间距优化、动画效果
    st.markdown("""
        <style>
        /* 隐藏右上角官方 Share 按钮和收藏图标 */
        div[data-testid="stStatusWidget"] { display: none !important; }
        .st-emotion-cache-15ec60u, .st-emotion-cache-zq59db { display: none !important; }
        
        /* 隐藏右下角内容 */
        footer { visibility: hidden !important; }

        /* 标题居中并适配主题颜色 */
        .stApp h1 {
            text-align: center !important;
            display: block;
            margin-left: auto;
            margin-right: auto;
            width: 100%;
            font-weight: 700;
            /* 关键修复：使用继承颜色，确保 Dark/Light 模式下都清晰 */
            color: inherit !important; 
            margin-bottom: 55px !important; 
            padding-top: 10px;
        }
        
        /* 页面切换自然过渡动画 */
        .main .block-container {
            animation: fadeIn 0.6s cubic-bezier(0.4, 0, 0.2, 1);
        }
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(8px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        /* 移动端间距适配 */
        @media (max-width: 640px) {
            .stApp h1 { 
                font-size: 1.8rem !important;
                margin-bottom: 40px !important; 
            }
        }
        
        /* 页脚留白 */
        .main .block-container { padding-bottom: 100px; }
        </style>
    """, unsafe_allow_html=True)

    # JS 注入：在 GitHub 左侧添加 Streamlit Logo 导航
    components.html("""
        <script>
        const targetUrl = "https://share.streamlit.io/user/attaboy0328";
        const logoSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 24 24" fill="#FF4B4B" style="margin-right:15px; cursor:pointer;"><path d="M12 2L2 19.72L12 22L22 19.72L12 2ZM12 16.5L6.5 15.5L12 6L17.5 15.5L12 16.5Z"/></svg>`;
        function injectLogo() {
            const header = window.parent.document.querySelector('header[data-testid="stHeader"]');
            const container = header ? header.querySelector('div:nth-child(2)') : null;
            if (container && !window.parent.document.getElementById('custom-streamlit-logo')) {
                const link = window.parent.document.createElement('a');
                link.id = 'custom-streamlit-logo';
                link.href = targetUrl; link.target = "_blank";
                link.innerHTML = logoSvg; link.style.display = "flex"; link.style.alignItems = "center";
                container.prepend(link);
            }
        }
        setInterval(injectLogo, 500);
        </script>
    """, height=0)

apply_custom_interface()

# --- 3. 业务内容 ---
st.title("🎓 内审员证书智能化工具")

# 第一步
st.markdown("### 第一步：选择录入模式")
mode = st.radio("选择方式：", ["Excel 文件上传", "网页表格填写 (支持粘贴)"], index=0, horizontal=True)

DEFAULT_TEMPLATE = "内审员证书.docx"
data_to_process = []

# 第二步
st.markdown("---")
st.markdown("### 第二步：填写或上传信息")

if mode == "网页表格填写 (支持粘贴)":
    st.info("💡 提示：点击左上角第一个单元格并按下 Ctrl+V 即可粘贴 Excel 数据。")
    init_df = pd.DataFrame({
        "序号": [i for i in range(1, 101)],
        "证书编号": [None] * 100, "姓名": [None] * 100, "身份证号": [None] * 100, "培训日期": [None] * 100, "标准号": [None] * 100,
    })
    edited_df = st.data_editor(
        init_df, num_rows="fixed", use_container_width=True, hide_index=True, height=380,
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
    data_to_process = [{k: str(v).strip() for k, v in row.items() if v is not None} for row in data_to_process if any(row.values())]

else:
    col1, col2 = st.columns([2, 3])
    with col1:
        example_data = {"证书编号": ["T-2025-001"],"姓名": ["张三"],"身份证号": ["440683..."],"培训日期": ["2025年9月"],"标准号": ["ISO9001"]}
        df_ex = pd.DataFrame(example_data)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            df_ex.to_excel(writer, index=False)
            ws = writer.sheets['Sheet1']
            for i in range(1, 6): ws.column_dimensions[get_column_letter(i)].width = 20
            for cell in ws[2]: cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        st.download_button("📥 下载标准模板", data=buf.getvalue(), file_name="学员信息上传模板.xlsx")
    with col2:
        up = st.file_uploader("上传文件", type=["xlsx", "csv"], label_visibility="collapsed")
        if up:
            df = pd.read_csv(up, dtype=str).fillna("") if up.name.endswith('.csv') else pd.read_excel(up, dtype=str).fillna("")
            data_to_process = [row for row in df.to_dict('records') if "示例" not in str(row.get('姓名', ''))]
            if data_to_process: st.success(f"✅ 已加载 {len(data_to_process)} 条数据")

# 第三步
st.markdown("---")
st.markdown("### 第三步：模板确认与生成")
if os.path.exists(DEFAULT_TEMPLATE):
    t_opt = st.radio("模板：", ["使用内置模板", "上传本地新模板"], horizontal=True)
    t_path = DEFAULT_TEMPLATE if t_opt == "使用内置模板" else st.file_uploader("上传 docx 模板", type=["docx"])
else:
    t_path = st.file_uploader("上传 Word 模板", type=["docx"])

if t_path and data_to_process:
    if st.button("🚀 启动批量制作", use_container_width=True):
        try:
            master, bar, count = None, st.progress(0), 0
            # 遍历数据进行处理
            for i, row in enumerate(data_to_process):
                name_v = str(row.get('姓名', '')).strip()
                if not name_v or name_v == 'nan':
                    continue
                
                count += 1
                doc = DocxTemplate(t_path)
                # 渲染 Word 模板
                doc.render({
                    'number': str(row.get('证书编号','')).strip(),
                    'name': name_v,
                    'id_card': str(row.get('身份证号','')).strip(),
                    'date': str(row.get('培训日期','')).strip(),
                    'standards': str(row.get('标准号','')).strip()
                })
                
                tmp = io.BytesIO()
                doc.save(tmp)
                tmp.seek(0)
                cur = Document(tmp)
                
                if master is None:
                    master = cur
                    composer = Composer(master)
                else:
                    master.add_page_break()
                    composer.append(cur)
                
                # 更新进度条
                bar.progress((i + 1) / len(data_to_process))
            
            # 循环结束后，检查是否有生成成功的文件
            if master:
                out = io.BytesIO()
                master.save(out)
                out.seek(0)
                st.balloons()
                st.download_button(
                    label=f"🎁 下载汇总文档({count}份)", 
                    data=out.getvalue(), 
                    file_name="证书汇总.docx", 
                    use_container_width=True
                )
        
        except Exception as e:
            # 必须包含这个 except 块，否则会报你遇到的那个错误
            st.error(f"制作过程中发生错误：{e}")
