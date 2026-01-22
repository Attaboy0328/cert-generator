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
st.set_page_config(page_title="证书智能制作工具", layout="centered")

# --- 2. 深度界面定制 (主题切换、居中、隐藏官方按钮) ---
def apply_advanced_customizations():
    # A. 注入 CSS：标题居中、移动端适配、隐藏 Share 按钮
    st.markdown("""
        <style>
        div[data-testid="stStatusWidget"] { display: none !important; }
        footer { visibility: hidden !important; }
        
        /* 标题居中 */
        .stApp h1 {
            text-align: center !important;
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 10px;
            width: 100%;
            margin-top: 0px;
        }

        /* 移动端优化 */
        @media (max-width: 640px) {
            .stApp h1 { font-size: 1.6rem !important; }
            .stApp .block-container { padding: 1rem !important; }
        }
        .main .block-container { padding-bottom: 100px; }
        </style>
    """, unsafe_allow_html=True)

    # B. 注入 JS：实现右上角 Streamlit Logo 跳转
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
                link.href = targetUrl;
                link.target = "_blank";
                link.innerHTML = logoSvg;
                link.style.display = "flex";
                link.style.alignItems = "center";
                container.prepend(link);
            }
        }
        setInterval(injectLogo, 500);
        </script>
    """, height=0)

apply_advanced_customizations()

# --- 3. 主题切换逻辑 ---
# 使用 st.toggle 作为切换开关
theme_col1, theme_col2 = st.columns([8, 2])
with theme_col2:
    is_dark = st.toggle("🌙 夜间模式", value=False)

# 通过 JS 动态更改全局主题颜色变量
if is_dark:
    components.html("""
        <script>
            const doc = window.parent.document;
            doc.documentElement.style.setProperty('--primary-color', '#FF4B4B');
            doc.body.style.backgroundColor = '#0E1117';
            doc.querySelectorAll('.stApp').forEach(el => {
                el.style.backgroundColor = '#0E1117';
                el.style.color = '#FAFAFA';
            });
        </script>
    """, height=0)
else:
    components.html("""
        <script>
            const doc = window.parent.document;
            doc.body.style.backgroundColor = '#FFFFFF';
            doc.querySelectorAll('.stApp').forEach(el => {
                el.style.backgroundColor = '#FFFFFF';
                el.style.color = '#31333F';
            });
        </script>
    """, height=0)

# --- 4. 业务内容 ---
st.markdown("<h1>🎓 内审员证书智能制作工具</h1>", unsafe_allow_html=True)

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
    raw_data = temp_df.dropna(how='all').to_dict('records')
    data_to_process = [{k: str(v).strip() for k, v in row.items() if v is not None} for row in raw_data if any(row.values())]
else:
    c1, c2 = st.columns([2, 3])
    with c1:
        example_data = {"证书编号":["T-2025-001 (示例)"],"姓名":["张三 (示例)"],"身份证号":["440683..."],"培训日期":["2025年9月"],"标准号":["ISO9001"]}
        df_ex = pd.DataFrame(example_data)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            df_ex.to_excel(writer, index=False)
            ws = writer.sheets['Sheet1']
            for i in range(1, 6): ws.column_dimensions[get_column_letter(i)].width = 20
            for cell in ws[2]: cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        st.download_button("📥 下载标准模板", data=buf.getvalue(), file_name="学员信息上传模板.xlsx")
    with c2:
        up = st.file_uploader("上传文件", type=["xlsx", "csv"], label_visibility="collapsed")
        if up:
            df = pd.read_csv(up, dtype=str).fillna("") if up.name.endswith('.csv') else pd.read_excel(up, dtype=str).fillna("")
            data_to_process = [row for row in df.to_dict('records') if "示例" not in str(row.get('姓名',''))]
            if data_to_process: st.success(f"✅ 已加载 {len(data_to_process)} 条有效数据")

# 第三步
st.markdown("---")
st.markdown("### 第三步：模板确认与生成")
if os.path.exists(DEFAULT_TEMPLATE):
    t_opt = st.radio("模板选择：", ["使用内置模板", "上传本地新模板"], horizontal=True)
    t_path = DEFAULT_TEMPLATE if t_opt == "使用内置模板" else st.file_uploader("上传自定义模板", type=["docx"])
else:
    t_path = st.file_uploader("请上传 Word 模板", type=["docx"])

if t_path and data_to_process:
    if st.button("🚀 开始批量制作合并文档", use_container_width=True):
        try:
            master, bar, count = None, st.progress(0), 0
            for i, row in enumerate(data_to_process):
                name_v = str(row.get('姓名', '')).replace('nan', '').strip()
                if not name_v: continue
                count += 1
                doc = DocxTemplate(t_path)
                doc.render({'number': str(row.get('证书编号','')), 'name': name_v, 'id_card': str(row.get('身份证号','')), 'date': str(row.get('培训日期','')), 'standards': str(row.get('标准号',''))})
                tmp = io.BytesIO(); doc.save(tmp); tmp.seek(0)
                cur = Document(tmp)
                if master is None:
                    master = cur
                    composer = Composer(master)
                else:
                    master.add_page_break(); composer.append(cur)
                bar.progress((i + 1) / len(data_to_process))
            if master:
                out = io.BytesIO(); master.save(out); out.seek(0)
                st.balloons()
                st.download_button(f"🎁 下载汇总文档({count}份)", out.getvalue(), "证书汇总.docx", use_container_width=True)
        except Exception as e: st.error(f"制作失败：{e}")

# --- 5. 底部 Logo 墙与版权 ---
st.markdown("---")
footer_html = """<div style="text-align:center;margin-top:40px;padding-bottom:20px;width:100%;"><div style="display:flex;justify-content:center;align-items:center;gap:20px;margin-bottom:15px;flex-wrap:wrap;"><img src="https://cdn.jsdelivr.net/gh/devicons/devicon/icons/github/github-original.svg" width="22" style="opacity:0.7;"><img src="https://www.vectorlogo.zone/logos/cloudflare/cloudflare-ar21.svg" width="55" style="opacity:0.7;"><img src="https://www.vectorlogo.zone/logos/vercel/vercel-ar21.svg" width="55" style="opacity:0.7;"><img src="https://cdn.jsdelivr.net/gh/devicons/devicon/icons/vuejs/vuejs-original.svg" width="22" style="opacity:0.7;"><img src="https://www.vectorlogo.zone/logos/tailwindcss/tailwindcss-icon.svg" width="22" style="opacity:0.7;"></div><div style="font-size:13px;color:#666;line-height:1.6;font-family:sans-serif;"><p style="margin:0;">© 2026 Jiachen Tu. All rights reserved.</p></div></div>"""
st.markdown(footer_html, unsafe_allow_html=True)
