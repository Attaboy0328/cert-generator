import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import os
from docx import Document
from docxcompose.composer import Composer
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# --- 1. 注入 Silk 丝绸着色器背景 (基于您提供的 Shader 逻辑) ---
def inject_silk_shader_bg():
    # 我们使用原生 Three.js 还原您的 React 代码逻辑
    silk_html = """
    <div id="silk-container" style="position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; z-index: -1;"></div>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
    <script>
        const vertexShader = `
            varying vec2 vUv;
            void main() {
                vUv = uv;
                gl_Position = projectionMatrix * modelViewMatrix * vec4(position, 1.0);
            }
        `;

        const fragmentShader = `
            varying vec2 vUv;
            uniform float uTime;
            uniform vec3  uColor;
            uniform float uSpeed;
            uniform float uScale;
            uniform float uNoiseIntensity;

            const float e = 2.71828182845904523536;

            float noise(vec2 texCoord) {
                vec2 r = (e * sin(e * texCoord));
                return fract(r.x * r.y * (1.0 + texCoord.x));
            }

            void main() {
                float rnd = noise(gl_FragCoord.xy);
                vec2 uv = vUv * uScale;
                float tOffset = uSpeed * uTime;

                uv.y += 0.03 * sin(8.0 * uv.x - tOffset);

                float pattern = 0.6 + 0.4 * sin(5.0 * (uv.x + uv.y + 
                                cos(3.0 * uv.x + 5.0 * uv.y) + 0.02 * tOffset) + 
                                sin(20.0 * (uv.x + uv.y - 0.1 * tOffset)));

                vec3 col = uColor * pattern - rnd / 15.0 * uNoiseIntensity;
                gl_FragColor = vec4(col, 1.0);
            }
        `;

        const scene = new THREE.Scene();
        const camera = new THREE.OrthographicCamera(-1, 1, 1, -1, 0, 1);
        const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
        renderer.setSize(window.innerWidth, window.innerHeight);
        document.getElementById('silk-container').appendChild(renderer.domElement);

        const uniforms = {
            uTime: { value: 0 },
            uColor: { value: new THREE.Color("#7B7481") },
            uSpeed: { value: 4.3 },
            uScale: { value: 0.5 },
            uNoiseIntensity: { value: 1.5 }
        };

        const geometry = new THREE.PlaneGeometry(2, 2);
        const material = new THREE.ShaderMaterial({ uniforms, vertexShader, fragmentShader });
        const mesh = new THREE.Mesh(geometry, material);
        scene.add(mesh);

        function animate(time) {
            uniforms.uTime.value = time * 0.001;
            renderer.render(scene, camera);
            requestAnimationFrame(animate);
        }
        
        window.addEventListener('resize', () => {
            renderer.setSize(window.innerWidth, window.innerHeight);
        });
        
        requestAnimationFrame(animate);
    </script>
    <style>
        .stApp { background: transparent !important; }
        /* 保持毛玻璃标题块 */
        h3 {
            background: rgba(255, 255, 255, 0.15) !important;
            backdrop-filter: blur(15px);
            padding: 10px 20px;
            border-radius: 12px;
            color: white !important;
            border: 1px solid rgba(255, 255, 255, 0.1);
        }
        div[data-testid="stVerticalBlock"] > div { background: transparent !important; }
    </style>
    """
    st.components.v1.html(silk_html, height=0)

# 配置
st.set_page_config(page_title="证书智能制作工具", layout="centered")
inject_silk_shader_bg()

st.title("🎓 内审员证书智能制作工具")

# --- 第一步：录入模式 ---
st.markdown("### 第一步：选择录入模式")
mode = st.radio("方式：", ["Excel 文件上传", "网页表格填写 (支持粘贴)"], index=0, horizontal=True, label_visibility="collapsed")

DEFAULT_TEMPLATE = "内审员证书.docx"
data_to_process = []

# --- 第二步：准备数据 ---
st.markdown("### 第二步：填写或上传信息")

if mode == "网页表格填写 (支持粘贴)":
    init_df = pd.DataFrame({
        "序号": [i for i in range(1, 101)],
        "证书编号": [None] * 100, "姓名": [None] * 100, "身份证号": [None] * 100, "培训日期": [None] * 100, "标准号": [None] * 100,
    })
    edited_df = st.data_editor(init_df, num_rows="fixed", use_container_width=True, hide_index=True, height=385)
    
    # 清洗逻辑
    temp_df = edited_df.drop(columns=["序号"]).dropna(how='all')
    data_to_process = []
    for _, row in temp_df.iterrows():
        clean_row = {k: str(v).strip() for k, v in row.items() if pd.notna(v) and str(v).lower() != 'none'}
        if clean_row: data_to_process.append(clean_row)
else:
    c1, c2 = st.columns([2, 3])
    with c1:
        # 模板下载
        df_ex = pd.DataFrame({"证书编号": ["T-2026-001 (示例)"], "姓名": ["张三 (示例)"], "身份证号": ["440683199001010001"], "培训日期": ["2026年1月23日"], "标准号": ["ISO9001"]})
        template_buffer = io.BytesIO()
        with pd.ExcelWriter(template_buffer, engine='openpyxl') as writer:
            df_ex.to_excel(writer, index=False)
            ws = writer.sheets['Sheet1']
            for cell in ws[2]: cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        st.download_button("📥 下载标准模板", data=template_buffer.getvalue(), file_name="模板.xlsx")
    with c2:
        uploaded_data = st.file_uploader("上传文件", type=["xlsx", "csv"], label_visibility="collapsed")
    if uploaded_data:
        df = pd.read_csv(uploaded_data, dtype=str).fillna("") if uploaded_data.name.endswith('.csv') else pd.read_excel(uploaded_data, dtype=str).fillna("")
        data_to_process = [row for row in df.to_dict('records') if "示例" not in str(row.get('姓名', ''))]

# --- 第三步：生成 ---
st.markdown("### 第三步：模板确认与生成")
if os.path.exists(DEFAULT_TEMPLATE):
    t_opt = st.radio("模板：", ["使用内置", "上传本地"], horizontal=True, label_visibility="collapsed")
    t_path = DEFAULT_TEMPLATE if t_opt == "使用内置" else st.file_uploader("上传 Word", type=["docx"])
else:
    t_path = st.file_uploader("请上传 Word 模板", type=["docx"])

if t_path and data_to_process:
    if st.button("🚀 开始批量制作合并文档", use_container_width=True):
        try:
            master_doc, prog, count = None, st.progress(0), 0
            for i, row in enumerate(data_to_process):
                name_val = row.get('姓名', '').replace('nan', '').strip()
                if not name_val: continue
                count += 1
                doc = DocxTemplate(t_path)
                doc.render({'number': row.get('证书编号',''), 'name': name_val, 'id_card': row.get('身份证号',''), 'date': row.get('培训日期',''), 'standards': row.get('标准号','')})
                t_io = io.BytesIO(); doc.save(t_io); t_io.seek(0)
                cur = Document(t_io)
                if master_doc is None:
                    master_doc = cur
                    composer = Composer(master_doc)
                else:
                    master_doc.add_page_break(); composer.append(cur)
                prog.progress((i + 1) / len(data_to_process))
            if master_doc:
                out = io.BytesIO(); master_doc.save(out); out.seek(0)
                st.balloons()
                st.download_button(f"🎁 下载汇总文档({count}份)", out.getvalue(), "证书汇总.docx", use_container_width=True)
        except Exception as e: st.error(f"制作失败：{e}")
