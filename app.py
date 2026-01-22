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

# --- 1. 注入 Silk 着色器背景（基于您提供的 Shader 逻辑） ---
def inject_silk_shader_bg():
    # 我们将 React 逻辑转译为原生 Three.js 脚本，直接嵌入 HTML
    silk_html = """
    <div id="silk-bg" style="position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; z-index: -1;"></div>
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

        const container = document.getElementById('silk-bg');
        const scene = new THREE.Scene();
        const camera = new THREE.OrthographicCamera(-1, 1, 1, -1, 0, 1);
        const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
        renderer.setSize(window.innerWidth, window.innerHeight);
        container.appendChild(renderer.domElement);

        const uniforms = {
            uTime: { value: 0 },
            uColor: { value: new THREE.Color("#7B7481") }, // 使用您代码中的颜色
            uSpeed: { value: 4.3 },
            uScale: { value: 0.5 },
            uNoiseIntensity: { value: 1.5 }
        };

        const geometry = new THREE.PlaneGeometry(2, 2);
        const material = new THREE.ShaderMaterial({ uniforms, vertexShader, fragmentShader });
        const mesh = new THREE.Mesh(geometry, material);
        scene.add(mesh);

        function animate(time) {
            uniforms.uTime.value = time * 0.0005; // 对应 React 中的 delta 逻辑
            renderer.render(scene, camera);
            requestAnimationFrame(animate);
        }

        window.onresize = () => {
            renderer.setSize(window.innerWidth, window.innerHeight);
        };

        requestAnimationFrame(animate);
    </script>
    <style>
        /* 强制 Streamlit 背景透明 */
        .stApp { background: transparent !important; }
        
        /* 步骤框去白、磨砂化 */
        div[data-testid="stVerticalBlock"] > div {
            background-color: transparent !important;
        }
        
        h3 {
            background: rgba(255, 255, 255, 0.15) !important;
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
            padding: 10px 15px !important;
            border-radius: 12px !important;
            color: white !important;
            border: 1px solid rgba(255, 255, 255, 0.1);
        }

        h1 { color: white !important; text-shadow: 2px 2px 10px rgba(0,0,0,0.2); }
    </style>
    """
    components.html(silk_html, height=0)

# --- 2. 核心功能代码 ---
st.set_page_config(page_title="证书智能制作工具", layout="centered")
inject_silk_shader_bg()

st.title("🎓 内审员证书智能制作工具")

# 第一步：模式选择
st.markdown("### 第一步：选择录入模式")
mode = st.radio("方式：", ["Excel 上传", "网页填写"], index=0, horizontal=True, label_visibility="collapsed")

DEFAULT_TEMPLATE = "内审员证书.docx"
data_to_process = []

# 第二步：准备数据
st.markdown("### 第二步：录入学员信息")

if mode == "网页填写":
    init_df = pd.DataFrame({
        "序号": range(1, 51),
        "证书编号": [None]*50, "姓名": [None]*50, "身份证号": [None]*50, "培训日期": [None]*50, "标准号": [None]*50
    })
    edited_df = st.data_editor(init_df, use_container_width=True, hide_index=True)
    temp_df = edited_df.drop(columns=["序号"]).dropna(how='all')
    data_to_process = [row for row in temp_df.to_dict('records') if row.get('姓名')]
else:
    c1, c2 = st.columns([2, 3])
    with c1:
        # 带有黄色示例行的模板生成
        df_ex = pd.DataFrame({"证书编号":["编号(示例)"], "姓名":["张三(示例)"], "身份证号":["123456..."], "培训日期":["2026-01"], "标准号":["ISO9001"]})
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            df_ex.to_excel(writer, index=False)
            ws = writer.sheets['Sheet1']
            for cell in ws[2]: cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        st.download_button("📥 下载带示例模板", data=buf.getvalue(), file_name="学员模板.xlsx")
    with c2:
        up = st.file_uploader("上传 Excel", type=["xlsx"], label_visibility="collapsed")
        if up:
            df = pd.read_excel(up, dtype=str).fillna("")
            data_to_process = [row for row in df.to_dict('records') if "示例" not in str(row.get('姓名'))]

# 第三步：生成
st.markdown("### 第三步：模板确认与生成")
if os.path.exists(DEFAULT_TEMPLATE):
    t_path = DEFAULT_TEMPLATE
    st.success("✅ 已检测到默认 Word 模板")
else:
    t_path = st.file_uploader("请上传 Word 模板", type=["docx"])

if t_path and data_to_process:
    if st.button("🚀 批量生成汇总文档", use_container_width=True):
        try:
            master = None
            bar = st.progress(0)
            for i, row in enumerate(data_to_process):
                doc = DocxTemplate(t_path)
                doc.render(row)
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
                bar.progress((i + 1) / len(data_to_process))
            
            out = io.BytesIO()
            master.save(out)
            st.balloons()
            st.download_button("🎁 下载汇总文档", out.getvalue(), "证书汇总.docx", use_container_width=True)
        except Exception as e:
            st.error(f"出错啦: {e}")
