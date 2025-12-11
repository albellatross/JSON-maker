import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import json
import io
import zipfile
import random

# ================= 🎨 1. COPILOT DESIGN TOKENS =================
# 基于 Copilot Consumer Design System 提取的变量
MY_DESIGN_TOKENS = {
    # 核心色彩
    "bg_color": "#FFF6F0",            # Oat 100 (Primary Background)
    "surface_color": "rgba(255, 255, 255, 0.60)", # Soft, translucent white
    "surface_active": "#FFFFFF",      # Active/Focus surface
    
    # 文本色彩
    "text_primary": "#311F10",        # Dark Oat (Primary Text)
    "text_secondary": "#594134",      # Secondary Text
    "accent_color": "#311F10",        # 使用 Dark Oat 作为主交互色，保持 Grounded
    
    # 几何与阴影
    "radius_card": "40px",            # Large radius for cards
    "radius_pill": "999px",           # Pill radius for buttons/chips
    "shadow_tinted": "0 8px 24px -4px rgba(210, 150, 120, 0.20), 0 4px 8px -2px rgba(210, 150, 120, 0.15)", # Tinted Shadow
    
    # 字体
    "font_family": "'Segoe UI', 'Helvetica Neue', Helvetica, Arial, sans-serif" # 近似 Ginto 的几何无衬线体
}

# ================= 2. CSS 注入引擎 (Copilot Style) =================
def inject_copilot_css(tokens):
    css = f"""
    <style>
        /* === 全局重置 === */
        .stApp {{
            background-color: {tokens['bg_color']};
            font-family: {tokens['font_family']};
            color: {tokens['text_primary']};
        }}
        
        /* 隐藏 Streamlit 默认头部，沉浸式体验 */
        header {{visibility: hidden;}}
        .block-container {{padding-top: 3rem; padding-bottom: 5rem; max-width: 1200px;}}

        /* === 排版系统 === */
        h1, h2, h3 {{
            color: {tokens['text_primary']} !important;
            font-weight: 600 !important;
            letter-spacing: -0.02em !important; /* Display tight spacing */
        }}
        p, label, .stMarkdown {{
            color: {tokens['text_secondary']} !important;
            line-height: 1.6 !important; /* Calm reading rhythm */
        }}

        /* === Copilot 卡片风格 === */
        /* 定制 Streamlit 的 container 为大圆角卡片 */
        [data-testid="stVerticalBlock"] > [style*="flex-direction: column;"] > [data-testid="stVerticalBlock"] {{
            background-color: {tokens['surface_color']};
            border-radius: {tokens['radius_card']};
            padding: 2rem;
            box-shadow: {tokens['shadow_tinted']};
            border: 1px solid rgba(255,255,255,0.4);
            backdrop-filter: blur(10px); /* 磨砂玻璃质感 */
        }}

        /* === 输入框美化 (Composer Style) === */
        .stTextArea textarea, .stTextInput input, .stNumberInput input {{
            background-color: #FFFFFF !important;
            border: 1px solid transparent !important;
            border-radius: 24px !important; /* 较大的圆角 */
            color: {tokens['text_primary']} !important;
            padding: 1rem !important;
            box-shadow: 0 2px 6px rgba(210, 150, 120, 0.05) !important;
            transition: all 0.3s ease;
        }}
        .stTextArea textarea:focus, .stTextInput input:focus {{
            box-shadow: 0 4px 12px rgba(210, 150, 120, 0.15) !important;
            transform: translateY(-1px);
        }}

        /* === 按钮美化 (Pill Shape) === */
        .stButton button {{
            border-radius: {tokens['radius_pill']} !important;
            font-weight: 600 !important;
            border: none !important;
            padding: 0.6rem 1.5rem !important;
            transition: all 0.2s ease !important;
        }}
        
        /* Primary Button (Dark Oat) */
        div[data-testid="stButton"] > button[kind="primary"] {{
            background-color: {tokens['accent_color']} !important;
            color: #FFFFFF !important;
            box-shadow: 0 4px 12px rgba(49, 31, 16, 0.2);
        }}
        div[data-testid="stButton"] > button[kind="primary"]:hover {{
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(49, 31, 16, 0.3);
        }}

        /* Secondary Button (Soft Outline) */
        div[data-testid="stButton"] > button[kind="secondary"] {{
            background-color: transparent !important;
            border: 1px solid rgba(89, 65, 52, 0.2) !important;
            color: {tokens['text_secondary']} !important;
        }}
        div[data-testid="stButton"] > button[kind="secondary"]:hover {{
            background-color: rgba(255, 246, 240, 0.8) !important;
            border-color: {tokens['text_primary']} !important;
            color: {tokens['text_primary']} !important;
        }}

        /* === 进度条颜色 (Tinted Warmth) === */
        .stProgress > div > div > div > div {{
            background-color: {tokens['text_primary']};
        }}
        
        /* === 图片圆角 === */
        img {{
            border-radius: 32px !important; 
        }}

    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

# ================= 3. 数据与逻辑 =================
REMIX_MASTER_LIST = [
    {"label": "Want a wider view?", "prompt": "Create an expanded image with extended space."},
    {"label": "Want to zoom in?", "prompt": "Create a micro-detail close-up variant of this image."},
    {"label": "Try paper cut style?", "prompt": "Remake this image in a modern paper cut style with layered colors and soft shadows."},
    {"label": "Make this embroidery style?", "prompt": "Remake this image in a textile embroidery style with visible stitched threads."},
    {"label": "Change to Pixel Art", "prompt": "Create this picture as a retro pixel art, with nostalgic detail and game shading."},
    {"label": "Apply Glitch Effect", "prompt": "Remake this image as a glitch digital art, with pixel splits and cyberpunk noise."},
    {"label": "Change to Watercolor", "prompt": "Create this picture as a watercolor painting."},
    {"label": "Change to Impressionism", "prompt": "Create this picture as an Impressionist painting, with loose brushwork, luminous color, and fleeting light."},
    {"label": "Draw with colored pencil?", "prompt": "Remake this image as a colored pencil drawing."},
    {"label": "Try fine-line style?", "prompt": "Remake this image as a Chinese Gongbi painting with precise outlines, soft washes, and detailed forms."},
    {"label": "Try Chinese paper cut style?", "prompt": "Remake this image as a Chinese paper cut, with red silhouettes, cultural motifs, and symmetrical patterns."},
    {"label": "Try Ukiyo-e style?", "prompt": "Remake this image as a Japanese Ukiyo-e, with woodblock texture, flat colors, and flowing lines."},
    {"label": "Make this a portrait?", "prompt": "Remake this image as a photo portrait, with natural light, and shallow depth."},
    {"label": "Make this stained glass?", "prompt": "Remake this image as a stained glass design with colorful panes, bold outlines, and glowing light."},
    {"label": "Try silkscreen style?", "prompt": "Remake this image as a silkscreen print."},
    {"label": "Make this anime?", "prompt": "Remake this image as an anime illustration with expressive light and a dynamic layout."},
    {"label": "Add sepia tone?", "prompt": "Remake this image as a sepia-toned memory with aged paper texture."},
    {"label": "Make this pop art?", "prompt": "Remake this image as a high-saturation pop art, with bold blocks and hues."},
    {"label": "Make this a gradient mesh?", "prompt": "Remake this image as a gradient mesh, blending colors seamlessly across the composition."},
    {"label": "Make this a 3D figure?", "prompt": "Remake this image as a photorealistic 3D render of a collectible figure, made of real materials like resin or plastic with cinematic lighting, studio backdrop, and ultra-fine modeling detail."},
    {"label": "Try duotone colors?", "prompt": "Remake this image as a duotone image."},
    {"label": "Make this monochrome?", "prompt": "Remake this image as a monochrome image."},
    {"label": "Add neon lighting?", "prompt": "Remake this image as a neon-lit scene with vibrant color contrasts."},
    {"label": "Make this mechanical?", "prompt": "Create a mechanical version of the subject with exposed gears, metallic joints, and precise components."},
    {"label": "Make this crystal?", "prompt": "Remake this image to be in an iridescent fantasy realm with the subject as translucent glass or crystal, glowing and refracted."}
]

def process_ppt_file(uploaded_file, start_id):
    prs = Presentation(uploaded_file)
    current_id = int(start_id)
    extracted_data = []
    image_storage = {}
    for index, slide in enumerate(prs.slides):
        slide_info = {"id": str(current_id), "original_prompt_text": "", "image_filename": ""}
        found_image = False
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and not found_image:
                image_filename = f"{current_id}.png"
                image_storage[image_filename] = shape.image.blob
                slide_info["image_filename"] = image_filename
                found_image = True
            if shape.has_text_frame:
                text = shape.text.strip()
                if len(text) > 10 and len(text) > len(slide_info["original_prompt_text"]):
                    slide_info["original_prompt_text"] = text
        if found_image:
            extracted_data.append(slide_info)
            current_id += 1
    return extracted_data, image_storage

def create_final_zip(processed_jsons, image_storage):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for filename, json_data in processed_jsons.items():
            json_str = json.dumps(json_data, indent=4, ensure_ascii=False)
            zip_file.writestr(f"jsons/{filename}", json_str)
        for json_filename, json_data in processed_jsons.items():
            img_name = f"{json_data['id']}.png"
            if img_name in image_storage:
                zip_file.writestr(f"images/{img_name}", image_storage[img_name])
    return zip_buffer

def get_random_remix():
    return random.choice(REMIX_MASTER_LIST)

# ================= 4. 主程序 =================

st.set_page_config(page_title="Copilot Design Studio", layout="wide", page_icon="🧶")
inject_copilot_css(MY_DESIGN_TOKENS)

# Init Session
if 'data' not in st.session_state: st.session_state.data = []
if 'images' not in st.session_state: st.session_state.images = {}
if 'processed_results' not in st.session_state: st.session_state.processed_results = {}
if 'current_idx' not in st.session_state: st.session_state.current_idx = 0

# --- Phase 1: Upload (Hero-first Layout) ---
if not st.session_state.data:
    # 居中布局，营造“Calm center zone”
    st.markdown("<br><br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown(f"""
        <div style='text-align: center; margin-bottom: 32px;'>
            <h1 style='font-size: 3rem;'>Let’s create together.</h1>
            <p style='font-size: 1.2rem; color: {MY_DESIGN_TOKENS['text_secondary']};'>
                Upload your PPT to start the prompt engineering workflow.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        with st.container(border=True): # Copilot Card Style
            st.markdown("### 📂 New Project")
            uploaded_ppt = st.file_uploader("Upload PPTX", type=["pptx"])
            st.markdown("<br>", unsafe_allow_html=True)
            col_in1, col_in2 = st.columns([2, 1])
            with col_in2:
                start_id = st.number_input("Start ID", value=453, step=1)
            
            st.markdown("<br>", unsafe_allow_html=True)
            if uploaded_ppt:
                if st.button("Start Workflow", type="primary", use_container_width=True):
                    with st.spinner("Analyzing content..."):
                        try:
                            data, images = process_ppt_file(uploaded_ppt, start_id)
                            st.session_state.data = data
                            st.session_state.images = images
                            st.session_state.current_idx = 0
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error: {e}")

# --- Phase 2: Editor (Content Detail Page) ---
else:
    # 侧边栏：Settings / Profile Page style
    with st.sidebar:
        st.markdown("### Progress")
        total = len(st.session_state.data)
        done = len(st.session_state.processed_results)
        st.progress(done / total if total > 0 else 0)
        st.caption(f"Completed: {done} / {total}")
        
        st.markdown("---")
        if done > 0:
            zip_buffer = create_final_zip(st.session_state.processed_results, st.session_state.images)
            st.download_button("Download ZIP", data=zip_buffer.getvalue(), file_name="dataset.zip", mime="application/zip", type="primary", use_container_width=True)
        else:
            st.info("Process one image to export.")
            
        st.markdown("---")
        if st.button("Reset Project", type="secondary"):
            st.session_state.clear()
            st.rerun()

    # 主界面
    current_item = st.session_state.data[st.session_state.current_idx]
    current_id = current_item['id']
    img_name = current_item['image_filename']

    # 顶部标题 (Ginto Nord style - Display Strong)
    st.markdown(f"## Editor <span style='font-weight: 400; color: #594134; font-size: 0.8em; margin-left: 12px;'>ID {current_id}</span>", unsafe_allow_html=True)

    col_L, col_R = st.columns([1, 1.4])

    with col_L:
        # 左侧图片区
        with st.container(border=True):
            if img_name in st.session_state.images:
                st.image(st.session_state.images[img_name], use_container_width=True)
            else:
                st.error("Image missing")

    with col_R:
        # 右侧编辑区 - Composer Style
        st.markdown("**Main Prompt**")
        default_text = current_item['original_prompt_text']
        if not default_text.strip().lower().startswith("create"):
            default_text = "Create an image of " + default_text
        main_prompt = st.text_area("main_prompt",
