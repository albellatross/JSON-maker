import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import json
import io
import zipfile
import random
import urllib.parse

# ================= 🎨 1. DESIGN TOKENS =================
MY_DESIGN_TOKENS = {
    "bg_color": "#FFF6F0",            
    "surface_color": "rgba(255, 255, 255, 0.90)", 
    "text_primary": "#311F10",        
    "text_secondary": "#594134",
    "accent_color": "#311F10",
    "radius_card": "24px",
    "radius_pill": "999px",
    "shadow_tinted": "0 4px 20px rgba(210, 150, 120, 0.15)",
    "font_family": "'Segoe UI', 'Microsoft YaHei', sans-serif"
}

def inject_copilot_css(tokens):
    css = f"""
    <style>
        .stApp {{ background-color: {tokens['bg_color']}; font-family: {tokens['font_family']}; color: {tokens['text_primary']}; }}
        header {{visibility: hidden;}}
        .block-container {{padding-top: 1rem; padding-bottom: 5rem; max-width: 1000px;}}
        
        h1, h2, h3 {{ color: {tokens['text_primary']} !important; font-weight: 600 !important; }}
        
        /* 卡片样式 */
        [data-testid="stVerticalBlock"] > [style*="flex-direction: column;"] > [data-testid="stVerticalBlock"] {{
            background-color: {tokens['surface_color']};
            border-radius: {tokens['radius_card']};
            padding: 1.5rem;
            box-shadow: {tokens['shadow_tinted']};
            border: 1px solid rgba(255,255,255,0.6);
        }}
        
        .stTextArea textarea, .stTextInput input {{
            background-color: #FFFFFF !important;
            border-radius: 12px !important;
            border: 1px solid transparent !important;
            box-shadow: 0 2px 6px rgba(210, 150, 120, 0.05) !important;
        }}

        .stButton button {{ border-radius: {tokens['radius_pill']} !important; font-weight: 600 !important; border: none !important; }}
        div[data-testid="stButton"] > button[kind="primary"] {{ background-color: {tokens['accent_color']} !important; color: #FFFFFF !important; }}
        
        img {{ border-radius: 16px !important; }}
        .css-1v0mbdj a {{ display: none; }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

# ================= 2. DATA LISTS =================

# 🇬🇧 English List (Original 25)
REMIX_LIST_EN = [
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

# 🇨🇳 Chinese List (Translations of the 25)
REMIX_LIST_ZH = [
    {"label": "想要更宽的视野？", "prompt": "生成一张视野更广阔的扩展图像，包含延伸的空间。"},
    {"label": "想要放大看？", "prompt": "生成这张图片的微距细节特写版本。"},
    {"label": "试试剪纸风格？", "prompt": "将这张图片重制为现代剪纸风格，带有分层色彩和柔和阴影。"},
    {"label": "做成刺绣风格？", "prompt": "将这张图片重制为纺织刺绣风格，带有可见的缝合线。"},
    {"label": "变为像素艺术", "prompt": "将这张图片创作成复古像素艺术风格，带有怀旧细节和游戏光影。"},
    {"label": "应用故障效果", "prompt": "将这张图片重制为数字故障艺术，带有像素撕裂和赛博朋克噪点。"},
    {"label": "变为水彩画", "prompt": "将这张图片创作成水彩画。"},
    {"label": "变为印象派", "prompt": "将这张图片创作成印象派画作，带有松散的笔触、明亮的色彩和稍纵即逝的光线。"},
    {"label": "用彩铅绘制？", "prompt": "将这张图片重制为彩色铅笔画。"},
    {"label": "试试工笔画风格？", "prompt": "将这张图片重制为中国工笔画，带有精确的轮廓、柔和的晕染和细腻的形态。"},
    {"label": "试试中国剪纸？", "prompt": "将这张图片重制为中国传统剪纸，带有红色剪影、文化纹样和对称图案。"},
    {"label": "试试浮世绘风格？", "prompt": "将这张图片重制为日本浮世绘，带有木刻纹理、平涂色彩和流畅线条。"},
    {"label": "做成肖像照？", "prompt": "将这张图片重制为摄影肖像，带有自然光和浅景深。"},
    {"label": "做成彩色玻璃？", "prompt": "将这张图片重制为彩色玻璃设计，带有多彩的玻璃块、粗轮廓和发光效果。"},
    {"label": "试试丝网印刷？", "prompt": "将这张图片重制为丝网印刷品。"},
    {"label": "做成动漫？", "prompt": "将这张图片重制为动漫插画，带有表现力的光影和动态布局。"},
    {"label": "增加怀旧色调？", "prompt": "将这张图片重制为带有陈旧纸张纹理的怀旧棕褐色调记忆。"},
    {"label": "做成波普艺术？", "prompt": "将这张图片重制为高饱和度的波普艺术，带有大胆的色块和色调。"},
    {"label": "做成渐变网格？", "prompt": "将这张图片重制为渐变网格风格，颜色在画面中无缝融合。"},
    {"label": "做成3D手办？", "prompt": "将这张图片重制为逼真的3D收藏手办渲染图，由树脂或塑料等真实材料制成，带有电影级布光、摄影棚背景和超精细建模细节。"},
    {"label": "试试双色调？", "prompt": "将这张图片重制为双色调图像。"},
    {"label": "做成单色？", "prompt": "将这张图片重制为单色图像。"},
    {"label": "增加霓虹灯光？", "prompt": "将这张图片重制为带有强烈色彩对比的霓虹灯场景。"},
    {"label": "做成机械风格？", "prompt": "生成主体的机械版本，带有外露的齿轮、金属关节和精密组件。"},
    {"label": "做成水晶质感？", "prompt": "将这张图片重制为彩虹色的幻想领域，主体呈现发光和折射的半透明玻璃或水晶质感。"}
]

COPILOT_GEN_INSTRUCTION = """A remix prompt consists of a short, 2–5-word title and an instruction.
Please write 5 remix prompts for me based on the uploaded image.
Format:
Label: [Title]
Prompt: [Instruction]"""

def get_random_remix(language_mode):
    if language_mode == "Chinese/中文":
        return random.choice(REMIX_LIST_ZH)
    else:
        return random.choice(REMIX_LIST_EN)

# 🔥 核心修复：把更新逻辑放进回调函数里
def randomize_callback(index, session_key_root, lang, current_id_val):
    new_remix = get_random_remix(lang)
    # 1. 更新后台数据列表
    st.session_state[session_key_root][index] = new_remix
    # 2. 直接更新组件的Key，这样下次渲染时输入框就会显示新值
    st.session_state[f"l_{current_id_val}_{index}"] = new_remix['label']
    st.session_state[f"p_{current_id_val}_{index}"] = new_remix['prompt']

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
                img_name = f"{current_id}.png"
                image_storage[img_name] = shape.image.blob
                slide_info["image_filename"] = img_name
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

# ================= 3. MAIN UI =================
st.set_page_config(page_title="Editor", layout="wide", page_icon="🧩")
inject_copilot_css(MY_DESIGN_TOKENS)

# Session Setup
if 'data' not in st.session_state: st.session_state.data = []
if 'images' not in st.session_state: st.session_state.images = {}
if 'processed_results' not in st.session_state: st.session_state.processed_results = {}
if 'current_idx' not in st.session_state: st.session_state.current_idx = 0

# --- Phase 1: Upload ---
if not st.session_state.data:
    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown(f"<div style='text-align: center;'><h1>Remix Editor</h1><p style='color:#594134'>Clean Workflow</p></div>", unsafe_allow_html=True)
        with st.container(border=True):
            uploaded_ppt = st.file_uploader("Upload PPTX", type=["pptx"])
            start_id = st.number_input("Start ID", value=453, step=1)
            st.markdown("<br>", unsafe_allow_html=True)
            if uploaded_ppt:
                if st.button("🚀 Load Slides", type="primary", use_container_width=True):
                    with st.spinner("Processing..."):
                        try:
                            data, images = process_ppt_file(uploaded_ppt, start_id)
                            st.session_state.data = data
                            st.session_state.images = images
                            st.session_state.current_idx = 0
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error: {e}")

# --- Phase 2: Editor (Clean Mode) ---
else:
    item = st.session_state.data[st.session_state.current_idx]
    current_id = item['id']
    img_name = item['image_filename']

    # Header: ID + Save Button (Minimal)
    col_h1, col_h2 = st.columns([4, 1])
    with col_h1:
        st.markdown(f"## ID {current_id}")
    with col_h2:
        if st.button("💾 Save & Next", type="primary", use_container_width=True, key="save_top"):
            # Trigger logic handled at bottom, button just for UI
            pass 

    # === SECTION 1: Top Image ===
    with st.container(border=True):
        if img_name in st.session_state.images:
            st.image(st.session_state.images[img_name], use_container_width=True)
        else:
            st.error("Missing Image")

    st.markdown("<br>", unsafe_allow_html=True)

    # === SECTION 2: Main Prompt ===
    st.markdown("#### 📝 Main Prompt")
    default_text = item['original_prompt_text']
    if not default_text.strip().lower().startswith("create"):
        default_text = "Create an image of " + default_text
    main_prompt = st.text_area("main_hidden", value=default_text, height=100, key=f"m_{current_id}", label_visibility="collapsed")
    
    st.markdown("<br>", unsafe_allow_html=True)

    # === SECTION 3: Element Remix Suggestions ===
    # 获取语言设置（默认为中文，如果需要更改，在底部菜单）
    lang_mode = st.session_state.get('lang_mode', 'Chinese/中文')
    
    st.markdown(f"#### 🎨 Remix Suggestions ({lang_mode})")
    
    session_key = f"remix_{current_id}"
    # 初始化
    if session_key not in st.session_state:
        st.session_state[session_key] = [get_random_remix(lang_mode) for _ in range(3)]
        
    current_remixes = st.session_state[session_key]

    with st.expander("📋 Show Copilot Instruction"):
        st.code(COPILOT_GEN_INSTRUCTION, language="text")

    # Grid Layout
    r_col1, r_col2, r_col3 = st.columns(3)
    cols = [r_col1, r_col2, r_col3]

    for i in range(3):
        with cols[i]:
            with st.container(border=True):
                # Title + Random Button
                c_title, c_btn = st.columns([4, 1])
                with c_title:
                    l_key = f"l_{current_id}_{i}"
                    l_val = st.text_input(f"L{i}", value=current_remixes[i]['label'], key=l_key, label_visibility="collapsed", placeholder="Label")
                with c_btn:
                    # 🔥 使用 on_click 回调来修复“无法修改 session_state”的错误
                    st.button("🎲", key=f"rnd_{current_id}_{i}", 
                             on_click=randomize_callback,
                             args=(i, session_key, lang_mode, current_id))

                p_key = f"p_{current_id}_{i}"
                p_val = st.text_area(f"P{i}", value=current_remixes[i]['prompt'], height=120, key=p_key, label_visibility="collapsed", placeholder="Prompt")
                
                # 同步回数据列表（防止手动修改丢失）
                # 注意：因为使用了回调，如果是点击按钮触发的刷新，这里的 current_remixes 已经是新的了
                # 如果是用户手动输入，这里会捕获输入值
                if current_remixes[i]['label'] != l_val:
                   current_remixes[i]['label'] = l_val
                if current_remixes[i]['prompt'] != p_val:
                   current_remixes[i]['prompt'] = p_val
                
                # 验证按钮
                if st.button(f"🎨 Verify", key=f"v_{current_id}_{i}", use_container_width=True):
                    clean_prompt = urllib.parse.quote(p_val)
                    seed = random.randint(0, 9999)
                    url = f"https://image.pollinations.ai/prompt/{clean_prompt}?seed={seed}&width=600&height=600&nologo=true"
                    st.session_state[f"poll_img_{current_id}_{i}"] = url
                
                if f"poll_img_{current_id}_{i}" in st.session_state:
                    st.image(st.session_state[f"poll_img_{current_id}_{i}"], caption="Preview", use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    
    # 底部隐形菜单：用于下载和设置语言，不占用主界面空间
    with st.expander("⚙️ Menu (Download / Language)", expanded=False):
        c_menu1, c_menu2 = st.columns(2)
        with c_menu1:
            st.markdown("##### Language")
            # 绑定 session state
            new_lang = st.radio("Select Language", ["English", "Chinese/中文"], index=1, key='lang_mode')
        with c_menu2:
            st.markdown("##### Export")
            total = len(st.session_state.data)
            done = len(st.session_state.processed_results)
            st.write(f"Progress: {done} / {total}")
            if done > 0:
                zip_buffer = create_final_zip(st.session_state.processed_results, st.session_state.images)
                st.download_button("⬇️ Download ZIP", data=zip_buffer.getvalue(), file_name="dataset.zip", mime="application/zip", type="primary")

    # 底部保存逻辑
    if st.button("💾 Save & Next", type="primary", use_container_width=True, key="save_bottom"):
        final_json = { 
            "id": current_id, 
            "prompt": main_prompt, 
            "remixSuggestions": current_remixes,
            "language": lang_mode
        }
        st.session_state.processed_results[f"{current_id}.json"] = final_json
        if st.session_state.current_idx < len(st.session_state.data) - 1:
            st.session_state.current_idx += 1
            st.rerun()
        else:
            st.balloons()
            st.success("All Done!")
