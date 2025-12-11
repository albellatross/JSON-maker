import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import json
import io
import zipfile
import random

# ================= 🎨 1. DESIGN TOKENS (Copilot Style) =================
MY_DESIGN_TOKENS = {
    "bg_color": "#FFF6F0",            # Oat 100
    "surface_color": "rgba(255, 255, 255, 0.70)", 
    "text_primary": "#311F10",        # Dark Oat
    "text_secondary": "#594134",
    "accent_color": "#311F10",
    "radius_card": "32px",
    "radius_pill": "999px",
    "shadow_tinted": "0 8px 24px -4px rgba(210, 150, 120, 0.20), 0 4px 8px -2px rgba(210, 150, 120, 0.15)",
    "font_family": "'Segoe UI', sans-serif"
}

def inject_copilot_css(tokens):
    css = f"""
    <style>
        .stApp {{ background-color: {tokens['bg_color']}; font-family: {tokens['font_family']}; color: {tokens['text_primary']}; }}
        header {{visibility: hidden;}}
        .block-container {{padding-top: 3rem; padding-bottom: 5rem; max-width: 1200px;}}
        h1, h2, h3 {{ color: {tokens['text_primary']} !important; font-weight: 600 !important; letter-spacing: -0.02em !important; }}
        
        /* 卡片风格 */
        [data-testid="stVerticalBlock"] > [style*="flex-direction: column;"] > [data-testid="stVerticalBlock"] {{
            background-color: {tokens['surface_color']};
            border-radius: {tokens['radius_card']};
            padding: 2rem;
            box-shadow: {tokens['shadow_tinted']};
            border: 1px solid rgba(255,255,255,0.5);
        }}
        
        /* 输入框美化 */
        .stTextArea textarea, .stTextInput input {{
            background-color: #FFFFFF !important;
            border-radius: 12px !important;
            border: 1px solid transparent !important;
            box-shadow: 0 2px 6px rgba(210, 150, 120, 0.05) !important;
        }}
        
        /* 代码块复制按钮 */
        .stCode {{ border-radius: 12px !important; }}

        /* 按钮 */
        .stButton button {{ border-radius: {tokens['radius_pill']} !important; font-weight: 600 !important; border: none !important; }}
        div[data-testid="stButton"] > button[kind="primary"] {{ background-color: {tokens['accent_color']} !important; color: #FFFFFF !important; }}
        div[data-testid="stButton"] > button[kind="secondary"] {{ border: 1px solid rgba(89, 65, 52, 0.2) !important; background: transparent !important; }}
        
        img {{ border-radius: 24px !important; }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

# ================= 2. DATA: LISTS & PROMPTS =================

# 你的随机备选库
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

# 你的Copilot生成指令
COPILOT_GEN_INSTRUCTION = """A remix prompt consists of a short style title and a style-transformation instruction that begins with “Create this picture as…”. The title should be a concise 2–5-word summary of the stylistic direction. The prompt should be a single, clear sentence that transforms the original image into a specific art style, medium, or visual treatment.

Please write 5 remix prompts for me based on the uploaded image (each with a title + a “Create this picture as…” prompt, all in different styles)."""

# ================= 3. LOGIC FUNCTIONS =================
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

# ================= 4. MAIN APP =================
st.set_page_config(page_title="Copilot Workflow", layout="wide", page_icon="🎨")
inject_copilot_css(MY_DESIGN_TOKENS)

# Session Init
if 'data' not in st.session_state: st.session_state.data = []
if 'images' not in st.session_state: st.session_state.images = {}
if 'processed_results' not in st.session_state: st.session_state.processed_results = {}
if 'current_idx' not in st.session_state: st.session_state.current_idx = 0

# --- Phase 1: Upload ---
if not st.session_state.data:
    st.markdown("<br><br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown(f"<div style='text-align: center; margin-bottom: 30px;'><h1>🎨 AI Prompt Workflow</h1><p style='color:#594134'>Upload PPT -> Generate -> Verify -> Export</p></div>", unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown("### 📂 Project Setup")
            uploaded_ppt = st.file_uploader("Upload PPTX", type=["pptx"])
            col_in1, col_in2 = st.columns([2, 1])
            with col_in2:
                start_id = st.number_input("Start ID", value=453, step=1)
            st.markdown("<br>", unsafe_allow_html=True)
            if uploaded_ppt:
                if st.button("Start Workflow", type="primary", use_container_width=True):
                    with st.spinner("Processing..."):
                        try:
                            data, images = process_ppt_file(uploaded_ppt, start_id)
                            st.session_state.data = data
                            st.session_state.images = images
                            st.session_state.current_idx = 0
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error: {e}")

# --- Phase 2: Editor (Generate & Verify) ---
else:
    # Sidebar
    with st.sidebar:
        st.markdown("### Progress")
        total = len(st.session_state.data)
        done = len(st.session_state.processed_results)
        st.progress(done / total if total > 0 else 0)
        st.caption(f"Done: {done} / {total}")
        st.markdown("---")
        if done > 0:
            zip_buffer = create_final_zip(st.session_state.processed_results, st.session_state.images)
            st.download_button("📦 Download ZIP", data=zip_buffer.getvalue(), file_name="dataset.zip", mime="application/zip", type="primary", use_container_width=True)
        st.markdown("---")
        if st.button("Reset Project", type="secondary"):
            st.session_state.clear()
            st.rerun()

    # Data Setup
    current_item = st.session_state.data[st.session_state.current_idx]
    current_id = current_item['id']
    img_name = current_item['image_filename']

    st.markdown(f"## Workstation <span style='font-weight:400; font-size:0.8em; color:#594134; margin-left:10px;'>ID {current_id}</span>", unsafe_allow_html=True)

    col_L, col_R = st.columns([1, 1.4])
    
    with col_L:
        # Image Display
        with st.container(border=True):
            if img_name in st.session_state.images:
                st.image(st.session_state.images[img_name], use_container_width=True)
            else:
                st.error("Missing Image")
                
    with col_R:
        # 1. Main Prompt
        st.markdown("**1. Main Prompt**")
        default_text = current_item['original_prompt_text']
        if not default_text.strip().lower().startswith("create"):
            default_text = "Create an image of " + default_text
        main_prompt = st.text_area("main", value=default_text, height=100, key=f"m_{current_id}", label_visibility="collapsed")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 2. Generation Method (Tabbed Interface)
        st.markdown("**2. Remix Suggestions**")
        
        # Initialize Remixes if new
        session_key = f"remix_{current_id}"
        if session_key not in st.session_state:
            st.session_state[session_key] = random.sample(REMIX_MASTER_LIST, 3)
            st.session_state[f"verified_{current_id}"] = [False, False, False]
            
        current_remixes = st.session_state[session_key]
        verified_status = st.session_state.get(f"verified_{current_id}", [False, False, False])

        # === 核心变化：增加生成选项卡 ===
        gen_tab1, gen_tab2 = st.tabs(["🎲 Quick Random (List)", "🧠 Copilot Generate (AI)"])
        
        # Tab 1: 随机列表模式
        with gen_tab1:
            st.caption("Use pre-defined styles from your master list.")
            if st.button("🎲 Randomize All 3", key=f"rand_all_{current_id}", type="secondary"):
                st.session_state[session_key] = random.sample(REMIX_MASTER_LIST, 3)
                # Reset verification
                st.session_state[f"verified_{current_id}"] = [False, False, False]
                st.rerun()

        # Tab 2: Copilot 生成模式
        with gen_tab2:
            st.caption("Copy this instruction to Copilot to generate FRESH ideas.")
            # 显示你的超长指令
            st.code(COPILOT_GEN_INSTRUCTION, language="text")
            st.markdown(f"<a href='https://copilot.microsoft.com/' target='_blank'>↗️ Go to Copilot</a> (Paste instruction + Image -> Copy results back below)", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("**3. Verify & Filter**")

        # 3. Cards for Verification
        for i in range(3):
            with st.container(border=True):
                # Header
                c_head1, c_head2 = st.columns([4, 1])
                with c_head1:
                     l_val = st.text_input(f"Label {i+1}", value=current_remixes[i]['label'], key=f"l_{current_id}_{i}", label_visibility="collapsed", placeholder="Style Title")
                with c_head2:
                    # Verification Checkbox
                    is_checked = st.checkbox("Valid", value=verified_status[i], key=f"chk_{current_id}_{i}")
                    verified_status[i] = is_checked
                
                # Prompt Area
                p_val = st.text_area(f"Prompt {i+1}", value=current_remixes[i]['prompt'], height=60, key=f"p_{current_id}_{i}", label_visibility="collapsed", placeholder="Paste Copilot result here...")
                current_remixes[i]['label'] = l_val
                current_remixes[i]['prompt'] = p_val
                
                # Copy & Swap Controls
                c_copy, c_swap = st.columns([5, 1])
                with c_copy:
                    st.code(p_val, language="text") # Easy copy for verification
                with c_swap:
                    st.write("")
                    if st.button("🎲", key=f"b_{current_id}_{i}", help="Swap single", type="secondary"):
                        st.session_state[session_key][i] = get_random_remix()
                        verified_status[i] = False
                        st.session_state[f"verified_{current_id}"] = verified_status
                        st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💾 Save Verified & Next", type="primary", use_container_width=True):
            final_json = { 
                "id": current_id, 
                "prompt": main_prompt, 
                "remixSuggestions": current_remixes,
                "verificationStatus": verified_status 
            }
            st.session_state.processed_results[f"{current_id}.json"] = final_json
            if st.session_state.current_idx < len(st.session_state.data) - 1:
                st.session_state.current_idx += 1
                st.rerun()
            else:
                st.balloons()
                st.success("All Done!")
