import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import json
import io
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

# ================= 2. DATA LISTS (Strictly 25 - English Only) =================

REMIX_LIST = [
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

COPILOT_GEN_INSTRUCTION = """A remix prompt consists of a short, 2–5-word title and an instruction.
Please write 5 remix prompts for me based on the uploaded image.
Format:
Label: [Title]
Prompt: [Instruction]"""

def get_random_remix():
    return random.choice(REMIX_LIST)

def randomize_callback(index, session_key_root, current_id_val):
    new_remix = get_random_remix()
    st.session_state[session_key_root][index] = new_remix
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

def generate_combined_json(processed_results):
    """
    Export strictly clean JSON structure:
    [ {id, prompt, remixSuggestions: [{label, prompt}, ...]}, ... ]
    """
    combined_list = []
    sorted_keys = sorted(processed_results.keys(), key=lambda x: int(processed_results[x]['id']))
    
    for key in sorted_keys:
        item = processed_results[key]
        clean_item = {
            "id": item.get("id"),
            "prompt": item.get("prompt"),
            "remixSuggestions": [
                {
                    "label": r.get("label"), 
                    "prompt": r.get("prompt")
                } for r in item.get("remixSuggestions", [])
            ]
        }
        combined_list.append(clean_item)
        
    return json.dumps(combined_list, indent=4, ensure_ascii=False)

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
        st.markdown(f"<div style='text-align: center;'><h1>Clean Studio</h1><p style='color:#594134'>Single JSON Output</p></div>", unsafe_allow_html=True)
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

    # Header: ID
    col_h1, col_h2 = st.columns([4, 1])
    with col_h1:
        st.markdown(f"## ID {current_id}")
    with col_h2:
        if st.button("💾 Save & Next", type="primary", use_container_width=True, key="save_top"):
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

    # === SECTION 3: Remix Suggestions ===
    st.markdown(f"#### 🎨 Remix Suggestions")
    
    session_key = f"remix_{current_id}"
    if session_key not in st.session_state:
        st.session_state[session_key] = [get_random_remix() for _ in range(3)]
        
    current_remixes = st.session_state[session_key]

    with st.expander("📋 Show Copilot Instruction"):
        st.code(COPILOT_GEN_INSTRUCTION, language="text")

    # Grid Layout
    r_col1, r_col2, r_col3 = st.columns(3)
    cols = [r_col1, r_col2, r_col3]

    for i in range(3):
        with cols[i]:
            with st.container(border=True):
                c_title, c_btn = st.columns([4, 1])
                with c_title:
                    l_key = f"l_{current_id}_{i}"
                    l_val = st.text_input(f"L{i}", value=current_remixes[i]['label'], key=l_key, label_visibility="collapsed", placeholder="Label")
                with c_btn:
                    # Callback without lang argument
                    st.button("🎲", key=f"rnd_{current_id}_{i}", 
                             on_click=randomize_callback,
                             args=(i, session_key, current_id))

                p_key = f"p_{current_id}_{i}"
                p_val = st.text_area(f"P{i}", value=current_remixes[i]['prompt'], height=120, key=p_key, label_visibility="collapsed", placeholder="Prompt")
                
                # Update List
                if current_remixes[i]['label'] != l_val:
                   current_remixes[i]['label'] = l_val
                if current_remixes[i]['prompt'] != p_val:
                   current_remixes[i]['prompt'] = p_val
                
                # Verify
                if st.button(f"🎨 Verify", key=f"v_{current_id}_{i}", use_container_width=True):
                    clean_prompt = urllib.parse.quote(p_val)
                    seed = random.randint(0, 9999)
                    url = f"https://image.pollinations.ai/prompt/{clean_prompt}?seed={seed}&width=600&height=600&nologo=true"
                    st.session_state[f"poll_img_{current_id}_{i}"] = url
                
                if f"poll_img_{current_id}_{i}" in st.session_state:
                    st.image(st.session_state[f"poll_img_{current_id}_{i}"], caption="Preview", use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    
    # 底部隐形菜单：仅用于下载，无语言选项
    with st.expander("⚙️ Export Menu", expanded=False):
        total = len(st.session_state.data)
        done = len(st.session_state.processed_results)
        st.write(f"Progress: {done} / {total}")
        if done > 0:
            json_str = generate_combined_json(st.session_state.processed_results)
            st.download_button(
                "⬇️ Download JSON (dataset.json)", 
                data=json_str, 
                file_name="dataset.json", 
                mime="application/json", 
                type="primary"
            )

    # 底部保存逻辑
    if st.button("💾 Save & Next", type="primary", use_container_width=True, key="save_bottom"):
        
        # 内部暂时保留完整结构，导出时会自动清洗
        final_json = { 
            "id": current_id, 
            "prompt": main_prompt, 
            "remixSuggestions": current_remixes
        }
        
        st.session_state.processed_results[f"{current_id}.json"] = final_json
        if st.session_state.current_idx < len(st.session_state.data) - 1:
            st.session_state.current_idx += 1
            st.rerun()
        else:
            st.balloons()
            st.success("All Done!")
