import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import json
import io
import zipfile
import random
import urllib.parse
import time

# ================= 🎨 1. DESIGN TOKENS (Copilot Style) =================
MY_DESIGN_TOKENS = {
    "bg_color": "#FFF6F0",            
    "surface_color": "rgba(255, 255, 255, 0.70)", 
    "text_primary": "#311F10",        
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
        .block-container {{padding-top: 2rem; padding-bottom: 5rem; max-width: 1200px;}}
        h1, h2, h3 {{ color: {tokens['text_primary']} !important; font-weight: 600 !important; letter-spacing: -0.02em !important; }}
        
        [data-testid="stVerticalBlock"] > [style*="flex-direction: column;"] > [data-testid="stVerticalBlock"] {{
            background-color: {tokens['surface_color']};
            border-radius: {tokens['radius_card']};
            padding: 2rem;
            box-shadow: {tokens['shadow_tinted']};
            border: 1px solid rgba(255,255,255,0.5);
        }}
        
        .stTextArea textarea, .stTextInput input {{
            background-color: #FFFFFF !important;
            border-radius: 12px !important;
            border: 1px solid transparent !important;
            box-shadow: 0 2px 6px rgba(210, 150, 120, 0.05) !important;
        }}
        
        /* 隐藏顶部链接图标 */
        .css-1v0mbdj a {{ display: none; }}

        .stButton button {{ border-radius: {tokens['radius_pill']} !important; font-weight: 600 !important; border: none !important; }}
        div[data-testid="stButton"] > button[kind="primary"] {{ background-color: {tokens['accent_color']} !important; color: #FFFFFF !important; }}
        
        img {{ border-radius: 24px !important; }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

# ================= 2. LOGIC & DATA =================
REMIX_MASTER_LIST = [
    {"label": "Pixel Art", "prompt": "Create this picture as a retro pixel art, with nostalgic detail and game shading."},
    {"label": "Watercolor", "prompt": "Create this picture as a watercolor painting with soft edges and pastel tones."},
    {"label": "Cyberpunk", "prompt": "Create this picture in a cyberpunk style with neon lights and high contrast."},
    {"label": "Oil Painting", "prompt": "Create this picture as a classical oil painting with rich textures."},
    {"label": "3D Render", "prompt": "Create this picture as a cute 3D render style, like a toy."}
]

COPILOT_GEN_INSTRUCTION = """A remix prompt consists of a short style title and a style-transformation instruction that begins with “Create this picture as…”. 

Please look at the attached image and write 3 creative remix prompts for me strictly in this format:
Label: [Style Name]
Prompt: [Instruction]

Make sure the styles fit the image content."""

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

def get_random_remix():
    return random.choice(REMIX_MASTER_LIST)

# ================= 3. MAIN UI =================
st.set_page_config(page_title="Free Remix Studio", layout="wide", page_icon="🎁")
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
        st.markdown(f"<div style='text-align: center;'><h1>🎁 Free Remix Studio</h1><p style='color:#594134'>Zero Cost. No API Key Needed.</p></div>", unsafe_allow_html=True)
        with st.container(border=True):
            uploaded_ppt = st.file_uploader("Upload PPTX", type=["pptx"])
            start_id = st.number_input("Start ID", value=453, step=1)
            st.markdown("<br>", unsafe_allow_html=True)
            if uploaded_ppt:
                if st.button("🚀 Load Slides", type="primary", use_container_width=True):
                    with st.spinner("Analyzing PPT..."):
                        try:
                            data, images = process_ppt_file(uploaded_ppt, start_id)
                            st.session_state.data = data
                            st.session_state.images = images
                            st.session_state.current_idx = 0
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error: {e}")

# --- Phase 2: Editor (Free Mode) ---
else:
    # Sidebar
    with st.sidebar:
        st.markdown("### Progress")
        total = len(st.session_state.data)
        done = len(st.session_state.processed_results)
        st.progress(done / total if total > 0 else 0)
        st.caption(f"Done: {done} / {total}")
        
        if done > 0:
            zip_buffer = create_final_zip(st.session_state.processed_results, st.session_state.images)
            st.download_button("⬇️ Download ZIP", data=zip_buffer.getvalue(), file_name="dataset.zip", mime="application/zip", type="primary")
        
        st.markdown("---")
        if st.button("Reset Project", type="secondary"):
            st.session_state.clear()
            st.rerun()

    item = st.session_state.data[st.session_state.current_idx]
    current_id = item['id']
    img_name = item['image_filename']

    st.markdown(f"## Editor <span style='font-weight:400; font-size:0.8em; color:#594134;'>ID {current_id}</span>", unsafe_allow_html=True)

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
        st.markdown("**Main Prompt**")
        default_text = item['original_prompt_text']
        if not default_text.strip().lower().startswith("create"):
            default_text = "Create an image of " + default_text
        main_prompt = st.text_area("main", value=default_text, height=100, key=f"m_{current_id}", label_visibility="collapsed")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 2. Generation Method (Free Tabs)
        st.markdown("**Remix Generator**")
        
        session_key = f"remix_{current_id}"
        if session_key not in st.session_state:
            st.session_state[session_key] = random.sample(REMIX_MASTER_LIST, 3)
            
        current_remixes = st.session_state[session_key]
        
        # Tabs for Free Workflow
        tab1, tab2 = st.tabs(["🎲 Quick Random", "🧠 Copilot Manual"])
        
        with tab1:
            if st.button("🎲 Randomize (from List)", key=f"rnd_{current_id}", type="secondary"):
                st.session_state[session_key] = random.sample(REMIX_MASTER_LIST, 3)
                st.rerun()
                
        with tab2:
            st.info("💡 Copy instruction -> Paste to Copilot/ChatGPT -> Paste results below.")
            st.code(COPILOT_GEN_INSTRUCTION, language="text")
            st.markdown("[Open Copilot ↗️](https://copilot.microsoft.com/)")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("**Verify & Edit**")

        # 3. Cards with Free Image Gen
        for i in range(3):
            with st.container(border=True):
                # Inputs
                l_val = st.text_input(f"Label {i+1}", value=current_remixes[i]['label'], key=f"l_{current_id}_{i}", label_visibility="collapsed", placeholder="Label")
                p_val = st.text_area(f"Prompt {i+1}", value=current_remixes[i]['prompt'], height=60, key=f"p_{current_id}_{i}", label_visibility="collapsed", placeholder="Prompt")
                current_remixes[i]['label'] = l_val
                current_remixes[i]['prompt'] = p_val
                
                # Free Verification Button
                if st.button(f"🎨 Free Verify (Pollinations)", key=f"v_btn_{current_id}_{i}", type="secondary", use_container_width=True):
                    # 构建 Pollinations URL (无需 Key)
                    clean_prompt = urllib.parse.quote(p_val)
                    # 添加随机种子防止缓存
                    seed = random.randint(0, 10000)
                    image_url = f"https://image.pollinations.ai/prompt/{clean_prompt}?seed={seed}&width=800&height=800&nologo=true"
                    st.session_state[f"poll_img_{current_id}_{i}"] = image_url

                # Show Image
                if f"poll_img_{current_id}_{i}" in st.session_state:
                    st.image(st.session_state[f"poll_img_{current_id}_{i}"], caption="Preview", use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)
        # Save
        if st.button("💾 Save & Next", type="primary", use_container_width=True):
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
