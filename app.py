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
    "surface_color": "rgba(255, 255, 255, 0.85)", 
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
        .block-container {{padding-top: 2rem; padding-bottom: 5rem; max-width: 1000px;}}
        
        h1, h2, h3 {{ color: {tokens['text_primary']} !important; font-weight: 600 !important; }}
        
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

# ================= 2. DATA: NEW ELEMENT-LEVEL LISTS =================

# 🇬🇧 English List (Element Modifications)
REMIX_LIST_EN = [
    {"label": "Replace Background", "prompt": "Create this picture by replacing the background with a busy city street at twilight."},
    {"label": "Add Festive Props", "prompt": "Create this picture with festive holiday decorations added to the surrounding environment."},
    {"label": "Change Time of Day", "prompt": "Create this picture with golden hour lighting casting long, soft shadows across the scene."},
    {"label": "Replace Handheld Object", "prompt": "Create this picture by replacing the object in the hand with a vintage camera."},
    {"label": "Add Weather Effect", "prompt": "Create this picture with a gentle rain falling and reflections on the ground surfaces."},
    {"label": "Modify Clothing", "prompt": "Create this picture by replacing the outfit with a formal tuxedo and bow tie."},
    {"label": "Add Indoor Plants", "prompt": "Create this picture with lush green potted plants arranged in the background corners."},
    {"label": "Switch to Winter", "prompt": "Create this picture with a layer of snow covering the outdoor surfaces and frost on the windows."},
    {"label": "Add Glasses", "prompt": "Create this picture with a pair of stylish reading glasses added to the character's face."},
    {"label": "Replace Furniture", "prompt": "Create this picture by replacing the chair with a modern mid-century armchair."},
    {"label": "Change Hair Color", "prompt": "Create this picture by replacing the hair color with a vibrant platinum blonde tone."},
    {"label": "Add Neon Sign", "prompt": "Create this picture with a glowing neon sign mounted on the wall behind the subject."},
    {"label": "Add Pet Companion", "prompt": "Create this picture with a sleeping cat resting on the surface nearby."},
    {"label": "Replace Beverage", "prompt": "Create this picture by replacing the drink with a steaming cup of coffee in a ceramic mug."},
    {"label": "Adjust Lighting", "prompt": "Create this picture with dramatic spotlighting focusing solely on the central object."}
]

# 🇨🇳 Chinese List (Element Modifications)
REMIX_LIST_ZH = [
    {"label": "替换背景", "prompt": "通过将背景替换为黄昏时繁忙的城市街道来创作这张图片。"},
    {"label": "增加节日装饰", "prompt": "创作这张图片，在周围环境中添加节日的装饰道具。"},
    {"label": "更改时间光影", "prompt": "创作这张图片，使用黄金时刻的灯光，在场景中投下柔和的长影。"},
    {"label": "替换手中物体", "prompt": "通过将手中的物体替换为一台复古相机来创作这张图片。"},
    {"label": "增加天气效果", "prompt": "创作这张图片，添加柔和的雨丝和地面上的倒影效果。"},
    {"label": "修改服装", "prompt": "通过将服装替换为正式的燕尾服和领结来创作这张图片。"},
    {"label": "添加室内植物", "prompt": "创作这张图片，在背景角落布置郁郁葱葱的绿色盆栽。"},
    {"label": "切换至冬季", "prompt": "创作这张图片，让一层积雪覆盖室外表面，窗户上带有霜花。"},
    {"label": "添加眼镜", "prompt": "创作这张图片，给人物脸上加上一副时尚的阅读眼镜。"},
    {"label": "替换家具", "prompt": "通过将椅子替换为现代的中世纪风格扶手椅来创作这张图片。"},
    {"label": "改变发色", "prompt": "通过将头发颜色替换为充满活力的白金色调来创作这张图片。"},
    {"label": "添加霓虹灯牌", "prompt": "创作这张图片，在主体后方的墙上添加一个发光的霓虹灯牌。"},
    {"label": "添加宠物", "prompt": "创作这张图片，在附近的表面上添加一只熟睡的猫。"},
    {"label": "替换饮料", "prompt": "通过将饮料替换为陶瓷杯中热气腾腾的咖啡来创作这张图片。"},
    {"label": "调整灯光", "prompt": "创作这张图片，使用戏剧性的聚光灯仅聚焦于中心物体。"}
]

# 🔥 更新后的 Copilot 指令 (User Provided)
COPILOT_GEN_INSTRUCTION = """A remix prompt consists of a short, 2–5-word title and an instruction that begins with “Create this picture with…” or “Create this picture by replacing…”, focusing on element-level modifications rather than stylistic changes.

The title should concisely summarize the modification direction (e.g., Replace Object, Add Light Effects, Switch to Seasonal Props).
The prompt should be one clear, professional sentence that performs a specific element change in the original picture—such as replacing an object, adding a prop, adjusting lighting on a specific element, or modifying text—while keeping the original art style, mood, composition, and character features fully preserved.

Each prompt should describe 1–2 concrete element modifications, with optional clarity descriptors (e.g., material, scale, position), but must not introduce new stylistic transformations.
The remix may include slight creative enhancements as long as the result remains visually consistent with the original image and clearly derived from it.

After understanding this structure, please write 5 remix prompts for me based on the uploaded image (each including a modification title + a “Create this picture with/by…” instruction), all targeting different element-level changes.
Format:
Label: [Title]
Prompt: [Instruction]"""

def get_random_remix(language_mode):
    if language_mode == "Chinese/中文":
        return random.choice(REMIX_LIST_ZH)
    else:
        return random.choice(REMIX_LIST_EN)

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
st.set_page_config(page_title="Element Editor", layout="wide", page_icon="🧩")
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
        st.markdown(f"<div style='text-align: center;'><h1>🧩 Element Editor Studio</h1><p style='color:#594134'>Element-level Modifications Workflow</p></div>", unsafe_allow_html=True)
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

# --- Phase 2: Editor ---
else:
    with st.sidebar:
        st.markdown("### Settings")
        lang_mode = st.radio("Prompt Language / 语言", ["English", "Chinese/中文"], index=0)
        
        st.markdown("---")
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

    # Header
    col_h1, col_h2 = st.columns([3, 1])
    with col_h1:
        st.markdown(f"## ID {current_id} <span style='font-size:0.6em; color:#888'>Editor</span>", unsafe_allow_html=True)
    with col_h2:
        if st.button("💾 Save & Next", type="primary", use_container_width=True, key="save_top"):
            pass 

    # === Top: Image ===
    with st.container(border=True):
        if img_name in st.session_state.images:
            st.image(st.session_state.images[img_name], use_container_width=True)
        else:
            st.error("Missing Image")

    st.markdown("<br>", unsafe_allow_html=True)

    # === Middle: Main Prompt ===
    st.markdown("#### 📝 Main Prompt")
    default_text = item['original_prompt_text']
    if not default_text.strip().lower().startswith("create"):
        default_text = "Create an image of " + default_text
    main_prompt = st.text_area("main_hidden", value=default_text, height=100, key=f"m_{current_id}", label_visibility="collapsed")
    
    st.markdown("<br>", unsafe_allow_html=True)

    # === Bottom: Element Remix Suggestions ===
    st.markdown(f"#### 🧩 Element Modifications ({lang_mode})")
    
    session_key = f"remix_{current_id}"
    if session_key not in st.session_state:
        st.session_state[session_key] = [get_random_remix(lang_mode) for _ in range(3)]
        
    current_remixes = st.session_state[session_key]

    # Instruction Expander
    with st.expander("📋 Show Copilot Instruction (Copy & Paste)"):
        st.code(COPILOT_GEN_INSTRUCTION, language="text")
        st.caption("Instructions updated to focus on element-level changes.")

    # Grid Layout
    r_col1, r_col2, r_col3 = st.columns(3)
    cols = [r_col1, r_col2, r_col3]

    for i in range(3):
        with cols[i]:
            with st.container(border=True):
                c_title, c_btn = st.columns([4, 1])
                with c_title:
                    l_val = st.text_input(f"L{i}", value=current_remixes[i]['label'], key=f"l_{current_id}_{i}", label_visibility="collapsed", placeholder="Label")
                with c_btn:
                    if st.button("🎲", key=f"rnd_{current_id}_{i}", help=f"Randomize in {lang_mode}"):
                        st.session_state[session_key][i] = get_random_remix(lang_mode)
                        st.rerun()

                p_val = st.text_area(f"P{i}", value=current_remixes[i]['prompt'], height=120, key=f"p_{current_id}_{i}", label_visibility="collapsed", placeholder="Prompt")
                current_remixes[i]['label'] = l_val
                current_remixes[i]['prompt'] = p_val
                
                if st.button(f"🎨 Verify", key=f"v_{current_id}_{i}", use_container_width=True):
                    clean_prompt = urllib.parse.quote(p_val)
                    seed = random.randint(0, 9999)
                    url = f"https://image.pollinations.ai/prompt/{clean_prompt}?seed={seed}&width=600&height=600&nologo=true"
                    st.session_state[f"poll_img_{current_id}_{i}"] = url
                
                if f"poll_img_{current_id}_{i}" in st.session_state:
                    st.image(st.session_state[f"poll_img_{current_id}_{i}"], caption="Preview", use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    
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
