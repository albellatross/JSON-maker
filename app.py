import streamlit as st
# 设置页面配置必须是第一个 Streamlit 命令
st.set_page_config(page_title="Remix Studio", layout="wide", page_icon="🧶")

import gc # 引入垃圾回收
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import json
import io
import zipfile
import random
import urllib.parse
import re
import base64

# ================= 🎨 1. DESIGN TOKENS =================
MY_DESIGN_TOKENS = {
    "bg_color": "#FFF6F0", 
    "surface_color": "rgba(255, 255, 255, 0.90)", 
    "text_primary": "#311F10",        
    "accent_color": "#311F10", 
    "radius_card": "16px",
    "radius_pill": "999px",
    "shadow_tinted": "0 2px 8px rgba(210, 150, 120, 0.08)",
    "font_family": "'Segoe UI', 'Microsoft YaHei', sans-serif"
}

def inject_layout_css(tokens):
    css = f"""
    <style>
        .stApp {{ background-color: {tokens['bg_color']}; font-family: {tokens['font_family']}; color: {tokens['text_primary']}; }}
        header, [data-testid="stHeader"] {{ display: none !important; }}
        .block-container {{
            padding: 1rem 2rem !important;
            max-width: 98% !important;
            margin-top: 0 !important;
        }}
        h1, h2, h3 {{ color: {tokens['text_primary']} !important; margin: 0 !important; }}
        
        .left-panel {{
            height: 85vh; 
            background-color: #EFEBE9; 
            border-radius: 16px;
            display: flex; justify-content: center; align_items: center;
            overflow: hidden; border: 1px solid rgba(0,0,0,0.05);
        }}
        .left-panel img {{
            max-width: 95%; max-height: 95%; width: auto; height: auto;
            object-fit: contain; box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }}
        .right-scroll-area {{
            height: 85vh; overflow-y: auto; padding-right: 12px; padding-left: 5px; padding-bottom: 60px;
        }}
        .right-scroll-area::-webkit-scrollbar {{ width: 6px; }}
        .right-scroll-area::-webkit-scrollbar-thumb {{ background-color: #D7CCC8; border-radius: 3px; }}
        [data-testid="stVerticalBlock"] > [style*="flex-direction: column;"] > [data-testid="stVerticalBlock"] {{
            background-color: {tokens['surface_color']};
            border-radius: {tokens['radius_card']};
            padding: 1.2rem;
            box-shadow: {tokens['shadow_tinted']};
            border: 1px solid rgba(255,255,255,0.6);
            margin-bottom: 1rem;
        }}
        .stTextArea textarea {{ font-size: 13px; min-height: 80px; }}
        .stTextInput input {{ font-size: 13px; }}
        .stButton button {{ border-radius: {tokens['radius_pill']} !important; font-weight: 600 !important; }}
        div[data-testid="stButton"] > button[kind="primary"] {{ 
            background-color: {tokens['accent_color']} !important; 
            color: #FFFFFF !important; 
            border: none !important;
        }}
        img {{ border-radius: 12px !important; }}
        .css-1v0mbdj a {{ display: none; }}
        .stProgress > div > div > div > div {{ background-color: {tokens['accent_color']}; }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

# ================= 2. DATA LISTS =================

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

COPILOT_GEN_INSTRUCTION = """A remix prompt consists of a short, 2–5-word title and an instruction.
Please write 5 remix prompts for me based on the uploaded image.
Format:
Label: [Title]
Prompt: [Instruction]"""

def get_random_remix(): return random.choice(REMIX_LIST_EN)

def randomize_callback(index, session_key_root, current_id_val):
    new_remix = get_random_remix()
    st.session_state[session_key_root][index] = new_remix
    st.session_state[f"l_{current_id_val}_{index}"] = new_remix['label']
    st.session_state[f"p_{current_id_val}_{index}"] = new_remix['prompt']

def parse_bulk_remix_text(raw_text):
    if not raw_text.strip(): return []
    # 关键词库
    ACTION_KEYWORDS = ("create", "change", "recreate", "replace", "generate", "make", "transform", "add", "switch", "use", "apply", "convert", "turn")
    
    lines = raw_text.split('\n')
    parsed_items = []
    processed_indices = set() 
    
    def clean_line_start(s): return re.sub(r'^[\d\.\-\*\s]+', '', s).strip()

    for i, line in enumerate(lines):
        clean_current = clean_line_start(line)
        lower_current = clean_current.lower()
        is_prompt_start = lower_current.startswith(ACTION_KEYWORDS)
        
        inline_split = line.split(":", 1)
        has_inline_title = len(inline_split) > 1 and clean_line_start(inline_split[1]).lower().startswith(ACTION_KEYWORDS)

        if is_prompt_start or has_inline_title:
            title = "Remix Option"
            prompt_text = ""
            if has_inline_title:
                title = clean_line_start(inline_split[0])
                prompt_text = clean_line_start(inline_split[1])
                processed_indices.add(i)
            else:
                prompt_text = clean_current
                processed_indices.add(i)
                k = i - 1
                while k >= 0:
                    prev_line = lines[k].strip()
                    if prev_line and k not in processed_indices:
                        title = clean_line_start(prev_line)
                        processed_indices.add(k)
                        break
                    k -= 1
            j = i + 1
            while j < len(lines):
                next_line = lines[j].strip()
                if not next_line: j+=1; continue
                clean_next = clean_line_start(next_line)
                if clean_next.lower().startswith(ACTION_KEYWORDS): break
                if j + 1 < len(lines):
                    clean_next_next = clean_line_start(lines[j+1])
                    if clean_next_next.lower().startswith(ACTION_KEYWORDS): break
                prompt_text += " " + next_line
                processed_indices.add(j)
                j += 1
            parsed_items.append({"label": title, "prompt": prompt_text})
    return parsed_items

def batch_parse_callback(session_key, current_id_val):
    batch_text = st.session_state.get("batch_input_area", "")
    parsed_items = parse_bulk_remix_text(batch_text)
    
    if parsed_items:
        final_items = parsed_items[:3]
        while len(final_items) < 3:
            final_items.append(get_random_remix())
        
        st.session_state[session_key] = final_items
        for idx, item in enumerate(final_items):
            st.session_state[f"l_{current_id_val}_{idx}"] = item['label']
            st.session_state[f"p_{current_id_val}_{idx}"] = item['prompt']
        
        st.session_state["batch_input_area"] = ""
        st.session_state["_parse_success"] = True
        st.session_state["_parsed_count"] = len(parsed_items)
    else:
        st.session_state["_parse_error"] = True

# --- File Operations (Memory Optimized) ---

def process_ppt_file(uploaded_file, start_id):
    # 强制清理内存
    gc.collect()
    
    try:
        uploaded_file.seek(0)
        prs = Presentation(uploaded_file)
    except zipfile.BadZipFile:
        raise ValueError("File is not a valid .pptx file.")
    except Exception as e:
        raise ValueError(f"Error reading PPT: {str(e)}")

    current_id = int(start_id)
    extracted_data = []
    image_storage = {}
    
    # 限制处理数量防止崩溃
    MAX_SLIDES = 100 
    
    for index, slide in enumerate(prs.slides):
        if index >= MAX_SLIDES:
            st.warning(f"⚠️ Limit reached: Processing first {MAX_SLIDES} slides to prevent memory crash.")
            break
            
        slide_info = {"id": str(current_id), "original_prompt_text": "", "image_filename": ""}
        found_image = False
        
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and not found_image:
                img_name = f"{current_id}.png"
                # 直接读取二进制，不进行图像处理以节省内存
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
    
    gc.collect() # 再次清理
    return extracted_data, image_storage

def create_final_zip(processed_jsons, image_storage):
    gc.collect()
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        combined_list = []
        sorted_keys = sorted(processed_jsons.keys(), key=lambda x: int(processed_jsons[x]['id']))
        for key in sorted_keys:
            item = processed_jsons[key]
            clean_item = {
                "id": item.get("id"),
                "prompt": item.get("prompt"),
                "remixSuggestions": [
                    {"label": r.get("label"), "prompt": r.get("prompt")} for r in item.get("remixSuggestions", [])
                ]
            }
            combined_list.append(clean_item)
            target_img_name = f"{item.get('id')}.png"
            if target_img_name in image_storage:
                zip_file.writestr(f"images/{target_img_name}", image_storage[target_img_name])
        json_str = json.dumps(combined_list, indent=4, ensure_ascii=False)
        zip_file.writestr("dataset.json", json_str)
    
    gc.collect()
    return zip_buffer

# ================= 3. MAIN UI =================
inject_layout_css(MY_DESIGN_TOKENS)

if 'data' not in st.session_state: st.session_state.data = []
if 'images' not in st.session_state: st.session_state.images = {}
if 'processed_results' not in st.session_state: st.session_state.processed_results = {}
if 'current_idx' not in st.session_state: st.session_state.current_idx = 0

# --- Phase 1: Upload ---
if not st.session_state.data:
    st.markdown("<br><br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown(f"<div style='text-align: center; margin-bottom:20px;'><h1>🧶 Remix Studio</h1><p style='color:#64748B'>Stable Version</p></div>", unsafe_allow_html=True)
        with st.container(border=True):
            uploaded_ppt = st.file_uploader("Upload Presentation", type=["pptx"])
            start_id = st.number_input("Start ID", value=453, step=1)
            st.markdown("<br>", unsafe_allow_html=True)
            if uploaded_ppt:
                if st.button("🚀 Load Slides", type="primary", use_container_width=True):
                    with st.spinner("Processing..."):
                        try:
                            data, images = process_ppt_file(uploaded_ppt, start_id)
                            if not data: st.error("No valid slides found.")
                            else:
                                st.session_state.data = data
                                st.session_state.images = images
                                st.session_state.current_idx = 0
                                st.rerun()
                        except ValueError as ve: st.error(str(ve))
                        except Exception as e: st.error(f"Error: {e}")

# --- Phase 2: Editor ---
else:
    item = st.session_state.data[st.session_state.current_idx]
    current_id = item['id']
    img_name = item['image_filename']

    # 1. Top Bar
    c1, c2, c3 = st.columns([1, 4, 1], gap="small")
    with c1:
        st.markdown(f"### ID {current_id}")
    with c2:
        total = len(st.session_state.data)
        done = len(st.session_state.processed_results)
        st.progress(done / total if total > 0 else 0)
    with c3:
        if st.button("💾 Next Image", type="primary", use_container_width=True):
            session_key = f"remix_{current_id}"
            current_remixes = st.session_state.get(session_key, [])
            main_p = st.session_state.get(f"m_{current_id}", "")
            
            final_json = { 
                "id": current_id, 
                "prompt": main_p, 
                "remixSuggestions": current_remixes
            }
            st.session_state.processed_results[f"{current_id}.json"] = final_json
            if st.session_state.current_idx < len(st.session_state.data) - 1:
                st.session_state.current_idx += 1
                st.rerun()
            else:
                st.balloons()
                st.success("All Done!")

    st.markdown("<hr style='margin: 0.5rem 0; opacity: 0.2;'>", unsafe_allow_html=True)

    # 2. Main Split View
    col_left, col_right = st.columns([2, 3], gap="medium")

    # === LEFT: Image ===
    with col_left:
        if img_name in st.session_state.images:
            b64_img = base64.b64encode(st.session_state.images[img_name]).decode()
            st.markdown(f"""
            <div class="left-panel">
                <img src="data:image/png;base64,{b64_img}" />
            </div>
            """, unsafe_allow_html=True)
        else:
            st.error("Image missing")

    # === RIGHT: Editor ===
    with col_right:
        st.markdown('<div class="right-scroll-area">', unsafe_allow_html=True)

        st.markdown("#### 📝 Main Prompt")
        default_text = item['original_prompt_text']
        if not default_text.strip().lower().startswith("create"):
            default_text = "Create an image of " + default_text
        main_prompt = st.text_area("main_hidden", value=default_text, height=80, key=f"m_{current_id}", label_visibility="collapsed")

        st.markdown("---")

        with st.expander("📋 Paste Remix Text (Replace)", expanded=False):
            st.text_area("Paste here", height=100, key="batch_input_area", label_visibility="collapsed", placeholder="Title\nCreate...")
            
            session_key = f"remix_{current_id}"
            st.button("Parse & Replace", on_click=batch_parse_callback, args=(session_key, current_id))
            
            if st.session_state.get("_parse_success"):
                st.success(f"Updated {st.session_state['_parsed_count']} items!")
                st.session_state["_parse_success"] = False
            if st.session_state.get("_parse_error"):
                st.warning("No valid prompts found.")
                st.session_state["_parse_error"] = False

        st.markdown("#### 🎨 Remix Suggestions")
        if session_key not in st.session_state:
            st.session_state[session_key] = [get_random_remix() for _ in range(3)]
        current_remixes = st.session_state[session_key]

        with st.expander("Show Instruction"):
            st.code(COPILOT_GEN_INSTRUCTION, language="text")

        for i in range(3):
            with st.container(border=True):
                c_t, c_b = st.columns([5, 1])
                with c_t:
                    l_key = f"l_{current_id}_{i}"
                    if l_key not in st.session_state: st.session_state[l_key] = current_remixes[i]['label']
                    l_val = st.text_input(f"L{i}", value=current_remixes[i]['label'], key=l_key, label_visibility="collapsed", placeholder="Label")
                with c_b:
                    st.button("🎲", key=f"rnd_{current_id}_{i}", on_click=randomize_callback, args=(i, session_key, current_id))

                p_key = f"p_{current_id}_{i}"
                if p_key not in st.session_state: st.session_state[p_key] = current_remixes[i]['prompt']
                p_val = st.text_area(f"P{i}", value=current_remixes[i]['prompt'], height=80, key=p_key, label_visibility="collapsed", placeholder="Prompt")

                if current_remixes[i]['label'] != l_val: current_remixes[i]['label'] = l_val
                if current_remixes[i]['prompt'] != p_val: current_remixes[i]['prompt'] = p_val

                if st.button("Verify", key=f"v_{current_id}_{i}", use_container_width=True):
                    clean = urllib.parse.quote(p_val)
                    seed = random.randint(0, 9999)
                    url = f"https://image.pollinations.ai/prompt/{clean}?seed={seed}&width=600&height=600&nologo=true"
                    st.session_state[f"poll_img_{current_id}_{i}"] = url
                
                if f"poll_img_{current_id}_{i}" in st.session_state:
                    st.image(st.session_state[f"poll_img_{current_id}_{i}"], use_container_width=True)

        st.markdown('</div>', unsafe_allow_html=True) # End scrollable

        st.markdown("---")
        with st.expander("⚙️ Download Data", expanded=False):
            if done > 0:
                zip_buffer = create_final_zip(st.session_state.processed_results, st.session_state.images)
                st.download_button("⬇️ Download ZIP", data=zip_buffer.getvalue(), file_name="dataset.zip", mime="application/zip", type="primary", use_container_width=True)
            else:
                st.info("Process at least one image to download.")
