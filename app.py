{\rtf1\ansi\ansicpg1252\cocoartf2867
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\paperw11900\paperh16840\margl1440\margr1440\vieww34020\viewh19880\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 import streamlit as st\
from pptx import Presentation\
from pptx.enum.shapes import MSO_SHAPE_TYPE\
import json\
import io\
import zipfile\
import random\
\
# ================= \uc0\u37197 \u32622 \u21306 \u65306 \u20320 \u30340  Remix Master List =================\
REMIX_MASTER_LIST = [\
    \{"label": "Want a wider view?", "prompt": "Create an expanded image with extended space."\},\
    \{"label": "Want to zoom in?", "prompt": "Create a micro-detail close-up variant of this image."\},\
    \{"label": "Try paper cut style?", "prompt": "Remake this image in a modern paper cut style with layered colors and soft shadows."\},\
    \{"label": "Make this embroidery style?", "prompt": "Remake this image in a textile embroidery style with visible stitched threads."\},\
    \{"label": "Change to Pixel Art", "prompt": "Create this picture as a retro pixel art, with nostalgic detail and game shading."\},\
    \{"label": "Apply Glitch Effect", "prompt": "Remake this image as a glitch digital art, with pixel splits and cyberpunk noise."\},\
    \{"label": "Change to Watercolor", "prompt": "Create this picture as a watercolor painting."\},\
    \{"label": "Change to Impressionism", "prompt": "Create this picture as an Impressionist painting, with loose brushwork, luminous color, and fleeting light."\},\
    \{"label": "Draw with colored pencil?", "prompt": "Remake this image as a colored pencil drawing."\},\
    \{"label": "Try fine-line style?", "prompt": "Remake this image as a Chinese Gongbi painting with precise outlines, soft washes, and detailed forms."\},\
    \{"label": "Try Chinese paper cut style?", "prompt": "Remake this image as a Chinese paper cut, with red silhouettes, cultural motifs, and symmetrical patterns."\},\
    \{"label": "Try Ukiyo-e style?", "prompt": "Remake this image as a Japanese Ukiyo-e, with woodblock texture, flat colors, and flowing lines."\},\
    \{"label": "Make this a portrait?", "prompt": "Remake this image as a photo portrait, with natural light, and shallow depth."\},\
    \{"label": "Make this stained glass?", "prompt": "Remake this image as a stained glass design with colorful panes, bold outlines, and glowing light."\},\
    \{"label": "Try silkscreen style?", "prompt": "Remake this image as a silkscreen print."\},\
    \{"label": "Make this anime?", "prompt": "Remake this image as an anime illustration with expressive light and a dynamic layout."\},\
    \{"label": "Add sepia tone?", "prompt": "Remake this image as a sepia-toned memory with aged paper texture."\},\
    \{"label": "Make this pop art?", "prompt": "Remake this image as a high-saturation pop art, with bold blocks and hues."\},\
    \{"label": "Make this a gradient mesh?", "prompt": "Remake this image as a gradient mesh, blending colors seamlessly across the composition."\},\
    \{"label": "Make this a 3D figure?", "prompt": "Remake this image as a photorealistic 3D render of a collectible figure, made of real materials like resin or plastic with cinematic lighting, studio backdrop, and ultra-fine modeling detail."\},\
    \{"label": "Try duotone colors?", "prompt": "Remake this image as a duotone image."\},\
    \{"label": "Make this monochrome?", "prompt": "Remake this image as a monochrome image."\},\
    \{"label": "Add neon lighting?", "prompt": "Remake this image as a neon-lit scene with vibrant color contrasts."\},\
    \{"label": "Make this mechanical?", "prompt": "Create a mechanical version of the subject with exposed gears, metallic joints, and precise components."\},\
    \{"label": "Make this crystal?", "prompt": "Remake this image to be in an iridescent fantasy realm with the subject as translucent glass or crystal, glowing and refracted."\}\
]\
\
# ================= \uc0\u26680 \u24515 \u36923 \u36753 \u65306 PPT\u22788 \u29702 \u19982 \u25171 \u21253  =================\
\
def process_ppt_file(uploaded_file, start_id):\
    """\uc0\u30452 \u25509 \u22312 \u20869 \u23384 \u20013 \u22788 \u29702 PPT\u65292 \u19981 \u20445 \u23384 \u21040 \u30828 \u30424 """\
    prs = Presentation(uploaded_file)\
    current_id = int(start_id)\
    \
    extracted_data = [] # \uc0\u23384 \u25918 \u20803 \u25968 \u25454 \
    image_storage = \{\}  # \uc0\u23384 \u25918 \u22270 \u29255 \u20108 \u36827 \u21046 \u25968 \u25454  \{ "453.png": bytes \}\
\
    for index, slide in enumerate(prs.slides):\
        slide_info = \{\
            "id": str(current_id),\
            "original_prompt_text": "",\
            "image_filename": ""\
        \}\
        \
        found_image = False\
        \
        # \uc0\u36941 \u21382 \u24418 \u29366 \
        for shape in slide.shapes:\
            # A. \uc0\u25552 \u21462 \u22270 \u29255 \
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:\
                if not found_image: # \uc0\u27599 \u39029 \u21482 \u21462 \u31532 \u19968 \u24352 \u22270 \
                    image_filename = f"\{current_id\}.png"\
                    # \uc0\u23558 \u22270 \u29255 \u20445 \u23384 \u22312 \u20869 \u23384 \u23383 \u20856 \u37324 \
                    image_storage[image_filename] = shape.image.blob\
                    slide_info["image_filename"] = image_filename\
                    found_image = True\
            \
            # B. \uc0\u25552 \u21462 \u25991 \u26412 \
            if shape.has_text_frame:\
                text = shape.text.strip()\
                if len(text) > 10:\
                    if len(text) > len(slide_info["original_prompt_text"]):\
                        slide_info["original_prompt_text"] = text\
        \
        if found_image:\
            extracted_data.append(slide_info)\
            current_id += 1\
            \
    return extracted_data, image_storage\
\
def create_final_zip(processed_jsons, image_storage):\
    """\uc0\u25171 \u21253 \u25152 \u26377  JSON \u21644  \u37325 \u21629 \u21517 \u21518 \u30340 \u22270 \u29255 """\
    zip_buffer = io.BytesIO()\
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:\
        # 1. \uc0\u20889 \u20837 \u29983 \u25104 \u30340  JSON \u25991 \u20214 \
        for filename, json_data in processed_jsons.items():\
            json_str = json.dumps(json_data, indent=4, ensure_ascii=False)\
            zip_file.writestr(f"jsons/\{filename\}", json_str)\
            \
        # 2. \uc0\u20889 \u20837 \u23545 \u24212 \u30340 \u22270 \u29255 \u25991 \u20214  (\u20174 \u20869 \u23384 \u20889 \u20837 Zip)\
        # \uc0\u21482 \u20889 \u20837 \u24050 \u32463 \u22788 \u29702 \u36807 \u65288 \u29983 \u25104 \u20102 JSON\u65289 \u30340 \u22270 \u29255 \u65292 \u33410 \u30465 \u20307 \u31215 \
        for json_filename, json_data in processed_jsons.items():\
            img_name = f"\{json_data['id']\}.png"\
            if img_name in image_storage:\
                zip_file.writestr(f"images/\{img_name\}", image_storage[img_name])\
                \
    return zip_buffer\
\
def get_random_remix():\
    return random.choice(REMIX_MASTER_LIST)\
\
# ================= \uc0\u32593 \u39029  UI \u37096 \u20998  =================\
\
st.set_page_config(page_title="One-Stop AI Tool", layout="wide", page_icon="\uc0\u9889 \u65039 ")\
\
st.markdown("""\
<style>\
    .stTextArea textarea \{ font-size: 14px; \}\
    .stButton button \{ width: 100%; border-radius: 8px; \}\
    .css-1v0mbdj.etr89bj1 \{ display: block; \} \
</style>\
""", unsafe_allow_html=True)\
\
st.title("\uc0\u9889 \u65039  PPT \u36716  AI \u25968 \u25454 \u38598  - \u20840 \u33021 \u24037 \u20316 \u21488 ")\
\
# \uc0\u21021 \u22987 \u21270  Session State\
if 'data' not in st.session_state:\
    st.session_state.data = []         # PPT\uc0\u25552 \u21462 \u20986 \u30340 \u21407 \u22987 \u21015 \u34920 \
if 'images' not in st.session_state:\
    st.session_state.images = \{\}       # \uc0\u22270 \u29255 \u20108 \u36827 \u21046 \u25968 \u25454 \
if 'processed_results' not in st.session_state:\
    st.session_state.processed_results = \{\} # \uc0\u26368 \u32456 \u20570 \u22909 \u30340 JSON\
if 'current_idx' not in st.session_state:\
    st.session_state.current_idx = 0   # \uc0\u24403 \u21069 \u20570 \u21040 \u31532 \u20960 \u24352 \
\
# ---------------------------------------------------------\
# \uc0\u38454 \u27573  1: \u22914 \u26524 \u27809 \u26377 \u25968 \u25454 \u65292 \u26174 \u31034 \u19978 \u20256 \u30028 \u38754 \
# ---------------------------------------------------------\
if not st.session_state.data:\
    st.info("\uc0\u55357 \u56395  \u27426 \u36814 \u65281 \u35831 \u30452 \u25509 \u19978 \u20256 \u20320 \u30340  PPT \u25991 \u20214 \u65292 \u25105 \u26469 \u24110 \u20320 \u22788 \u29702 \u19968 \u20999 \u12290 ")\
    \
    col_u1, col_u2 = st.columns([2, 1])\
    with col_u1:\
        uploaded_ppt = st.file_uploader("\uc0\u25302 \u20837  PPTX \u25991 \u20214 ", type=["pptx"])\
    with col_u2:\
        start_id = st.number_input("\uc0\u36215 \u22987 \u24207 \u21495  (ID)", value=453, step=1)\
    \
    if uploaded_ppt:\
        if st.button("\uc0\u55357 \u56960  \u24320 \u22987 \u25552 \u21462 \u32032 \u26448 ", type="primary"):\
            with st.spinner("\uc0\u27491 \u22312 \u25286 \u35299  PPT\u65292 \u25552 \u21462 \u22270 \u29255 \u21644 \u25991 \u23383 ..."):\
                try:\
                    data, images = process_ppt_file(uploaded_ppt, start_id)\
                    st.session_state.data = data\
                    st.session_state.images = images\
                    st.session_state.current_idx = 0\
                    st.rerun() # \uc0\u21047 \u26032 \u39029 \u38754 \u36827 \u20837 \u38454 \u27573  2\
                except Exception as e:\
                    st.error(f"\uc0\u22788 \u29702 \u20986 \u38169 : \{e\}")\
\
# ---------------------------------------------------------\
# \uc0\u38454 \u27573  2: \u32534 \u36753 \u19982 \u23548 \u20986 \u30028 \u38754 \
# ---------------------------------------------------------\
else:\
    # \uc0\u20391 \u36793 \u26639 \u65306 \u29366 \u24577 \u19982 \u23548 \u20986 \
    with st.sidebar:\
        st.header("\uc0\u55357 \u56550  \u23548 \u20986 \u20013 \u24515 ")\
        total = len(st.session_state.data)\
        done = len(st.session_state.processed_results)\
        \
        st.write(f"\uc0\u36827 \u24230 : \{done\} / \{total\}")\
        st.progress(done / total if total > 0 else 0)\
        \
        if done > 0:\
            st.success("\uc0\u24050 \u20934 \u22791 \u22909 \u19979 \u36733 ")\
            zip_buffer = create_final_zip(st.session_state.processed_results, st.session_state.images)\
            st.download_button(\
                label="\uc0\u11015 \u65039  \u19979 \u36733 \u26368 \u32456  ZIP \u21253  (\u22270 +JSON)",\
                data=zip_buffer.getvalue(),\
                file_name="ai_dataset_pack.zip",\
                mime="application/zip",\
                type="primary"\
            )\
        \
        st.divider()\
        if st.button("\uc0\u55357 \u56580  \u37325 \u32622 /\u19978 \u20256 \u26032 \u25991 \u20214 "):\
            st.session_state.clear()\
            st.rerun()\
\
    # \uc0\u20027 \u32534 \u36753 \u21306 \
    current_item = st.session_state.data[st.session_state.current_idx]\
    current_id = current_item['id']\
    img_name = current_item['image_filename']\
    \
    col_left, col_right = st.columns([1, 1.3])\
    \
    # \uc0\u24038 \u20391 \u65306 \u30475 \u22270 \
    with col_left:\
        if img_name in st.session_state.images:\
            st.image(st.session_state.images[img_name], caption=f"ID: \{current_id\}", use_container_width=True)\
        else:\
            st.error("\uc0\u22270 \u29255 \u20002 \u22833 ")\
\
    # \uc0\u21491 \u20391 \u65306 \u32534 \u36753 \
    with col_right:\
        st.subheader(f"\uc0\u32534 \u36753  ID: \{current_id\}")\
        \
        # 1. \uc0\u20027  Prompt \u33258 \u21160 \u34917 \u20840 \
        default_text = current_item['original_prompt_text']\
        if not default_text.strip().lower().startswith("create"):\
            default_text = "Create an image of " + default_text\
            \
        main_prompt = st.text_area("Main Prompt", value=default_text, height=100, key=f"main_\{current_id\}")\
        \
        st.markdown("---")\
        st.write("**Remix \uc0\u24314 \u35758  (\u28857 \u20987  \u55356 \u57266  \u38543 \u26426 \u26356 \u25442 )**")\
        \
        # 2. Remix \uc0\u25277 \u21345 \u36923 \u36753 \
        session_key = f"remix_\{current_id\}"\
        if session_key not in st.session_state:\
            st.session_state[session_key] = random.sample(REMIX_MASTER_LIST, 3)\
        \
        current_remixes = st.session_state[session_key]\
        \
        # 3 \uc0\u20010 \u21345 \u29255 \
        for i in range(3):\
            c1, c2 = st.columns([6, 1])\
            with c1:\
                # \uc0\u32039 \u20945 \u24067 \u23616 \
                l_val = st.text_input(f"Label \{i+1\}", value=current_remixes[i]['label'], key=f"l_\{current_id\}_\{i\}", label_visibility="collapsed", placeholder="Label")\
                p_val = st.text_area(f"Prompt \{i+1\}", value=current_remixes[i]['prompt'], height=60, key=f"p_\{current_id\}_\{i\}", label_visibility="collapsed", placeholder="Prompt")\
                # \uc0\u26356 \u26032 \
                current_remixes[i]['label'] = l_val\
                current_remixes[i]['prompt'] = p_val\
            with c2:\
                if st.button("\uc0\u55356 \u57266 ", key=f"btn_\{current_id\}_\{i\}"):\
                    st.session_state[session_key][i] = get_random_remix()\
                    st.rerun()\
            st.write("") # \uc0\u38388 \u38548 \
\
        # \uc0\u24213 \u37096 \u20445 \u23384 \
        if st.button("\uc0\u55357 \u56510  \u20445 \u23384 \u24182 \u19979 \u19968 \u24352 ", type="primary"):\
            # \uc0\u20445 \u23384 \u21040  session\
            final_json = \{\
                "id": current_id,\
                "prompt": main_prompt,\
                "remixSuggestions": current_remixes\
            \}\
            st.session_state.processed_results[f"\{current_id\}.json"] = final_json\
            \
            # \uc0\u33258 \u21160 \u36339 \u39029 \
            if st.session_state.current_idx < len(st.session_state.data) - 1:\
                st.session_state.current_idx += 1\
                st.rerun()\
            else:\
                st.balloons()\
                st.success("\uc0\u20840 \u37096 \u23436 \u25104 \u65281 \u35831 \u22312 \u24038 \u20391 \u19979 \u36733 \u12290 ")}