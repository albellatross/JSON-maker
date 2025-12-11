import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import json
import io
import zipfile
import random

# ================= 配置区 =================
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

# ================= 逻辑函数 =================

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

# ================= 🎨 UI 设置与 CSS 美化 =================

st.set_page_config(page_title="AI Dataset Studio", layout="wide", page_icon="✨")

# 注入自定义CSS，让界面更像一个专业软件
st.markdown("""
<style>
    /* 调整顶部留白 */
    .block-container { padding-top: 2rem; padding-bottom: 5rem; }
    
    /* 优化文本域样式 */
    .stTextArea textarea { 
        font-family: 'Inter', sans-serif;
        font-size: 15px; 
        line-height: 1.5;
        border-radius: 8px;
    }
    
    /* 卡片容器样式 */
    [data-testid="stVerticalBlock"] > [style*="flex-direction: column;"] > [data-testid="stVerticalBlock"] {
        background-color: #f8f9fa; /* 浅灰底色 */
        padding: 1rem;
        border-radius: 10px;
    }

    /* 按钮美化 */
    .stButton button {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.3s;
    }
    
    /* 侧边栏样式微调 */
    section[data-testid="stSidebar"] {
        background-color: #f0f2f6;
    }
</style>
""", unsafe_allow_html=True)

# 初始化 Session State
if 'data' not in st.session_state: st.session_state.data = []
if 'images' not in st.session_state: st.session_state.images = {}
if 'processed_results' not in st.session_state: st.session_state.processed_results = {}
if 'current_idx' not in st.session_state: st.session_state.current_idx = 0

# ================= 🚀 阶段 1: 欢迎页 / 上传页 =================
if not st.session_state.data:
    st.markdown("<h1 style='text-align: center;'>✨ AI Dataset Studio</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: grey;'>将你的 PPT 一键转化为高质量 AI 训练数据集</p>", unsafe_allow_html=True)
    st.markdown("---")

    col_u1, col_u2, col_u3 = st.columns([1, 2, 1])
    with col_u2:
        with st.container(border=True):
            st.subheader("📂 开始工作")
            uploaded_ppt = st.file_uploader("拖入 PPTX 文件", type=["pptx"], label_visibility="collapsed")
            start_id = st.number_input("设置起始 ID (Start ID)", value=453, step=1)
            
            if uploaded_ppt:
                if st.button("🚀 启动提取引擎", type="primary", use_container_width=True):
                    with st.spinner("正在逐页解析 PPT，提取高清图片与文本..."):
                        try:
                            data, images = process_ppt_file(uploaded_ppt, start_id)
                            st.session_state.data = data
                            st.session_state.images = images
                            st.session_state.current_idx = 0
                            st.rerun()
                        except Exception as e:
                            st.error(f"解析失败: {e}")

# ================= 🛠️ 阶段 2: 编辑工作台 =================
else:
    # --- 侧边栏：状态面板 ---
    with st.sidebar:
        st.header("📊 工作进度")
        total = len(st.session_state.data)
        done = len(st.session_state.processed_results)
        
        # 进度条
        st.progress(done / total if total > 0 else 0)
        st.caption(f"已完成: {done} / {total}")
        
        st.divider()
        st.subheader("📦 导出")
        if done > 0:
            zip_buffer = create_final_zip(st.session_state.processed_results, st.session_state.images)
            st.download_button(
                label="⬇️ 下载数据集 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="ai_dataset_pack.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )
        else:
            st.info("处理完第一张图后即可下载。")
        
        st.divider()
        if st.button("🔄 重置所有数据", type="secondary"):
            st.session_state.clear()
            st.rerun()

    # --- 主操作区 ---
    current_item = st.session_state.data[st.session_state.current_idx]
    current_id = current_item['id']
    img_name = current_item['image_filename']
    
    # 顶部导航
    c_nav1, c_nav2 = st.columns([3, 1])
    with c_nav1:
        st.subheader(f"🖼️ 编辑中: ID {current_id}")
    with c_nav2:
        st.caption(f"文件名: {img_name}")

    col_left, col_right = st.columns([1, 1.3])
    
    # === 左侧：图片预览 ===
    with col_left:
        # 使用 Container 给图片加个框
        with st.container(border=True):
            if img_name in st.session_state.images:
                st.image(st.session_state.images[img_name], use_container_width=True)
            else:
                st.error("图片数据丢失")

    # === 右侧：编辑表单 ===
    with col_right:
        # 1. 主 Prompt (使用 Expander 保持整洁，或者直接展示)
        st.markdown("##### 📝 主指令 (Main Prompt)")
        default_text = current_item['original_prompt_text']
        if not default_text.strip().lower().startswith("create"):
            default_text = "Create an image of " + default_text
            
        main_prompt = st.text_area("main_prompt", value=default_text, height=120, key=f"main_{current_id}", label_visibility="collapsed")
        
        st.markdown("---")
        st.markdown("##### 🎨 风格变奏 (Remix Suggestions)")
        
        # 2. Remix 区域 (卡片化设计)
        session_key = f"remix_{current_id}"
        if session_key not in st.session_state:
            st.session_state[session_key] = random.sample(REMIX_MASTER_LIST, 3)
        current_remixes = st.session_state[session_key]
        
        # 遍历 3 个建议，每个放进一个带边框的容器
        for i in range(3):
            with st.container(border=True):
                c_text, c_btn = st.columns([5, 1])
                with c_text:
                    # 使用 text_input 的 label 直接作为标题，减少 label 占用
                    l_val = st.text_input(f"Label {i+1}", value=current_remixes[i]['label'], key=f"l_{current_id}_{i}", label_visibility="collapsed", placeholder="输入风格标题...")
                    p_val = st.text_area(f"Prompt {i+1}", value=current_remixes[i]['prompt'], height=70, key=f"p_{current_id}_{i}", label_visibility="collapsed", placeholder="输入具体指令...")
                    # 更新数据
                    current_remixes[i]['label'] = l_val
                    current_remixes[i]['prompt'] = p_val
                with c_btn:
                    st.write("") 
                    st.write("") 
                    # 图标按钮
                    if st.button("🎲", key=f"btn_{current_id}_{i}", help="随机换一个风格"):
                        st.session_state[session_key][i] = get_random_remix()
                        st.rerun()

        # 底部大按钮
        st.markdown("")
        if st.button("💾 保存当前并继续 (Next)", type="primary", use_container_width=True):
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
                st.success("🎉 全部完成！请点击左侧下载按钮。")
