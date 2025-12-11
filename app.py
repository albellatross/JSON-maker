import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import json
import io
import zipfile
import random

# ================= 配置区：你的 Remix Master List =================
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

# ================= 核心逻辑：PPT处理与打包 =================

def process_ppt_file(uploaded_file, start_id):
    """直接在内存中处理PPT，不保存到硬盘"""
    prs = Presentation(uploaded_file)
    current_id = int(start_id)
    
    extracted_data = [] # 存放元数据
    image_storage = {}  # 存放图片二进制数据 { "453.png": bytes }

    for index, slide in enumerate(prs.slides):
        slide_info = {
            "id": str(current_id),
            "original_prompt_text": "",
            "image_filename": ""
        }
        
        found_image = False
        
        # 遍历形状
        for shape in slide.shapes:
            # A. 提取图片
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                if not found_image: # 每页只取第一张图
                    image_filename = f"{current_id}.png"
                    # 将图片保存在内存字典里
                    image_storage[image_filename] = shape.image.blob
                    slide_info["image_filename"] = image_filename
                    found_image = True
            
            # B. 提取文本
            if shape.has_text_frame:
                text = shape.text.strip()
                if len(text) > 10:
                    if len(text) > len(slide_info["original_prompt_text"]):
                        slide_info["original_prompt_text"] = text
        
        if found_image:
            extracted_data.append(slide_info)
            current_id += 1
            
    return extracted_data, image_storage

def create_final_zip(processed_jsons, image_storage):
    """打包所有 JSON 和 重命名后的图片"""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        # 1. 写入生成的 JSON 文件
        for filename, json_data in processed_jsons.items():
            json_str = json.dumps(json_data, indent=4, ensure_ascii=False)
            zip_file.writestr(f"jsons/{filename}", json_str)
            
        # 2. 写入对应的图片文件 (从内存写入Zip)
        # 只写入已经处理过（生成了JSON）的图片，节省体积
        for json_filename, json_data in processed_jsons.items():
            img_name = f"{json_data['id']}.png"
            if img_name in image_storage:
                zip_file.writestr(f"images/{img_name}", image_storage[img_name])
                
    return zip_buffer

def get_random_remix():
    return random.choice(REMIX_MASTER_LIST)

# ================= 网页 UI 部分 =================

st.set_page_config(page_title="One-Stop AI Tool", layout="wide", page_icon="⚡️")

st.markdown("""
<style>
    .stTextArea textarea { font-size: 14px; }
    .stButton button { width: 100%; border-radius: 8px; }
    .css-1v0mbdj.etr89bj1 { display: block; } 
</style>
""", unsafe_allow_html=True)

st.title("⚡️ PPT 转 AI 数据集 - 全能工作台")

# 初始化 Session State
if 'data' not in st.session_state:
    st.session_state.data = []         # PPT提取出的原始列表
if 'images' not in st.session_state:
    st.session_state.images = {}       # 图片二进制数据
if 'processed_results' not in st.session_state:
    st.session_state.processed_results = {} # 最终做好的JSON
if 'current_idx' not in st.session_state:
    st.session_state.current_idx = 0   # 当前做到第几张

# ---------------------------------------------------------
# 阶段 1: 如果没有数据，显示上传界面
# ---------------------------------------------------------
if not st.session_state.data:
    st.info("👋 欢迎！请直接上传你的 PPT 文件，我来帮你处理一切。")
    
    col_u1, col_u2 = st.columns([2, 1])
    with col_u1:
        uploaded_ppt = st.file_uploader("拖入 PPTX 文件", type=["pptx"])
    with col_u2:
        start_id = st.number_input("起始序号 (ID)", value=453, step=1)
    
    if uploaded_ppt:
        if st.button("🚀 开始提取素材", type="primary"):
            with st.spinner("正在拆解 PPT，提取图片和文字..."):
                try:
                    data, images = process_ppt_file(uploaded_ppt, start_id)
                    st.session_state.data = data
                    st.session_state.images = images
                    st.session_state.current_idx = 0
                    st.rerun() # 刷新页面进入阶段 2
                except Exception as e:
                    st.error(f"处理出错: {e}")

# ---------------------------------------------------------
# 阶段 2: 编辑与导出界面
# ---------------------------------------------------------
else:
    # 侧边栏：状态与导出
    with st.sidebar:
        st.header("📦 导出中心")
        total = len(st.session_state.data)
        done = len(st.session_state.processed_results)
        
        st.write(f"进度: {done} / {total}")
        st.progress(done / total if total > 0 else 0)
        
        if done > 0:
            st.success("已准备好下载")
            zip_buffer = create_final_zip(st.session_state.processed_results, st.session_state.images)
            st.download_button(
                label="⬇️ 下载最终 ZIP 包 (图+JSON)",
                data=zip_buffer.getvalue(),
                file_name="ai_dataset_pack.zip",
                mime="application/zip",
                type="primary"
            )
        
        st.divider()
        if st.button("🔄 重置/上传新文件"):
            st.session_state.clear()
            st.rerun()

    # 主编辑区
    current_item = st.session_state.data[st.session_state.current_idx]
    current_id = current_item['id']
    img_name = current_item['image_filename']
    
    col_left, col_right = st.columns([1, 1.3])
    
    # 左侧：看图
    with col_left:
        if img_name in st.session_state.images:
            st.image(st.session_state.images[img_name], caption=f"ID: {current_id}", use_container_width=True)
        else:
            st.error("图片丢失")

    # 右侧：编辑
    with col_right:
        st.subheader(f"编辑 ID: {current_id}")
        
        # 1. 主 Prompt 自动补全
        default_text = current_item['original_prompt_text']
        if not default_text.strip().lower().startswith("create"):
            default_text = "Create an image of " + default_text
            
        main_prompt = st.text_area("Main Prompt", value=default_text, height=100, key=f"main_{current_id}")
        
        st.markdown("---")
        st.write("**Remix 建议 (点击 🎲 随机更换)**")
        
        # 2. Remix 抽卡逻辑
        session_key = f"remix_{current_id}"
        if session_key not in st.session_state:
            st.session_state[session_key] = random.sample(REMIX_MASTER_LIST, 3)
        
        current_remixes = st.session_state[session_key]
        
        # 3 个卡片
        for i in range(3):
            c1, c2 = st.columns([6, 1])
            with c1:
                # 紧凑布局
                l_val = st.text_input(f"Label {i+1}", value=current_remixes[i]['label'], key=f"l_{current_id}_{i}", label_visibility="collapsed", placeholder="Label")
                p_val = st.text_area(f"Prompt {i+1}", value=current_remixes[i]['prompt'], height=60, key=f"p_{current_id}_{i}", label_visibility="collapsed", placeholder="Prompt")
                # 更新
                current_remixes[i]['label'] = l_val
                current_remixes[i]['prompt'] = p_val
            with c2:
                if st.button("🎲", key=f"btn_{current_id}_{i}"):
                    st.session_state[session_key][i] = get_random_remix()
                    st.rerun()
            st.write("") # 间隔

        # 底部保存
        if st.button("💾 保存并下一张", type="primary"):
            # 保存到 session
            final_json = {
                "id": current_id,
                "prompt": main_prompt,
                "remixSuggestions": current_remixes
            }
            st.session_state.processed_results[f"{current_id}.json"] = final_json
            
            # 自动跳页
            if st.session_state.current_idx < len(st.session_state.data) - 1:
                st.session_state.current_idx += 1
                st.rerun()
            else:
                st.balloons()
                st.success("全部完成！请在左侧下载。")
