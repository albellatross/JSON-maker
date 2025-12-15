import streamlit as st
import io
import zipfile
import json
import base64
import random

# ================= 1. 核心配置与样式 (CSS Tokens) =================
STYLING = {
    "bg_main": "#FFF8F3",         # 主背景色（暖米色）
    "bg_card": "#FFFFFF",         # 卡片背景（纯白）
    "bg_left_panel": "#F2EBE6",   # 左侧图片区背景
    "text_dark": "#4A3A32",       # 主要文字颜色（深棕）
    "primary_btn": "#D35F5F",     # 主按钮颜色（暖红/棕红）
    "secondary_btn": "#ECE0D8",   # 次要按钮/边框颜色
}

# 注入自定义 CSS，实现完美的左右布局
def inject_custom_css():
    st.markdown(f"""
        <style>
            /* 全局设置 */
            .stApp {{
                background-color: {STYLING["bg_main"]};
                color: {STYLING["text_dark"]};
            }}
            /* 隐藏顶部 Header */
            header[data-testid="stHeader"] {{ display: none; }}
            .block-container {{ padding-top: 1rem; }}

            /* === 左侧面板样式 === */
            .left-image-container {{
                background-color: {STYLING["bg_left_panel"]};
                border-radius: 16px;
                padding: 20px;
                height: 85vh; /* 固定高度 */
                display: flex;
                justify-content: center;
                align-items: center;
                border: 2px solid {STYLING["secondary_btn"]};
            }}
            .left-image-container img {{
                max-height: 100%;
                max-width: 100%;
                object-fit: contain;
                border-radius: 8px;
                box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            }}

            /* === 右侧滚动区域样式 === */
            .right-scroll-container {{
                height: 85vh; /* 与左侧同高 */
                overflow-y: auto; /* 启用垂直滚动 */
                padding-right: 15px; /* 给滚动条留位置 */
                padding-left: 5px;
            }}
            /* 自定义滚动条 */
            .right-scroll-container::-webkit-scrollbar {{ width: 8px; }}
            .right-scroll-container::-webkit-scrollbar-track {{ background: transparent; }}
            .right-scroll-container::-webkit-scrollbar-thumb {{ background-color: #D0C0B4; border-radius: 4px; }}

            /* === 组件通用样式 === */
            .stTextArea textarea, .stTextInput input {{
                border-radius: 10px;
                border: 1px solid {STYLING["secondary_btn"]};
            }}
            /* 主按钮样式 (Save & Next) */
            div[data-testid="stButton"] > button[kind="primary"] {{
                background-color: {STYLING["primary_btn"]};
                border: none;
                border-radius: 20px;
                padding: 0.5rem 1rem;
                font-weight: 600;
                width: 100%;
            }}
             /* 普通按钮样式 (Verify/Apply) */
            div[data-testid="stButton"] > button[kind="secondary"] {{
                 background-color: {STYLING["secondary_btn"]};
                 border: none;
                 border-radius: 20px;
                 color: {STYLING["text_dark"]};
                 font-weight: 600;
            }}
            /* 进度条颜色 */
            .stProgress > div > div {{ background-color: {STYLING["primary_btn"]}; }}
            
            /* 卡片容器样式 */
            [data-testid="stVerticalBlockBorderWrapper"] > div {{
                background-color: {STYLING["bg_card"]};
                border-radius: 16px;
                border: 1px solid {STYLING["secondary_btn"]};
                box-shadow: 0 2px 6px rgba(0,0,0,0.04);
            }}

        </style>
    """, unsafe_allow_html=True)

# ================= 2. 初始化 Session State (修复 NameError) =================
if 'data' not in st.session_state: st.session_state.data = []
if 'images' not in st.session_state: st.session_state.images = {}
if 'processed_results' not in st.session_state: st.session_state.processed_results = {}
if 'current_idx' not in st.session_state: st.session_state.current_idx = 0
# 修复 APIException 的关键：不要在这里初始化 batch_input_area

# ================= 3. 辅助函数 (数据处理) =================
# 模拟 PPTX 处理 (为了演示，这里用占位符。你需要替换回你真实的PPTX处理逻辑)
def process_ppt_file_mock(uploaded_file, start_id):
    # 这里应该用 python-pptx 读取文件
    # 为了代码可运行，我创建一些假数据
    mock_data = []
    mock_images = {}
    curr_id = start_id
    for i in range(5): # 假设读取了5张图
        img_name = f"{curr_id}.png"
        # 创建一个假图片 (1x1 像素红色点) 用于演示
        mock_images[img_name] = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==")
        mock_data.append({
            "id": str(curr_id),
            "original_prompt_text": f"This is the original prompt for image {curr_id}. It contains some details about the scene.",
            "image_filename": img_name
        })
        curr_id += 1
    return mock_data, mock_images

def create_final_zip(processed_jsons, image_storage):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        final_list = []
        # 按ID排序
        sorted_ids = sorted(processed_jsons.keys(), key=lambda x: int(x))
        
        for pid in sorted_ids:
            item_data = processed_jsons[pid]
            final_list.append(item_data)
            # 将对应的图片加入ZIP
            img_name = f"{pid}.png"
            if img_name in image_storage:
                zip_file.writestr(f"images/{img_name}", image_storage[img_name])
                
        # 写入最终的 JSON 文件
        json_str = json.dumps(final_list, indent=4, ensure_ascii=False)
        zip_file.writestr("dataset.json", json_str)
    return zip_buffer

# ================= 4. 回调函数 (修复 StreamlitAPIException) =================
def on_batch_apply():
    """处理批量粘贴文本的回调函数。安全地更新 State。"""
    raw_text = st.session_state.get("batch_input_widget", "").strip()
    if not raw_text:
        st.warning("Cannot define batch_input_area because it's empty.")
        return

    # 简单的解析逻辑 (你可以替换为你更复杂的正则解析)
    new_remixes = []
    for line in raw_text.split('\n'):
        if ':' in line:
            parts = line.split(':', 1)
            new_remixes.append({"label": parts[0].strip(), "prompt": parts[1].strip()})
        elif line.strip():
             new_remixes.append({"label": "Remix", "prompt": line.strip()})

    if new_remixes:
        # 获取当前ID
        current_id = st.session_state.data[st.session_state.current_idx]['id']
        # 更新当前页面的 Remix Suggestions
        st.session_state[f"remix_{current_id}"] = new_remixes
        # 清空输入框 (通过设置 widget 的 key 对应的值)
        st.session_state["batch_input_widget"] = ""
        st.success(f"Successfully applied {len(new_remixes)} remix suggestions!")
    else:
        st.error("Could not parse any valid suggestions.")

def on_save_and_next():
    """保存当前进度并跳转下一页的回调。"""
    current_item = st.session_state.data[st.session_state.current_idx]
    current_id = current_item['id']
    
    # 1. 获取 Main Prompt
    main_prompt_val = st.session_state.get(f"main_prompt_{current_id}", current_item['original_prompt_text'])
    
    # 2. 获取 Remix Suggestions
    remix_suggestions = st.session_state.get(f"remix_{current_id}", [])
    # 如果用户在卡片上手动修改了，需要从 widget state 中获取最新值 (这里简化处理，假设直接用 stored state)
    
    # 3. 保存结果
    st.session_state.processed_results[current_id] = {
        "id": current_id,
        "prompt": main_prompt_val,
        "remixSuggestions": remix_suggestions
    }
    
    # 4. 跳转逻辑
    if st.session_state.current_idx < len(st.session_state.data) - 1:
        st.session_state.current_idx += 1
    else:
        st.balloons()
        st.success("🎉 All images processed! You can now download the dataset.")

# ================= 5. 主界面构建 =================
st.set_page_config(layout="wide", page_title="Image Dataset Maker", page_icon="🎨")
inject_custom_css()

# --- 阶段 1: 上传文件 ---
if not st.session_state.data:
    st.markdown("## 🎨 Create Your Image Dataset")
    with st.container(border=True):
        uploaded_file = st.file_uploader("Upload PPTX", type=["pptx"], help="Upload your presentation file.")
        start_id_input = st.number_input("Start ID", min_value=1, value=100, step=1)
        
        if uploaded_file is not None:
            if st.button("🚀 Load Slides & Begin", type="primary", use_container_width=True):
                with st.spinner("Processing PPTX and extracting images..."):
                    # 替换为你真实的函数: process_ppt_file(uploaded_file, start_id_input)
                    data, images = process_ppt_file_mock(uploaded_file, start_id_input) 
                    
                    if data:
                        st.session_state.data = data
                        st.session_state.images = images
                        st.session_state.current_idx = 0
                        st.rerun()
                    else:
                        st.error("No valid slides or images found in the PPTX.")

# --- 阶段 2: 主编辑界面 (完美的左右布局) ---
else:
    current_item = st.session_state.data[st.session_state.current_idx]
    current_id = current_item['id']
    img_filename = current_item['image_filename']
    
    # 计算进度
    total_count = len(st.session_state.data)
    processed_count = len(st.session_state.processed_results)
    progress_val = (processed_count / total_count) if total_count > 0 else 0

    # 使用 columns 创建左右布局，比例设为 [1, 1.2] 让右侧稍宽
    col_left, col_right = st.columns([1, 1.2], gap="medium")

    # ====== 左侧栏：固定图片展示 ======
    with col_left:
        st.subheader(f"ID {current_id}")
        img_data = st.session_state.images.get(img_filename)
        if img_data:
            # 使用自定义 CSS 类包裹图片
            st.markdown(
                f"""
                <div class="left-image-container">
                    <img src="data:image/png;base64,{base64.b64encode(img_data).decode()}" alt="Image {current_id}">
                </div>
                """,
                unsafe_allow_html=True
            )
        else:
            st.error(f"Image {img_filename} not found!")

    # ====== 右侧栏：可滚动编辑区 ======
    with col_right:
        # --- 顶部控制栏 (进度条 + 按钮) ---
        c_prog, c_btn = st.columns([3, 1])
        with c_prog:
            st.caption(f"Progress: {processed_count} / {total_count}")
            st.progress(progress_val)
        with c_btn:
            # 使用回调函数处理保存和跳转，避免直接修改 state 导致的错误
            st.button("💾 Save & Next", type="primary", use_container_width=True, on_click=on_save_and_next)

        # --- 开始滚动区域 ---
        st.markdown('<div class="right-scroll-container">', unsafe_allow_html=True)
        
        st.divider()

        # 1. Main Prompt 编辑
        st.subheader("📝 Main Prompt")
        st.text_area(
            "Edit the main description:",
            value=current_item['original_prompt_text'],
            height=150,
            key=f"main_prompt_{current_id}", # 使用唯一key绑定state
            label_visibility="collapsed"
        )

        st.divider()

        # 2. 批量粘贴功能 (修复 APIException 的核心)
        with st.expander("📋 Paste Remix Text (Replace Existing)"):
            # 注意：这里使用了一个固定的 key "batch_input_widget"
            st.text_area(
                "Paste generated options here (Format: 'Label: Prompt' per line):",
                height=120,
                key="batch_input_widget", 
                label_visibility="collapsed"
            )
            # 按钮绑定回调函数 on_batch_apply
            st.button("Apply Bulk Text", type="secondary", on_click=on_batch_apply)

        # 3. Remix Suggestions 展示
        st.subheader("🎨 Remix Suggestions")
        
        # 获取当前页面的建议列表，如果不存在则初始化为空
        if f"remix_{current_id}" not in st.session_state:
            st.session_state[f"remix_{current_id}"] = []
        
        current_suggestions = st.session_state[f"remix_{current_id}"]
        
        if not current_suggestions:
            st.info("No remix suggestions yet. Paste text above to add some.")
        else:
            # 遍历显示建议卡片
            for i, remix in enumerate(current_suggestions):
                with st.container(border=True):
                    # 使用列布局让标签和按钮在一行
                    c_label, c_btn = st.columns([4, 1])
                    with c_label:
                        # 简化的展示，实际应用中可以做成输入框供修改
                        st.text_input(f"Label {i+1}", value=remix['label'], key=f"lbl_{current_id}_{i}", disabled=True)
                    with c_btn:
                         st.button("✨ Verify", key=f"vfy_{current_id}_{i}", type="secondary", use_container_width=True, help="Click to verify this prompt (Mock Function)")
                    
                    st.text_area(f"Prompt {i+1}", value=remix['prompt'], height=80, key=f"prmt_{current_id}_{i}")

        st.divider()

        # 4. 下载区域 (在最后显示)
        if processed_count > 0:
            st.subheader("📦 Export Dataset")
            # 创建 ZIP 文件
            zip_data = create_final_zip(st.session_state.processed_results, st.session_state.images)
            st.download_button(
                label=f"⬇️ Download Dataset ({processed_count} items)",
                data=zip_data,
                file_name="image_dataset.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )

        # --- 结束滚动区域 ---
        st.markdown('</div>', unsafe_allow_html=True)
