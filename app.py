#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公文格式化工具 - Web 前端界面
基于 Streamlit 框架，调用 gongwen_formatter 模块生成符合 GB/T 9704-2012 规范的公文 Word 文档
"""

import sys
import os
import tempfile
from datetime import datetime

# 将当前脚本所在目录加入 sys.path，确保能导入同目录的 gongwen_formatter
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import streamlit as st
import gongwen_formatter


# ============================================================
# 页面配置
# ============================================================

st.set_page_config(
    page_title="公文格式化工具",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 自定义 CSS：政务风格主题（内联硬编码，不依赖 .streamlit/config.toml）
st.markdown("""
<style>
    /* 政务风格主题 */
    .stApp {
        --primary-color: #1a3a5c;
    }
    .block-title, h1, h2, h3 {
        color: #1a3a5c !important;
    }
    /* 按钮样式 */
    .stButton > button {
        background-color: #1a3a5c;
        color: white;
        border: none;
        border-radius: 4px;
        font-size: 16px;
        padding: 8px 24px;
    }
    .stButton > button:hover {
        background-color: #8b0000;
        color: white;
    }
    /* 侧边栏背景 */
    .stSidebar, [data-testid="stSidebar"] {
        background-color: #f5f5f5;
    }
    /* 下载按钮 */
    .stDownloadButton > button {
        background-color: #8b0000;
        color: white;
        border: none;
        border-radius: 4px;
    }
    .stDownloadButton > button:hover {
        background-color: #a00000;
        color: white;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# 示例公文文本
# ============================================================

EXAMPLE_TEXT = """关于加强信息化建设工作的通知

国发〔2024〕15号

各省、自治区、直辖市人民政府，国务院各部委、各直属机构：

一、总体要求
为深入贯彻落实党中央、国务院关于加快信息化建设的决策部署，推动数字政府建设取得新成效，现就加强信息化建设工作通知如下。

（一）指导思想
以习近平新时代中国特色社会主义思想为指导，全面贯彻党的二十大精神，坚持系统观念，强化顶层设计。

（二）基本原则
1. 统筹规划、协调发展。加强信息化建设的整体谋划和统筹协调。
2. 创新驱动、融合发展。以科技创新为动力，推动信息技术与经济社会发展深度融合。

二、重点任务
各地区、各部门要高度重视信息化建设工作，切实加强组织领导，确保各项任务落到实处。

附件：信息化建设重点工作清单
"""


# ============================================================
# 侧边栏：格式规范说明
# ============================================================

with st.sidebar:
    st.markdown("## 📖 格式规范说明")
    st.caption("依据 GB/T 9704-2012《党政机关公文格式》标准")

    with st.expander("📏 页面设置", expanded=False):
        st.markdown("""
- **纸张**：A4 (210mm × 297mm)，纵向
- **页边距**：上 37mm，下 35mm，左 28mm，右 26mm
- **页脚距边界**：25mm
- **页码**：底部居中，格式 "- 1 -"，宋体四号
""")

    with st.expander("📝 标题格式", expanded=False):
        st.markdown("""
**主标题**
- 字体：方正小标宋简体 / 黑体
- 字号：二号（22磅）
- 对齐：居中

**一级标题**（一、二、三、）
- 字体：黑体
- 字号：三号（16磅）
- 无首行缩进

**二级标题**（（一）（二））
- 字体：楷体_GB2312 / 楷体
- 字号：三号（16磅）
- 无首行缩进

**三级标题**（1. 2. 3.）
- 字体：仿宋_GB2312 加粗
- 字号：三号（16磅）
- 无首行缩进
""")

    with st.expander("📄 正文格式", expanded=False):
        st.markdown("""
- **字体**：仿宋_GB2312 / 仿宋
- **字号**：三号（16磅）
- **行距**：固定值 28 磅
- **首行缩进**：2 字符
- **对齐**：两端对齐

**发文字号**
- 字体：仿宋，三号
- 对齐：居中

**附件标识**
- 字体：仿宋，三号
- 对齐：左对齐
""")

    with st.expander("📊 表格格式", expanded=False):
        st.markdown("""
- **宽度**：与正文区域对齐
- **表头**：黑体，三号，居中
- **表内文字**：仿宋，三号，居中
- **边框**：全部实线，黑色，0.5磅
- **单元格对齐**：垂直居中，水平居中
""")

    with st.expander("🖼️ 图片格式", expanded=False):
        st.markdown("""
- **默认宽度**：页面有效宽度的80%（约12.5cm）
- **对齐**：居中
- **图题**：仿宋，小四号（12磅），居中
- **图题位置**：图片下方

**图片标记用法**：
- `[图片1]` — 基本插入
- `[图片1,10cm]` — 指定宽度10cm
- `[图片1,50%]` — 页面宽度的50%
- `[图片1:图题文字]` — 带图题
- `[图片1,8cm:图题文字]` — 指定宽度+图题
- 也支持 `[图1]`、`【图片1】` 等格式
""")

    st.markdown("---")

    # 公文类型选择
    doc_type = st.selectbox(
        "📋 公文类型",
        ["通知", "决定", "意见", "报告", "请示", "批复", "函", "讲话稿", "会议纪要"],
        index=0,
        help="不同类型公文的格式略有差异"
    )

    # 根据公文类型显示格式提示
    if doc_type in ["讲话稿", "会议纪要"]:
        st.caption('💡 讲话稿/会议纪要：问候语行（如"同志们："）不缩进')

    # 图片上传
    uploaded_images = st.file_uploader(
        "🖼️ 上传图片（可选）",
        type=["png", "jpg", "jpeg", "bmp", "gif"],
        accept_multiple_files=True,
        help="上传图片后，在文本中使用 [图片1]、[图片2] 标记图片位置"
    )

    # 显示上传图片的预览和标记名提示
    if uploaded_images:
        st.markdown("**已上传图片：**")
        for i, img in enumerate(uploaded_images, 1):
            col_img, col_info = st.columns([1, 2])
            with col_img:
                st.image(img, width=60)
            with col_info:
                st.caption(f"标记：`[图片{i}]`")

    st.markdown("---")
    st.markdown("""
<small>💡 **提示**：粘贴文本时，系统会自动识别主标题、发文字号、各级标题、正文、表格和图片标记等结构。</small>
""", unsafe_allow_html=True)


# ============================================================
# 主页面内容
# ============================================================

st.markdown("# 📄 公文格式化工具")
st.markdown("---")

# 使用两列布局
col_input, col_output = st.columns([3, 2], gap="large")

with col_input:
    st.markdown("### ✏️ 输入公文内容")

    # 文件上传区
    uploaded_file = st.file_uploader(
        "📂 上传文件（.txt 或 .docx）",
        type=["txt", "docx"],
        help="支持上传纯文本(.txt)或 Word(.docx)文件，内容将自动填入下方文本框",
    )

    # 处理上传文件，提取文本
    file_extracted_text = ""
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith(".txt"):
                file_extracted_text = uploaded_file.read().decode("utf-8")
            elif uploaded_file.name.endswith(".docx"):
                from docx import Document as DocxDocument
                # 保存到临时文件再读取
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name
                try:
                    doc_obj = DocxDocument(tmp_path)
                    paragraphs = [para.text for para in doc_obj.paragraphs]
                    file_extracted_text = "\n".join(paragraphs)
                finally:
                    os.unlink(tmp_path)
            st.toast(f"✅ 已读取文件：{uploaded_file.name}", icon="📄")
        except Exception as e:
            st.error(f"读取文件失败：{e}")

    # 文本输入区
    default_text = file_extracted_text if file_extracted_text else ""
    text_input = st.text_area(
        "公文文本",
        value=default_text,
        height=400,
        placeholder="请在此粘贴公文内容...\n\n示例格式：\n标题\n\n发文字号\n\n正文内容...\n\n一、一级标题\n（一）二级标题\n1. 三级标题\n\n[图片1] — 插入图片标记",
        label_visibility="collapsed",
    )

    # 如果刚上传了文件，更新文本框内容
    if file_extracted_text and text_input == "" and default_text != "":
        text_input = default_text

    # 示例按钮
    col_example, col_generate = st.columns([1, 2])
    with col_example:
        if st.button("📋 填入示例", use_container_width=True):
            st.rerun() if text_input == EXAMPLE_TEXT else None
            # 用 session_state 方式填入示例
            st.session_state["example_filled"] = True

    with col_generate:
        generate_clicked = st.button("🚀 一键生成公文", use_container_width=True, type="primary")

    # 处理示例填充
    if st.session_state.get("example_filled"):
        text_input = EXAMPLE_TEXT
        st.session_state["example_filled"] = False


with col_output:
    st.markdown("### 📥 生成结果")

    if generate_clicked:
        if not text_input or not text_input.strip():
            st.warning("⚠️ 请先输入公文内容！")
        else:
            with st.spinner("⏳ 正在生成公文文档，请稍候..."):
                try:
                    # 生成时间戳文件名
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename = f"公文_{timestamp}.docx"
                    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), output_filename)

                    # 构建图片字典
                    images = {}
                    if uploaded_images:
                        for i, img in enumerate(uploaded_images, 1):
                            images[f"图片{i}"] = img.read()

                    # 调用格式化模块
                    result_path = gongwen_formatter.main(
                        text_input.strip(),
                        output_path,
                        images=images if images else None,
                        doc_type=doc_type
                    )

                    st.success("✅ 公文文档生成成功！")

                    # 读取文件供下载
                    with open(result_path, "rb") as f:
                        file_bytes = f.read()

                    # 下载按钮
                    st.download_button(
                        label="⬇️ 下载公文文档",
                        data=file_bytes,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )

                    # 显示文件信息
                    file_size_kb = len(file_bytes) / 1024
                    st.info(f"📁 文件名：{output_filename}\n\n📊 文件大小：{file_size_kb:.1f} KB\n\n📋 公文类型：{doc_type}")

                    # 简要预览解析结果
                    with st.expander("🔍 解析结构预览", expanded=False):
                        elements = gongwen_formatter.parse_text(text_input.strip())
                        type_labels = {
                            "main_title": "📌 主标题",
                            "blank": "⬜ 空行",
                            "fawen_hao": "🔢 发文字号",
                            "body": "📝 正文",
                            "level1": "1️⃣ 一级标题",
                            "level2": "2️⃣ 二级标题",
                            "level3": "3️⃣ 三级标题",
                            "attachment_marker": "📎 附件标识",
                            "attachment_title": "📎 附件标题",
                            "table": "📊 表格",
                            "image": "🖼️ 图片",
                        }
                        for i, elem in enumerate(elements):
                            etype = elem["type"]
                            label = type_labels.get(etype, etype)
                            if etype == "table":
                                rows = elem.get("rows", [])
                                st.write(f"{i+1}. {label} — {len(rows)} 行 × {len(rows[0]) if rows else 0} 列")
                            elif etype == "image":
                                img_name = elem.get("name", "")
                                img_caption = elem.get("caption", "")
                                img_width = elem.get("width_spec", "")
                                detail = img_name
                                if img_width:
                                    detail += f"，宽度：{img_width}"
                                if img_caption:
                                    detail += f"，图题：{img_caption}"
                                st.write(f"{i+1}. {label} — {detail}")
                            else:
                                text_preview = elem.get("text", "")[:50]
                                st.write(f"{i+1}. {label} — {text_preview}")

                except Exception as e:
                    st.error(f"❌ 生成失败：{e}")
                    st.exception(e)
    else:
        # 未生成时的占位提示
        st.info("👆 请在左侧输入公文内容，点击「一键生成公文」按钮")


# ============================================================
# 页脚
# ============================================================

st.markdown("---")
st.markdown(
    '<p style="text-align: center; color: #888; font-size: 12px;">'
    "公文格式化工具 · 依据 GB/T 9704-2012 标准 · 仅供格式参考"
    "</p>",
    unsafe_allow_html=True,
)
