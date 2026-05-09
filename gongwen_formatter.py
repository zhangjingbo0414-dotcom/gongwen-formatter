#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中国党政机关公文格式自动化排版工具
依据 GB/T 9704-2012 标准
使用 python-docx 库生成规范格式的 Word 文档

模块结构：
  1. 页面设置
  2. 样式定义
  3. 文本解析
  4. 格式应用
  5. 文件输出
"""

import re
import os
from datetime import datetime

from docx import Document
from docx.shared import Pt, Mm, Cm, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml


# ============================================================
# 1. 页面设置模块
# ============================================================

def setup_page(doc):
    """
    设置文档页面参数（GB/T 9704-2012）
    - 页边距：上37mm，下35mm，左28mm，右26mm
    - 纸张方向：纵向
    - 纸张大小：A4 (210mm x 297mm)
    """
    section = doc.sections[0]

    # 纸张方向：纵向
    section.orientation = WD_ORIENT.PORTRAIT

    # 纸张大小：A4
    section.page_width = Mm(210)
    section.page_height = Mm(297)

    # 页边距
    section.top_margin = Mm(37)
    section.bottom_margin = Mm(35)
    section.left_margin = Mm(28)
    section.right_margin = Mm(26)

    # 页脚距离
    section.footer_distance = Cm(2.5)

    return section


# ============================================================
# 2. 样式定义模块
# ============================================================

# 字体名称常量
FONT_FANGSONG = '仿宋_GB2312'    # 正文
FONT_FANGSONG_ALT = '仿宋'       # 正文备选
FONT_HEITI = '黑体'              # 一级标题
FONT_KAITI = '楷体_GB2312'       # 二级标题
FONT_KAITI_ALT = '楷体'          # 二级标题备选
FONT_XIAOBIAOSONG = '方正小标宋简体'  # 主标题
FONT_SONGTI = '宋体'             # 页码

# 字号常量（磅值）
SIZE_ERHAO = Pt(22)      # 二号，用于主标题
SIZE_SANHAO = Pt(16)     # 三号，用于正文
SIZE_SIHAO = Pt(14)      # 四号，用于页码
SIZE_XIAOSI = Pt(12)     # 小四号，用于图标题

# 行距
LINE_SPACING = Pt(28)    # 固定值28磅


def set_run_font(run, font_name, font_size, bold=False, color=None):
    """
    设置 run 的字体属性
    - font_name: 中文字体名称
    - font_size: 字号（Pt值）
    - bold: 是否加粗
    - color: 字体颜色（RGBColor），默认黑色
    """
    run.font.size = font_size
    run.font.bold = bold
    if color is None:
        run.font.color.rgb = RGBColor(0, 0, 0)
    else:
        run.font.color.rgb = color

    # 设置中文字体
    run.font.name = font_name
    # 通过 rPr 设置东亚字体（确保中文字体生效）
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:eastAsia="{font_name}"/>')
        rPr.insert(0, rFonts)
    else:
        rFonts.set(qn('w:eastAsia'), font_name)


def set_paragraph_format(paragraph, alignment=None, first_line_indent=None,
                         line_spacing=LINE_SPACING, space_before=Pt(0),
                         space_after=Pt(0)):
    """
    设置段落格式
    - alignment: 对齐方式
    - first_line_indent: 首行缩进（如 2字符 用 Pt(32) 近似）
    - line_spacing: 行距
    - space_before: 段前距
    - space_after: 段后距
    """
    pf = paragraph.paragraph_format
    if alignment is not None:
        pf.alignment = alignment
    if first_line_indent is not None:
        pf.first_line_indent = first_line_indent
    pf.line_spacing = line_spacing
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.space_before = space_before
    pf.space_after = space_after


def add_page_number(doc):
    """
    添加页码：页面底部居中，格式为 "- 1 -"
    字体：宋体，四号
    """
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False

    # 清空已有页脚内容
    for p in footer.paragraphs:
        p._element.getparent().remove(p._element)

    # 新建页脚段落
    footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 添加页码字段 "- 页码 -"
    run1 = footer_para.add_run('- ')
    set_run_font(run1, FONT_SONGTI, SIZE_SIHAO)

    # 插入页码域代码
    fld_char_begin = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
    run_page = footer_para.add_run()
    run_page._element.append(fld_char_begin)
    set_run_font(run_page, FONT_SONGTI, SIZE_SIHAO)

    instr_text = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>')
    run_instr = footer_para.add_run()
    run_instr._element.append(instr_text)
    set_run_font(run_instr, FONT_SONGTI, SIZE_SIHAO)

    fld_char_end = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
    run_end = footer_para.add_run()
    run_end._element.append(fld_char_end)
    set_run_font(run_end, FONT_SONGTI, SIZE_SIHAO)

    run2 = footer_para.add_run(' -')
    set_run_font(run2, FONT_SONGTI, SIZE_SIHAO)


# ============================================================
# 3. 文本解析模块
# ============================================================

# 正则表达式：识别各层级标题
RE_LEVEL1 = re.compile(r'^[一二三四五六七八九十]+、')          # 一级标题：一、二、
RE_LEVEL2 = re.compile(r'^[（\(][一二三四五六七八九十]+[）\)]')  # 二级标题：（一）（二）
RE_LEVEL3 = re.compile(r'^\d+[\.．]')                          # 三级标题：1. 2.
RE_FAWEN_HAO = re.compile(r'^[^\s]*〔\d{4}〕\d+号')           # 发文字号
RE_TABLE_SEPARATOR = re.compile(r'^\|[-\s|]+\|$')              # 表格分隔行
RE_TABLE_ROW = re.compile(r'^\|.*\|$')                         # 表格数据行


def parse_text(text):
    """
    解析纯文本，识别各段落类型并结构化返回
    
    返回结构：
    [
        {'type': 'main_title', 'text': '...'},      # 公文主标题
        {'type': 'blank', 'text': ''},               # 空行（标题后空行）
        {'type': 'fawen_hao', 'text': '...'},        # 发文字号
        {'type': 'body', 'text': '...'},             # 正文
        {'type': 'level1', 'text': '...'},           # 一级标题
        {'type': 'level2', 'text': '...'},           # 二级标题
        {'type': 'level3', 'text': '...'},           # 三级标题
        {'type': 'attachment_marker', 'text': '...'}, # 附件标识
        {'type': 'attachment_title', 'text': '...'},  # 附件标题
        {'type': 'table', 'rows': [[...], ...]},     # 表格
        {'type': 'image', ...},                       # 图片（从docx输入时）
    ]
    """
    lines = text.split('\n')
    elements = []
    i = 0
    main_title_found = False

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # 跳过空行（但在特定位置保留）
        if not stripped:
            i += 1
            continue

        # --- 公文主标题：第一个非空行，且不超过20字 ---
        if not main_title_found:
            if len(stripped) <= 20:
                elements.append({'type': 'main_title', 'text': stripped})
                # 标题与正文之间空一行
                elements.append({'type': 'blank', 'text': ''})
                main_title_found = True
                i += 1
                continue
            else:
                # 超过20字，仍作为主标题处理（宽松处理）
                elements.append({'type': 'main_title', 'text': stripped})
                elements.append({'type': 'blank', 'text': ''})
                main_title_found = True
                i += 1
                continue

        # --- 表格识别 ---
        if RE_TABLE_ROW.match(stripped):
            table_rows = []
            while i < len(lines):
                row_line = lines[i].strip()
                if not RE_TABLE_ROW.match(row_line):
                    break
                # 跳过分隔行（|---|---|）
                if RE_TABLE_SEPARATOR.match(row_line):
                    i += 1
                    continue
                # 解析单元格
                cells = [c.strip() for c in row_line.split('|') if c.strip() != '']
                table_rows.append(cells)
                i += 1
            if table_rows:
                elements.append({'type': 'table', 'rows': table_rows})
            continue

        # --- 附件识别 ---
        if stripped.startswith('附件') or stripped.startswith('附件：'):
            elements.append({'type': 'blank', 'text': ''})  # 附件前空一行
            elements.append({'type': 'attachment_marker', 'text': stripped})
            i += 1
            continue

        # --- 发文字号 ---
        if RE_FAWEN_HAO.match(stripped):
            elements.append({'type': 'fawen_hao', 'text': stripped})
            i += 1
            continue

        # --- 一级标题 ---
        if RE_LEVEL1.match(stripped):
            elements.append({'type': 'level1', 'text': stripped})
            i += 1
            continue

        # --- 二级标题 ---
        if RE_LEVEL2.match(stripped):
            elements.append({'type': 'level2', 'text': stripped})
            i += 1
            continue

        # --- 三级标题 ---
        if RE_LEVEL3.match(stripped):
            elements.append({'type': 'level3', 'text': stripped})
            i += 1
            continue

        # --- 默认为正文 ---
        elements.append({'type': 'body', 'text': stripped})
        i += 1

    return elements


# ============================================================
# 4. 格式应用模块
# ============================================================

def apply_main_title(doc, text):
    """应用公文主标题格式：方正小标宋简体/黑体，二号，居中，行距28磅"""
    para = doc.add_paragraph()
    set_paragraph_format(
        para,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=None,
        line_spacing=LINE_SPACING,
        space_before=Pt(0),
        space_after=Pt(0)
    )
    run = para.add_run(text)
    # 优先使用方正小标宋简体，若不可用则用黑体替代
    set_run_font(run, FONT_XIAOBIAOSONG, SIZE_ERHAO, bold=False)
    apply_number_font(para)


def apply_blank_line(doc):
    """添加空行（标题与正文之间），字号三号，无缩进"""
    para = doc.add_paragraph()
    set_paragraph_format(
        para,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=None,
        line_spacing=LINE_SPACING,
        space_before=Pt(0),
        space_after=Pt(0)
    )
    run = para.add_run('')
    set_run_font(run, FONT_FANGSONG, SIZE_SANHAO)


def apply_fawen_hao(doc, text):
    """应用发文字号格式：仿宋，三号，居中，不加粗"""
    para = doc.add_paragraph()
    set_paragraph_format(
        para,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=None,
        line_spacing=LINE_SPACING,
        space_before=Pt(0),
        space_after=Pt(0)
    )
    run = para.add_run(text)
    set_run_font(run, FONT_FANGSONG, SIZE_SANHAO, bold=False)
    apply_number_font(para)


def apply_body(doc, text, first_indent=True):
    """
    应用正文格式：仿宋_GB2312，三号，首行缩进2字符，行距28磅
    - first_indent: 是否首行缩进，附件正文不缩进
    """
    para = doc.add_paragraph()
    indent = Pt(32) if first_indent else None  # 2字符约32磅（三号字2个字符宽度）
    set_paragraph_format(
        para,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_indent=indent,
        line_spacing=LINE_SPACING,
        space_before=Pt(0),
        space_after=Pt(0)
    )
    run = para.add_run(text)
    set_run_font(run, FONT_FANGSONG, SIZE_SANHAO, bold=False)
    apply_number_font(para)


def apply_number_font(paragraph):
    """将段落中的数字和英文字母设置为 Times New Roman 字体，其余保持原字体"""
    runs_to_remove = []
    for run in paragraph.runs:
        text = run.text
        if not text:
            continue
        # 用正则分割：数字+英文部分 和 非数字英文部分
        parts = re.split(r'([0-9a-zA-Z]+)', text)
        # 检查是否需要拆分（即是否有混合内容）
        non_empty_parts = [p for p in parts if p]
        has_alnum = any(re.match(r'^[0-9a-zA-Z]+$', p) for p in non_empty_parts)
        has_non_alnum = any(not re.match(r'^[0-9a-zA-Z]+$', p) for p in non_empty_parts)
        if not (has_alnum and has_non_alnum):
            # 无需拆分，但如果全是数字/英文则改字体
            if has_alnum and not has_non_alnum:
                run.font.name = 'Times New Roman'
                # 更新 rFonts 中的 ascii 和 hAnsi 属性
                r = run._element
                rPr = r.get_or_add_rPr()
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is not None:
                    rFonts.set(qn('w:ascii'), 'Times New Roman')
                    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
            continue
        # 需要拆分：为每个 part 创建新 run
        for part in non_empty_parts:
            new_run = paragraph.add_run(part)
            # 复制原 run 的格式
            new_run.font.size = run.font.size
            new_run.font.bold = run.font.bold
            new_run.font.italic = run.font.italic
            if run.font.color.rgb is not None:
                new_run.font.color.rgb = run.font.color.rgb
            # 设置字体
            if re.match(r'^[0-9a-zA-Z]+$', part):
                new_run.font.name = 'Times New Roman'
                # 复制 eastAsia 字体保持原样
                r = run._element
                rPr = r.get_or_add_rPr()
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is not None:
                    ea_font = rFonts.get(qn('w:eastAsia'))
                else:
                    ea_font = None
                nr = new_run._element
                nrPr = nr.get_or_add_rPr()
                nrFonts = nrPr.find(qn('w:rFonts'))
                if nrFonts is None:
                    nrFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
                    nrPr.insert(0, nrFonts)
                nrFonts.set(qn('w:ascii'), 'Times New Roman')
                nrFonts.set(qn('w:hAnsi'), 'Times New Roman')
                if ea_font:
                    nrFonts.set(qn('w:eastAsia'), ea_font)
            else:
                # 中文部分保持原字体
                new_run.font.name = run.font.name
                # 复制 rFonts 属性
                r = run._element
                rPr = r.get_or_add_rPr()
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is not None:
                    nr = new_run._element
                    nrPr = nr.get_or_add_rPr()
                    nrFonts = nrPr.find(qn('w:rFonts'))
                    if nrFonts is None:
                        nrFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
                        nrPr.insert(0, nrFonts)
                    for attr_key, attr_val in rFonts.attrib.items():
                        nrFonts.set(attr_key, attr_val)
        # 记录需要删除的原 run
        runs_to_remove.append(run)
    # 删除原来的 run
    for run in runs_to_remove:
        run._element.getparent().remove(run._element)


def apply_level1_title(doc, text):
    """一级标题格式：黑体，三号，首行缩进2字符，行距28磅"""
    para = doc.add_paragraph()
    set_paragraph_format(
        para,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_indent=Pt(32),
        line_spacing=LINE_SPACING,
        space_before=Pt(0),
        space_after=Pt(0)
    )
    run = para.add_run(text)
    set_run_font(run, FONT_HEITI, SIZE_SANHAO, bold=False)
    apply_number_font(para)


def apply_level2_title(doc, text):
    """二级标题格式：楷体_GB2312/楷体，三号，首行缩进2字符，行距28磅"""
    para = doc.add_paragraph()
    set_paragraph_format(
        para,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_indent=Pt(32),
        line_spacing=LINE_SPACING,
        space_before=Pt(0),
        space_after=Pt(0)
    )
    run = para.add_run(text)
    set_run_font(run, FONT_KAITI, SIZE_SANHAO, bold=False)
    apply_number_font(para)


def apply_level3_title(doc, text):
    """三级标题格式：仿宋_GB2312加粗，三号，首行缩进2字符，行距28磅"""
    para = doc.add_paragraph()
    set_paragraph_format(
        para,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_indent=Pt(32),
        line_spacing=LINE_SPACING,
        space_before=Pt(0),
        space_after=Pt(0)
    )
    run = para.add_run(text)
    set_run_font(run, FONT_FANGSONG, SIZE_SANHAO, bold=True)
    apply_number_font(para)


def apply_attachment_marker(doc, text):
    """附件标识格式：仿宋，三号，无首行缩进"""
    para = doc.add_paragraph()
    set_paragraph_format(
        para,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=None,
        line_spacing=LINE_SPACING,
        space_before=Pt(0),
        space_after=Pt(0)
    )
    run = para.add_run(text)
    set_run_font(run, FONT_FANGSONG, SIZE_SANHAO, bold=False)
    apply_number_font(para)


def apply_table(doc, table_data):
    """
    应用表格格式
    - 表格宽度与正文对齐
    - 表头：黑体，三号，居中
    - 表内文字：仿宋，三号，居中
    - 边框：全部实线，黑色，0.5磅
    - 单元格对齐：垂直居中，水平居中
    """
    if not table_data or not table_data[0]:
        return

    num_cols = len(table_data[0])
    num_rows = len(table_data)

    # 计算正文区域宽度（A4宽 - 左右边距）
    section = doc.sections[0]
    content_width = section.page_width - section.left_margin - section.right_margin

    # 添加表格前空行
    para_before = doc.add_paragraph()
    set_paragraph_format(
        para_before,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=None,
        line_spacing=LINE_SPACING,
        space_before=Pt(0),
        space_after=Pt(0)
    )

    # 创建表格
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 设置表格总宽度
    table_element = table._tbl
    tblPr = table_element.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>')
        table_element.insert(0, tblPr)

    # 设置表格宽度为正文宽度
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="{content_width}" w:type="dxa"/>')
        tblPr.append(tblW)
    else:
        tblW.set(qn('w:w'), str(content_width))
        tblW.set(qn('w:type'), 'dxa')

    # 设置每列等宽
    col_width = int(content_width / num_cols)
    tblGrid = table_element.find(qn('w:tblGrid'))
    if tblGrid is None:
        tblGrid = parse_xml(f'<w:tblGrid {nsdecls("w")}/>')
        table_element.insert(1, tblGrid)
    else:
        # 清空已有列定义
        for gc in tblGrid.findall(qn('w:gridCol')):
            tblGrid.remove(gc)
    for _ in range(num_cols):
        gridCol = parse_xml(f'<w:gridCol {nsdecls("w")} w:w="{col_width}"/>')
        tblGrid.append(gridCol)

    # 设置表格边框：全部实线，黑色，0.5磅
    set_table_borders(table)

    # 填充表格内容
    for row_idx, row_data in enumerate(table_data):
        row = table.rows[row_idx]
        for col_idx, cell_text in enumerate(row_data):
            if col_idx >= num_cols:
                break
            cell = row.cells[col_idx]
            # 清空默认段落
            cell_para = cell.paragraphs[0]
            cell_para.clear()

            # 设置单元格对齐：垂直居中
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            # 水平居中
            cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 设置单元格段落格式
            pf = cell_para.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = Pt(28)
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY

            run = cell_para.add_run(cell_text.strip())

            # 表头（第一行）用黑体，其余用仿宋
            if row_idx == 0:
                set_run_font(run, FONT_HEITI, SIZE_SANHAO, bold=False)
            else:
                set_run_font(run, FONT_FANGSONG, SIZE_SANHAO, bold=False)

    # 表格后空行
    para_after = doc.add_paragraph()
    set_paragraph_format(
        para_after,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=None,
        line_spacing=LINE_SPACING,
        space_before=Pt(0),
        space_after=Pt(0)
    )


def set_table_borders(table):
    """
    设置表格边框：全部实线，黑色，0.5磅
    """
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>')
        tbl.insert(0, tblPr)

    # 移除已有边框设置
    existing_borders = tblPr.find(qn('w:tblBorders'))
    if existing_borders is not None:
        tblPr.remove(existing_borders)

    # 添加边框
    borders_xml = f'''<w:tblBorders {nsdecls("w")}>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tblBorders>'''
    # w:sz="4" = 0.5磅 (1/8磅为单位, 4*1/8=0.5磅)
    borders_element = parse_xml(borders_xml)
    tblPr.append(borders_element)


# ============================================================
# 5. 文件输出模块
# ============================================================

def format_gongwen(text, output_path=None):
    """
    主格式化函数：接收纯文本，输出规范格式的 .docx 文件
    
    参数：
      text: 纯文本字符串
      output_path: 输出文件路径，默认为 "公文_时间戳.docx"
    
    返回：
      生成的文件路径
    """
    # 创建文档
    doc = Document()

    # 1. 页面设置
    setup_page(doc)

    # 2. 解析文本
    elements = parse_text(text)

    # 3. 应用格式
    for elem in elements:
        etype = elem['type']

        if etype == 'main_title':
            apply_main_title(doc, elem['text'])

        elif etype == 'blank':
            apply_blank_line(doc)

        elif etype == 'fawen_hao':
            apply_fawen_hao(doc, elem['text'])

        elif etype == 'body':
            apply_body(doc, elem['text'], first_indent=True)

        elif etype == 'level1':
            apply_level1_title(doc, elem['text'])

        elif etype == 'level2':
            apply_level2_title(doc, elem['text'])

        elif etype == 'level3':
            apply_level3_title(doc, elem['text'])

        elif etype == 'attachment_marker':
            apply_attachment_marker(doc, elem['text'])

        elif etype == 'table':
            apply_table(doc, elem['rows'])

    # 4. 添加页码
    add_page_number(doc)

    # 5. 输出文件
    if output_path is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = f'公文_{timestamp}.docx'

    # 确保目录存在
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    doc.save(output_path)
    return output_path


# ============================================================
# 入口函数
# ============================================================

def main(text, output_path=None):
    """
    入口函数
    
    参数：
      text: 纯文本字符串（公文内容）
      output_path: 输出文件路径（可选）
    
    返回：
      生成的文件路径
    """
    return format_gongwen(text, output_path)


if __name__ == '__main__':
    import sys

    if len(sys.argv) >= 2:
        # 从命令行参数获取文本（或文件路径）
        arg = sys.argv[1]
        if os.path.isfile(arg):
            with open(arg, 'r', encoding='utf-8') as f:
                input_text = f.read()
        else:
            input_text = arg

        out = sys.argv[2] if len(sys.argv) >= 3 else None
        result = main(input_text, out)
        print(f'公文已生成：{result}')
    else:
        print('用法：python gongwen_formatter.py <文本或文件路径> [输出路径]')
        print('示例：python gongwen_formatter.py "公文内容.txt" "输出.docx"')
