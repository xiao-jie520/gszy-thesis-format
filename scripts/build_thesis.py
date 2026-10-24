# -*- coding: utf-8 -*-
"""
build_thesis.py — 依据《甘肃中医药大学 医学信息工程学院本科毕业论文（设计）撰写规范》
把一份 Markdown 稿件排版为符合规范的 .docx。

用法:
    python build_thesis.py manuscript.md -o 论文.docx
    python build_thesis.py manuscript.md -o 论文.docx --max-figure-cm 11
    python build_thesis.py --skeleton 论文.docx        # 只看规范骨架长什么样

稿件语法见 references/manuscript-syntax.md。
依赖: python-docx
"""
import argparse
import os
import re
import sys

try:
    from docx import Document
    from docx.enum.section import WD_SECTION
    from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
    from docx.enum.text import (WD_ALIGN_PARAGRAPH, WD_LINE_SPACING,
                                WD_TAB_ALIGNMENT, WD_TAB_LEADER)
    from docx.enum.style import WD_STYLE_TYPE
    from docx.image.image import Image as DocxImage
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:  # pragma: no cover
    sys.exit("缺少依赖 python-docx，请先执行: pip install python-docx")


# ---------------------------------------------------------------- 规范常量
PAGE_W, PAGE_H = 21.0, 29.7            # A4
MARGIN_TOP, MARGIN_BOTTOM = 3.5, 2.5   # cm
MARGIN_LR = 2.5
HEADER_DIST, FOOTER_DIST = 2.5, 2.0
HEADER_TEXT = "甘肃中医药大学本科毕业论文（设计）"

SZ = {
    "cover_school": 18.0,   # 小二 楷体加粗
    "cover_title": 22.0,    # 二号 黑体加粗
    "h1": 15.0,             # 小三 黑体
    "h2": 12.0,             # 小四 黑体
    "body": 12.0,           # 小四 宋体
    "small": 10.5,          # 五号
    "page": 9.0,            # 小五
    "cover_label": 16.0,    # 三号
}

SONG, HEI, KAI, TNR = "宋体", "黑体", "楷体", "Times New Roman"

CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"


# ---------------------------------------------------------------- 底层工具
def style_run(run, size, ea=SONG, latin=TNR, bold=False, italic=False, underline=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    run.font.color.rgb = RGBColor(0, 0, 0)
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    rf.set(qn("w:ascii"), latin)
    rf.set(qn("w:hAnsi"), latin)
    rf.set(qn("w:eastAsia"), ea)
    rf.set(qn("w:cs"), latin)
    rf.set(qn("w:hint"), "eastAsia")


def para_spacing(p, line=1.5, before_pt=None, after_pt=None,
                 before_lines=None, after_lines=None):
    if line is not None:
        p.paragraph_format.line_spacing = line
    ppr = p._p.get_or_add_pPr()
    sp = ppr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        ppr.append(sp)
    if before_pt is not None:
        sp.set(qn("w:before"), str(int(round(before_pt * 20))))
    elif before_lines is not None:
        sp.set(qn("w:beforeLines"), str(int(before_lines * 100)))
        sp.set(qn("w:before"), str(int(round(before_lines * 312))))
    if after_pt is not None:
        sp.set(qn("w:after"), str(int(round(after_pt * 20))))
    elif after_lines is not None:
        sp.set(qn("w:afterLines"), str(int(after_lines * 100)))
        sp.set(qn("w:after"), str(int(round(after_lines * 312))))
    if before_pt is None and after_pt is None and before_lines is None and after_lines is None:
        sp.set(qn("w:before"), "0")
        sp.set(qn("w:after"), "0")


def para_indent(p, first_line_chars=0.0, left_cm=None, hanging_chars=0.0):
    if left_cm is not None:
        p.paragraph_format.left_indent = Cm(left_cm)
    ppr = p._p.get_or_add_pPr()
    if first_line_chars or hanging_chars:
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            ppr.append(ind)
        if hanging_chars:
            ind.set(qn("w:hangingChars"), str(int(hanging_chars * 100)))
            ind.set(qn("w:hanging"), str(int(hanging_chars * SZ["body"] * 20)))
        else:
            ind.set(qn("w:firstLineChars"), str(int(first_line_chars * 100)))
            ind.set(qn("w:firstLine"), str(int(first_line_chars * SZ["body"] * 20)))
    # 取消“对齐到网格”，避免行距被文档网格改写
    snap = ppr.find(qn("w:snapToGrid"))
    if snap is None:
        snap = OxmlElement("w:snapToGrid")
        ppr.append(snap)
    snap.set(qn("w:val"), "0")


def para_align(p, align):
    p.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }[align]


def _add_math(p, s, size, ea=SONG, latin=TNR):
    """行内公式：变量斜体，支持 x_1 / x^{2} 上下标。"""
    buf = []

    def flush():
        if buf:
            r = p.add_run("".join(buf))
            style_run(r, size, ea=ea, latin=latin, italic=True)
            buf.clear()

    i = 0
    while i < len(s):
        c = s[i]
        if c in "_^" and i + 1 < len(s):
            flush()
            va = "sub" if c == "_" else "sup"
            if s[i + 1] == "{":
                j = s.find("}", i + 2)
                if j == -1:
                    buf.append(c)
                    i += 1
                    continue
                txt, i = s[i + 2:j], j + 1
            else:
                txt, i = s[i + 1], i + 2
            r = p.add_run(txt)
            style_run(r, size, ea=ea, latin=latin, italic=True)
            r.font.subscript = va == "sub"
            r.font.superscript = va == "sup"
        else:
            buf.append(c)
            i += 1
    flush()


def add_runs(p, text, size, ea=SONG, latin=TNR, bold=False, base_bold=None):
    """支持 **加粗** 与 $行内公式$（变量斜体、支持上下标）。"""
    text = re.sub(r"\$\$(.+?)\$\$", r"$\1$", text)
    for i, seg in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if not seg:
            continue
        b = bold if base_bold is None else base_bold
        if i % 2 == 1:
            b = True
        for j, piece in enumerate(re.split(r"\$(.+?)\$", seg)):
            if not piece:
                continue
            if j % 2 == 1:
                _add_math(p, piece, size, ea, latin)
            else:
                r = p.add_run(piece)
                style_run(r, size, ea=ea, latin=latin, bold=b)


def kill_section_break_paragraphs(doc):
    """把承载分节符的空段落压缩到 1pt，避免多出空白行。"""
    for p in doc.paragraphs:
        ppr = p._p.pPr
        if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
            para_spacing(p, line=None, before_pt=0, after_pt=0)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            p.paragraph_format.line_spacing = Pt(1)
            if not p.runs:
                p.add_run("")
            for r in p.runs:
                r.font.size = Pt(1)


# ---------------------------------------------------------------- 节属性
SECTPR_ORDER = [
    "w:footnotePr", "w:endnotePr", "w:type", "w:pgSz", "w:pgMar", "w:paperSrc",
    "w:pgBorders", "w:lnNumType", "w:pgNumType", "w:cols", "w:formProt",
    "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection", "w:bidi",
    "w:rtlGutter", "w:docGrid", "w:printerSettings", "w:sectPrChange",
]


def _insert_ordered(parent, el):
    idx = next(i for i, t in enumerate(SECTPR_ORDER) if qn(t) == el.tag)
    for child in list(parent):
        for i, t in enumerate(SECTPR_ORDER):
            if qn(t) == child.tag and i > idx:
                child.addprevious(el)
                return
    parent.append(el)


def set_page_numbering(section, fmt="decimal", start=1):
    sectpr = section._sectPr
    old = sectpr.find(qn("w:pgNumType"))
    if old is not None:
        sectpr.remove(old)
    el = OxmlElement("w:pgNumType")
    el.set(qn("w:fmt"), fmt)
    el.set(qn("w:start"), str(start))
    _insert_ordered(sectpr, el)


def setup_page(section):
    section.page_width = Cm(PAGE_W)
    section.page_height = Cm(PAGE_H)
    section.top_margin = Cm(MARGIN_TOP)
    section.bottom_margin = Cm(MARGIN_BOTTOM)
    section.left_margin = Cm(MARGIN_LR)
    section.right_margin = Cm(MARGIN_LR)
    section.header_distance = Cm(HEADER_DIST)
    section.footer_distance = Cm(FOOTER_DIST)
    sectpr = section._sectPr
    if sectpr.find(qn("w:docGrid")) is None:
        grid = OxmlElement("w:docGrid")
        grid.set(qn("w:type"), "none")
        grid.set(qn("w:linePitch"), "312")
        _insert_ordered(sectpr, grid)


def add_page_field(p, size=SZ["page"]):
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), " PAGE ")
    inner = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rf = OxmlElement("w:rFonts")
    for k in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(qn(k), TNR)
    rf.set(qn("w:eastAsia"), SONG)
    rpr.append(rf)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rpr.append(sz)
    inner.append(rpr)
    t = OxmlElement("w:t")
    t.text = "1"
    inner.append(t)
    fld.append(inner)
    p._p.append(fld)


def config_section(section, fmt, start=1, header_text=None, page_number=True):
    setup_page(section)
    set_page_numbering(section, fmt, start)

    section.header.is_linked_to_previous = False
    hp = section.header.paragraphs[0]
    for r in list(hp.runs):
        r._element.getparent().remove(r._element)
    para_align(hp, "center")
    para_spacing(hp, line=1.0, before_pt=0, after_pt=0)
    if header_text:
        r = hp.add_run(header_text)
        style_run(r, SZ["small"], ea=SONG, latin=TNR)
        ppr = hp._p.get_or_add_pPr()
        bdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "thinThickSmallGap")
        bottom.set(qn("w:sz"), "12")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "auto")
        bdr.append(bottom)
        ppr.append(bdr)

    section.footer.is_linked_to_previous = False
    fp = section.footer.paragraphs[0]
    for r in list(fp.runs):
        r._element.getparent().remove(r._element)
    para_align(fp, "center")
    para_spacing(fp, line=1.0, before_pt=0, after_pt=0)
    if page_number:
        add_page_field(fp)


# ---------------------------------------------------------------- 样式
def setup_styles(doc):
    styles = doc.styles

    # 文档默认字体：去掉主题字体，全部锁定为 宋体 / Times New Roman
    dd = styles.element.find(qn("w:docDefaults"))
    if dd is not None:
        rpr_def = dd.find(qn("w:rPrDefault"))
        if rpr_def is not None:
            rpr = rpr_def.find(qn("w:rPr"))
            if rpr is not None:
                rf = rpr.find(qn("w:rFonts"))
                if rf is None:
                    rf = OxmlElement("w:rFonts")
                    rpr.insert(0, rf)
                for attr in ("w:asciiTheme", "w:eastAsiaTheme", "w:hAnsiTheme", "w:cstheme"):
                    if rf.get(qn(attr)) is not None:
                        del rf.attrib[qn(attr)]
                rf.set(qn("w:ascii"), TNR)
                rf.set(qn("w:hAnsi"), TNR)
                rf.set(qn("w:eastAsia"), SONG)
                rf.set(qn("w:cs"), TNR)
    normal = styles["Normal"]
    normal.font.size = Pt(SZ["body"])
    normal.font.color.rgb = RGBColor(0, 0, 0)
    rf = normal.element.get_or_add_rPr().get_or_add_rFonts()
    rf.set(qn("w:ascii"), TNR)
    rf.set(qn("w:hAnsi"), TNR)
    rf.set(qn("w:eastAsia"), SONG)
    rf.set(qn("w:cs"), TNR)
    pf = normal.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    specs = [
        ("Heading 1", SZ["h1"], "center", True, True),
        ("Heading 2", SZ["h2"], "left", False, False),
        ("Heading 3", SZ["h2"], "left", False, False),
    ]
    for name, size, align, one_line, pbb in specs:
        st = styles[name]
        st.font.size = Pt(size)
        st.font.bold = False
        st.font.color.rgb = RGBColor(0, 0, 0)
        rf = st.element.get_or_add_rPr().get_or_add_rFonts()
        for attr in ("w:asciiTheme", "w:eastAsiaTheme", "w:hAnsiTheme", "w:cstheme"):
            if rf.get(qn(attr)) is not None:
                del rf.attrib[qn(attr)]
        rf.set(qn("w:ascii"), TNR)
        rf.set(qn("w:hAnsi"), TNR)
        rf.set(qn("w:eastAsia"), HEI)
        rf.set(qn("w:cs"), TNR)
        ppr = st.element.get_or_add_pPr()

        for tag in ("w:spacing", "w:jc", "w:pageBreakBefore", "w:keepNext"):
            el = ppr.find(qn(tag))
            if el is not None:
                ppr.remove(el)

        sp = OxmlElement("w:spacing")
        if one_line:
            sp.set(qn("w:beforeLines"), "100")
            sp.set(qn("w:afterLines"), "100")
            sp.set(qn("w:before"), "312")
            sp.set(qn("w:after"), "312")
        else:
            sp.set(qn("w:before"), "120")
            sp.set(qn("w:after"), "120")
        sp.set(qn("w:line"), "360")
        sp.set(qn("w:lineRule"), "auto")
        ppr.append(sp)

        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), align)
        ppr.append(jc)
        if pbb:
            ppr.append(OxmlElement("w:pageBreakBefore"))
        ppr.append(OxmlElement("w:keepNext"))

        sm = ppr.find(qn("w:snapToGrid"))
        if sm is None:
            sm = OxmlElement("w:snapToGrid")
            ppr.append(sm)
        sm.set(qn("w:val"), "0")

    setup_toc_styles(doc)


def setup_toc_styles(doc):
    """预置 toc 1/2/3：宋体小四、1.5 倍行距、右对齐点状前导符（否则 Word 会照抄标题字号）。"""
    for lvl, indent in ((1, 0.0), (2, 0.74), (3, 1.48)):
        name = f"toc {lvl}"
        try:
            st = doc.styles[name]
        except KeyError:
            st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = doc.styles["Normal"]
        st.font.size = Pt(SZ["body"])
        st.font.bold = False
        st.font.color.rgb = RGBColor(0, 0, 0)
        rf = st.element.get_or_add_rPr().get_or_add_rFonts()
        rf.set(qn("w:ascii"), TNR)
        rf.set(qn("w:hAnsi"), TNR)
        rf.set(qn("w:eastAsia"), SONG)
        pf = st.paragraph_format
        pf.line_spacing = 1.5
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.left_indent = Cm(indent)
        pf.tab_stops.add_tab_stop(Cm(PAGE_W - 2 * MARGIN_LR), WD_TAB_ALIGNMENT.RIGHT,
                                  WD_TAB_LEADER.DOTS)


def _append_ppr(p, el):
    """按 schema 顺序把元素插到 pPr 中（rPr 之前）。"""
    ppr = p._p.get_or_add_pPr()
    rpr = ppr.find(qn("w:rPr"))
    if rpr is not None:
        rpr.addprevious(el)
    else:
        ppr.append(el)


def hide_from_toc(p):
    """把段落大纲级别设为正文，使其不进入目录（备用：目录标题现由 TOC Heading 样式处理）。"""
    el = OxmlElement("w:outlineLvl")
    el.set(qn("w:val"), "9")
    _append_ppr(p, el)


# ---------------------------------------------------------------- 表格 / 图 / 公式
TBLPR_ORDER = [
    "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual", "w:tblStyleRowBandSize",
    "w:tblStyleColBandSize", "w:tblW", "w:jc", "w:tblCellSpacing", "w:tblInd",
    "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
    "w:tblCaption", "w:tblDescription",
]


def _reorder_tblpr(tblpr):
    idx = {qn(t): i for i, t in enumerate(TBLPR_ORDER)}
    for child in sorted(list(tblpr), key=lambda c: idx.get(c.tag, 99)):
        tblpr.append(child)


def _set_tbl_borders(table, sz_outer=12):
    tblpr = table._tbl.tblPr
    old = tblpr.find(qn("w:tblBorders"))
    if old is not None:
        tblpr.remove(old)
    b = OxmlElement("w:tblBorders")
    for edge, val, sz in (
        ("top", "single", sz_outer),
        ("left", "none", 0),
        ("bottom", "single", sz_outer),
        ("right", "none", 0),
        ("insideH", "none", 0),
        ("insideV", "none", 0),
    ):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        b.append(el)
    tblpr.append(b)


def _cell_bottom_border(cell, sz=4):
    tcpr = cell._tc.get_or_add_tcPr()
    old = tcpr.find(qn("w:tcBorders"))
    if old is not None:
        tcpr.remove(old)
    b = OxmlElement("w:tcBorders")
    el = OxmlElement("w:bottom")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), "000000")
    b.append(el)
    tcpr.append(b)


def add_caption(text, kind="figure"):
    p = doc_current.add_paragraph()
    para_align(p, "center")
    para_spacing(p, line=1.5, before_pt=0, after_pt=0)
    para_indent(p, 0)
    if kind == "table":
        p.paragraph_format.keep_with_next = True
    add_runs(p, text, SZ["small"], ea=SONG, latin=TNR)
    return p


def _disp_width(s):
    return sum(2 if ord(c) > 127 else 1 for c in s)


def _apply_col_widths(table, rows):
    """按各列内容的显示宽度分配列宽，避免“分类维/护”这类难看的断行。"""
    n_cols = len(rows[0]) if rows else 0
    if not n_cols:
        return
    total_cm = PAGE_W - 2 * MARGIN_LR
    weights = []
    for ci in range(n_cols):
        w = max((_disp_width(r[ci]) for r in rows if ci < len(r)), default=4)
        weights.append(max(w, 4))
    # 极短列压低权重、极长列设上限，避免一列吃掉整表
    weights = [min(w, 26) for w in weights]
    s = sum(weights)
    widths = [max(2.0, total_cm * w / s) for w in weights]
    scale = total_cm / sum(widths)
    widths = [w * scale for w in widths]

    for ci, w in enumerate(widths):
        for row in table.rows:
            row.cells[ci].width = Cm(w)
    # 同步 tblGrid，并把表宽锁定为固定布局，Word 才会按此渲染
    grid = table._tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for ci, gc in enumerate(grid.findall(qn("w:gridCol"))):
            if ci < len(widths):
                gc.set(qn("w:w"), str(int(widths[ci] * 567)))
    tblpr = table._tbl.tblPr
    for tag in ("w:tblW", "w:tblLayout"):
        old = tblpr.find(qn(tag))
        if old is not None:
            tblpr.remove(old)
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:type"), "dxa")
    tw.set(qn("w:w"), str(int(sum(widths) * 567)))
    tblpr.append(tw)
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblpr.append(lay)
    mar = OxmlElement("w:tblCellMar")
    for edge, w in (("top", 60), ("left", 108), ("bottom", 60), ("right", 108)):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:w"), str(w))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblpr.append(mar)
    _reorder_tblpr(tblpr)


def add_table(rows, caption=None, col_widths=None):
    if caption:
        add_caption(caption, "table")
    n_cols = max(len(r) for r in rows)
    table = doc_current.add_table(rows=len(rows), cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    tblpr = table._tbl.tblPr
    old_w = tblpr.find(qn("w:tblW"))
    if old_w is not None:
        tblpr.remove(old_w)
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:type"), "pct")
    tw.set(qn("w:w"), "5000")
    tblpr.append(tw)
    _set_tbl_borders(table)
    _reorder_tblpr(tblpr)
    for ri, row in enumerate(rows):
        for ci in range(n_cols):
            cell = table.cell(ri, ci)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            para_align(p, "center")
            para_spacing(p, line=1.0, before_pt=0, after_pt=0)
            para_indent(p, 0)
            txt = row[ci] if ci < len(row) else ""
            add_runs(p, txt, SZ["small"], ea=SONG, latin=TNR)
            if ri == 0:
                _cell_bottom_border(cell, 4)
    if col_widths:
        for ci, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[ci].width = Cm(w)
    else:
        _apply_col_widths(table, rows)
    # 表格与下文留一行空白
    spacer = doc_current.add_paragraph()
    para_spacing(spacer, line=1.0, before_pt=0, after_pt=0)
    para_indent(spacer, 0)
    spacer.add_run("").font.size = Pt(SZ["small"])
    return table


def add_figure(path, caption=None, max_cm=11.0, max_h_cm=17.0):
    p = doc_current.add_paragraph()
    para_align(p, "center")
    para_spacing(p, line=1.0, before_pt=6, after_pt=0)
    para_indent(p, 0)
    p.paragraph_format.keep_with_next = True
    width = max_cm
    try:
        img = DocxImage.from_file(path)
        w_cm, h_cm = img.width.cm, img.height.cm
        scale = min(1.0, max_cm / w_cm if w_cm else 1.0, max_h_cm / h_cm if h_cm else 1.0)
        width = max(2.0, w_cm * scale)
    except Exception:
        pass
    p.add_run().add_picture(path, width=Cm(width))
    if caption:
        add_caption(caption, "figure")
    return p


def add_formula(expr, number):
    table = doc_current.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        b.append(el)
    table._tbl.tblPr.append(b)
    for ci, w in enumerate((Cm(13.4), Cm(2.6))):
        table.columns[ci].width = w
        table.cell(0, ci).width = w
    p1 = table.cell(0, 0).paragraphs[0]
    para_align(p1, "center")
    para_spacing(p1, line=1.5, before_pt=6, after_pt=6)
    para_indent(p1, 0)
    add_runs(p1, expr, SZ["body"], ea=SONG, latin=TNR)
    p2 = table.cell(0, 1).paragraphs[0]
    para_align(p2, "right")
    para_spacing(p2, line=1.5, before_pt=6, after_pt=6)
    para_indent(p2, 0)
    add_runs(p2, f"({number})", SZ["body"], ea=SONG, latin=TNR)
    return table


# ---------------------------------------------------------------- 文档骨架
doc_current = None


def blank(n=1, size=SZ["body"]):
    for _ in range(n):
        p = doc_current.add_paragraph()
        para_spacing(p, line=1.0, before_pt=0, after_pt=0)
        para_indent(p, 0)
        p.add_run("").font.size = Pt(size)


def add_title_runs(p, text, size, ea, bold=True):
    """标题支持用 | 指定换行位置（题目过长时手动断行，避免末行只剩一两个字）。"""
    for i, part in enumerate(text.split("|")):
        if i:
            p.add_run().add_break()
        add_runs(p, part, size, ea=ea, latin=TNR, base_bold=bold)


def write_cover(meta):
    school = meta.get("college") or "医学信息工程学院"
    blank(3)

    p = doc_current.add_paragraph()
    para_align(p, "center")
    para_spacing(p, line=1.5, before_pt=0, after_pt=12)
    para_indent(p, 0)
    add_runs(p, f"{school}毕业论文（设计）", SZ["cover_school"], ea=KAI, latin=TNR, base_bold=True)

    p = doc_current.add_paragraph()
    para_align(p, "center")
    para_spacing(p, line=1.5, before_pt=18, after_pt=18)
    para_indent(p, 0)
    add_title_runs(p, meta.get("title", "论文题目"), SZ["cover_title"], HEI)

    blank(4)

    fields = [
        ("指导教师姓名职称", meta.get("advisor", "")),
        ("专　业　名　称", meta.get("major", "医学信息工程")),
        ("学生年级姓名", f'{meta.get("grade", "")}  {meta.get("author", "")}'),
        ("论文答辩日期", meta.get("defense_date", "    年    月    日")),
    ]
    for label, value in fields:
        p = doc_current.add_paragraph()
        para_align(p, "center")
        para_spacing(p, line=1.5, before_pt=6, after_pt=6)
        para_indent(p, 0)
        add_runs(p, f"{label}：", SZ["cover_label"], ea=SONG, latin=TNR)
        width = sum(2 if ord(c) > 127 else 1 for c in value)
        pad = " " * max(0, 20 - width)
        add_runs(p, value + pad, SZ["body"], ea=SONG, latin=TNR)
        for r in p.runs[1:]:
            r.font.underline = True

    blank(4)
    p = doc_current.add_paragraph()
    para_align(p, "center")
    para_spacing(p, line=1.5, before_pt=12, after_pt=0)
    para_indent(p, 0)
    add_runs(p, meta.get("date") or "二〇二六年六月", 16.0, ea=KAI, latin=TNR)


def write_inner_title(meta):
    school = meta.get("college") or "医学信息工程学院"
    p = doc_current.add_paragraph()
    para_align(p, "center")
    para_spacing(p, line=1.5, before_pt=0, after_pt=24)
    para_indent(p, 0)
    add_runs(p, f"{school}毕业论文（设计）", SZ["cover_school"], ea=KAI, latin=TNR, base_bold=True)

    p = doc_current.add_paragraph()
    para_align(p, "center")
    para_spacing(p, line=1.5, before_pt=6, after_pt=18)
    para_indent(p, 0)
    add_title_runs(p, meta.get("title", ""), 22.0, HEI)
    p = doc_current.add_paragraph()
    para_align(p, "center")
    para_spacing(p, line=1.5, before_pt=6, after_pt=0)
    para_indent(p, 0)
    add_title_runs(p, meta.get("title_en", ""), 16.0, TNR)

    blank(6)
    p = doc_current.add_paragraph()
    para_align(p, "center")
    para_spacing(p, line=1.5, before_pt=12, after_pt=0)
    para_indent(p, 0)
    add_runs(p, meta.get("date") or "二〇二六年六月", 16.0, ea=KAI, latin=TNR)


def write_authorization(meta):
    p = doc_current.add_paragraph()
    para_align(p, "center")
    para_spacing(p, line=1.5, before_pt=0, after_pt=24)
    para_indent(p, 0)
    add_runs(p, "关于论文使用授权的说明", SZ["h1"], ea=HEI, latin=TNR)

    body = ("本人完全了解甘肃中医药大学有关保留、使用学位论文的规定，即：学校有权保留送交论文的复印件，"
            "允许论文被查阅和借阅；学校可以公布论文的全部或部分内容，可以采用影印、缩印或其他复制手段保存论文。")
    p = doc_current.add_paragraph()
    para_align(p, "justify")
    para_spacing(p, line=1.5, before_pt=0, after_pt=0)
    para_indent(p, 2)
    add_runs(p, body, SZ["body"], ea=SONG, latin=TNR)

    blank(3)
    for label in ("作者签名：", "导师签名："):
        p = doc_current.add_paragraph()
        para_align(p, "left")
        para_spacing(p, line=1.5, before_pt=12, after_pt=12)
        para_indent(p, 0)
        add_runs(p, label + " " * 28, SZ["body"], ea=SONG, latin=TNR)
        add_runs(p, "日期：    年    月    日", SZ["body"], ea=SONG, latin=TNR)


def render_toc_title(first_in_section=False):
    """目录标题用内置「TOC Heading」样式（大纲级别=正文），避免目录里出现“目录”自己。"""
    p = doc_current.add_paragraph(style="TOC Heading")
    p.paragraph_format.page_break_before = not first_in_section
    para_align(p, "center")
    para_spacing(p, line=1.5, before_lines=1, after_lines=1)
    para_indent(p, 0)
    p.add_run("目　　录")
    return p


def insert_toc_field():
    """在当前位置插入 TOC 域，Word 打开时可一键更新出真目录。"""
    p = doc_current.add_paragraph()
    para_indent(p, 0)
    para_spacing(p, line=1.5, before_pt=0, after_pt=0)

    r1 = p.add_run()
    style_run(r1, SZ["body"], ea=SONG, latin=TNR)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = r' TOC \o "1-3" \h \z \u '
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    r1._element.append(begin)
    r1._element.append(instr)
    r1._element.append(sep)

    r2 = p.add_run("（在 Word 中全选后按 F9 更新域即可生成目录，本文件已设为打开时自动更新）")
    style_run(r2, SZ["small"], ea=SONG, latin=TNR)

    r3 = p.add_run()
    style_run(r3, SZ["body"], ea=SONG, latin=TNR)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    r3._element.append(end)


def enable_update_fields(doc):
    settings = doc.settings.element
    el = settings.find(qn("w:updateFields"))
    if el is None:
        el = OxmlElement("w:updateFields")
        settings.append(el)
    el.set(qn("w:val"), "true")


# ---------------------------------------------------------------- Markdown 解析
CAP_FIG = re.compile(r"^图\s*[A-Z]?\d+\s*[-－.]\s*\d+")
CAP_TAB = re.compile(r"^表\s*[A-Z]?\d+\s*[-－.]\s*\d+")
IMG = re.compile(r"^!\[(?P<cap>[^\]]*)\]\((?P<src>[^)]+)\)\s*$")
REF = re.compile(r"^\[\d+\]")
CHAPTER_NUM = re.compile(r"^(\d+)\s+(.+)$")
TOC_LINE = re.compile(r"^(.*?)\s*[\t\.·…]{1,}\s*[0-9ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+$")


def parse_frontmatter(lines):
    meta = {}
    if not lines or lines[0].strip() != "---":
        return meta, lines
    end = None
    for i in range(1, len(lines)):
        if lines[i].strip() == "---":
            end = i
            break
    if end is None:
        return meta, lines
    for raw in lines[1:end]:
        if ":" not in raw:
            continue
        k, v = raw.split(":", 1)
        v = v.strip().strip('"').strip("'")
        if v:
            meta[k.strip()] = v
    return meta, lines[end + 1:]


def split_blocks(lines):
    blocks = []
    i, n = 0, len(lines)
    while i < n:
        line = lines[i]
        s = line.strip()
        if not s:
            i += 1
            continue

        m = re.match(r"^(#{1,4})\s+(.*)$", s)
        if m:
            blocks.append(("heading", len(m.group(1)), m.group(2).strip()))
            i += 1
            continue

        m = IMG.match(s)
        if m:
            blocks.append(("img", m.group("cap").strip(), m.group("src").strip()))
            i += 1
            continue

        if re.match(r"^([-*]\s+|\d+[\.、]\s+)", s):
            blocks.append(("li", re.sub(r"^([-*]\s+|\d+[\.、]\s+)", "", s).strip()))
            i += 1
            continue

        if s.startswith("```"):
            info = s[3:].strip()
            buf = []
            i += 1
            while i < n and not lines[i].strip().startswith("```"):
                buf.append(lines[i].rstrip("\n"))
                i += 1
            i += 1
            if info.lower().startswith(("formula", "eq")):
                num = ""
                mm = re.search(r"[\(（]([^\)）]+)[\)）]", info)
                if mm:
                    num = mm.group(1)
                blocks.append(("formula", "\n".join(buf).strip(), num))
            else:
                blocks.append(("code", info, buf))
            continue

        if s.startswith("|") and i + 1 < n and re.match(r"^\|?[\s:\-|]+\|?$", lines[i + 1].strip()):
            rows = [[c.strip() for c in s.strip("|").split("|")]]
            i += 2
            while i < n and lines[i].strip().startswith("|"):
                rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            caption = ""
            if blocks and blocks[-1][0] == "p" and CAP_TAB.match(blocks[-1][1]):
                caption = blocks.pop()[1]
            blocks.append(("table", caption, rows))
            continue

        buf = [s]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if (not nxt or nxt.startswith("#") or IMG.match(nxt) or nxt.startswith("```")
                    or re.match(r"^([-*]\s+|\d+[\.、]\s+)", nxt)
                    or (nxt.startswith("|") and i + 1 < n
                        and re.match(r"^\|?[\s:\-|]+\|?$", lines[i + 1].strip()))):
                break
            buf.append(nxt)
            i += 1
        blocks.append(("p", "".join(buf)))
    return blocks


DISP = {
    "摘要": "摘　　要",
    "目录": "目　　录",
    "总结": "总　　结",
    "致谢": "致　　谢",
    "附录": "附　　录",
    "参考文献": "参考文献",
    "Abstract": "Abstract",
}


def is_chapter_heading(text):
    return bool(re.match(r"^(第\s*\d+\s*章|\d+\s+\S)", text))


def render_heading(level, text, first_in_section):
    if level >= 4:
        raise SystemExit(f"规范只允许到三级标题，请把四级标题改为正文编号：{text}")
    p = doc_current.add_paragraph(style=f"Heading {level}")
    if level == 1:
        p.paragraph_format.page_break_before = not first_in_section
        para_align(p, "center")
        para_spacing(p, line=1.5, before_lines=1, after_lines=1)
        para_indent(p, 0)
        m = CHAPTER_NUM.match(text)
        if m and not re.match(r"^第\s*\d+\s*章", text):
            text = f"第{m.group(1)}章 {m.group(2)}"
        key = re.sub(r"\s+", "", text)
        text = DISP.get(key, text)
    else:
        para_align(p, "left")
        para_spacing(p, line=1.5, before_pt=6, after_pt=6)
        para_indent(p, 0)
    # 标题不写字符级格式，全部交给样式：
    # 这样 Word 更新目录时，目录条目才会用 toc 样式（宋体小四）而不是照抄标题的黑体
    p.add_run(text)
    return p


def render_paragraph(text):
    plain = re.sub(r"\*\*", "", text)
    plain = re.sub(r"\$\$(.+?)\$\$", r"\1", plain)
    low = plain.lower()

    is_kw = low.startswith("【关键词】") or low.startswith("【keywords】")
    if is_kw:
        blank(1, SZ["small"])          # 关键词与摘要之间空一行

    p = doc_current.add_paragraph()
    para_align(p, "justify")
    para_spacing(p, line=1.5, before_pt=0, after_pt=0)
    para_indent(p, 0)

    if is_kw:
        if low.startswith("【关键词】"):
            body = plain[len("【关键词】"):].strip()
            add_runs(p, "【关键词】", SZ["body"], ea=HEI, latin=TNR, base_bold=True)
            add_runs(p, body, SZ["body"], ea=SONG, latin=TNR)
        else:
            body = plain[len("【Keywords】"):].strip()
            add_runs(p, "【Keywords】", SZ["body"], ea=HEI, latin=TNR, base_bold=True)
            add_runs(p, body, SZ["body"], ea=TNR, latin=TNR)
        return p
    if REF.match(plain):
        para_align(p, "left")
        add_runs(p, plain, SZ["small"], ea=SONG, latin=TNR)
        return p
    if CAP_FIG.match(plain) or CAP_TAB.match(plain):
        para_align(p, "center")
        add_runs(p, plain, SZ["small"], ea=SONG, latin=TNR)
        return p

    para_indent(p, 2)
    add_runs(p, text, SZ["body"], ea=SONG, latin=TNR)
    return p


def render_list(items):
    for idx, item in enumerate(items):
        p = doc_current.add_paragraph()
        para_align(p, "justify")
        para_spacing(p, line=1.5, before_pt=0, after_pt=0)
        para_indent(p, 2)
        marker = f"（{idx + 1}）" if idx < 20 else f"({idx + 1})"
        if not re.match(r"^[（(①-⑳]", item):
            item = marker + item
        add_runs(p, item, SZ["body"], ea=SONG, latin=TNR)


def render_code(code_lines):
    for line in code_lines:
        p = doc_current.add_paragraph()
        para_align(p, "left")
        para_spacing(p, line=1.0, before_pt=0, after_pt=0)
        para_indent(p, 0)
        add_runs(p, line if line else " ", SZ["small"], ea=SONG, latin="Consolas")


class Renderer:
    def __init__(self, manuscript):
        self.base = os.path.dirname(os.path.abspath(manuscript))
        self.max_fig = 11.0
        self.first = True
        self.pending = []

    def flush(self):
        if self.pending:
            render_list([x for x in self.pending])
            self.pending = []

    def block(self, b):
        kind = b[0]
        if kind == "li":
            self.pending.append(b[1])
            return
        self.flush()
        if kind == "heading":
            render_heading(b[1], b[2], self.first)
            self.first = False
        elif kind == "p":
            render_paragraph(b[1])
        elif kind == "img":
            src = b[2]
            if not os.path.isabs(src):
                src = os.path.join(self.base, src)
            if os.path.exists(src):
                add_figure(src, b[1], self.max_fig)
            else:
                render_paragraph(f"[缺失插图：{src}]")
        elif kind == "table":
            add_table(b[2], b[1])
        elif kind == "formula":
            add_formula(b[1], b[2] or "0-0")
        elif kind == "code":
            render_code(b[2])


# ---------------------------------------------------------------- 主流程
def build(manuscript, out_path, max_figure_cm=11.0):
    global doc_current
    with open(manuscript, "r", encoding="utf-8") as f:
        raw = f.read().replace("\r\n", "\n")
    lines = raw.split("\n")
    meta, rest = parse_frontmatter(lines)
    blocks = split_blocks(rest)

    doc = Document()
    doc_current = doc
    setup_styles(doc)
    config_section(doc.sections[0], "upperRoman", 1, header_text=None, page_number=False)

    write_cover(meta)

    config_section(doc.add_section(WD_SECTION.NEW_PAGE), "upperRoman", 1,
                   header_text=None, page_number=False)
    write_inner_title(meta)

    config_section(doc.add_section(WD_SECTION.NEW_PAGE), "upperRoman", 1,
                   header_text=None, page_number=False)
    write_authorization(meta)

    # 前置部分（摘要 / Abstract / 目录），罗马页码
    config_section(doc.add_section(WD_SECTION.NEW_PAGE), "upperRoman", 1,
                   header_text=None, page_number=True)

    body_start = None
    for idx, b in enumerate(blocks):
        if b[0] == "heading" and b[1] == 1 and is_chapter_heading(b[2]):
            body_start = idx
            break
    front = blocks if body_start is None else blocks[:body_start]
    body = [] if body_start is None else blocks[body_start:]

    renderer = Renderer(manuscript)
    renderer.max_fig = max_figure_cm

    toc_seen = False
    for b in front:
        if toc_seen:
            break
        if b[0] == "heading" and b[1] == 1 and re.sub(r"\s+", "", b[2]) == "目录":
            render_toc_title(renderer.first)
            renderer.first = False
            insert_toc_field()
            toc_seen = True
            continue
        renderer.block(b)
    if not toc_seen:
        render_toc_title(False)
        insert_toc_field()

    # 正文（阿拉伯页码，从 1 开始）
    header_text = meta.get("header") or HEADER_TEXT
    config_section(doc.add_section(WD_SECTION.NEW_PAGE), "decimal", 1,
                   header_text=header_text, page_number=True)
    renderer.first = True
    for b in body:
        renderer.block(b)
    renderer.flush()

    enable_update_fields(doc)
    kill_section_break_paragraphs(doc)
    doc.save(out_path)
    print(f"已生成：{out_path}")


def build_skeleton(out_path):
    global doc_current
    doc = Document()
    doc_current = doc
    setup_styles(doc)
    meta = {
        "title": "论文题目（不超过30字，可两行）",
        "title_en": "English Title of the Thesis",
        "advisor": "×××  教授",
        "major": "医学信息工程",
        "grade": "2022级",
        "author": "×××",
        "date": "二〇二六年六月",
    }
    config_section(doc.sections[0], "upperRoman", 1, header_text=None, page_number=False)
    write_cover(meta)
    config_section(doc.add_section(WD_SECTION.NEW_PAGE), "upperRoman", 1,
                   header_text=None, page_number=False)
    write_inner_title(meta)
    config_section(doc.add_section(WD_SECTION.NEW_PAGE), "upperRoman", 1,
                   header_text=None, page_number=False)
    write_authorization(meta)
    config_section(doc.add_section(WD_SECTION.NEW_PAGE), "upperRoman", 1,
                   header_text=None, page_number=True)

    render_heading(1, "摘    要", True)
    render_paragraph("（中文摘要 300～500 字，含目的、方法、结果、结论。宋体小四，首行缩进 2 字符，1.5 倍行距。）")
    render_paragraph("【关键词】关键词一；关键词二；关键词三；关键词四；关键词五")
    render_heading(1, "Abstract", False)
    render_paragraph("（英文摘要与中文摘要完全对应，Times New Roman 小四，首行缩进 2 字符。）")
    render_paragraph("【Keywords】Keyword One  Keyword Two  Keyword Three  Keyword Four")
    p_toc = render_toc_title(False)
    insert_toc_field()

    config_section(doc.add_section(WD_SECTION.NEW_PAGE), "decimal", 1,
                   header_text=HEADER_TEXT, page_number=True)
    render_heading(1, "第1章 引言", True)
    render_heading(2, "1.1 研究背景与意义", False)
    render_paragraph("（正文：宋体小四，英文与数字 Times New Roman，首行缩进 2 字符，1.5 倍行距，段前段后 0。）")
    render_heading(3, "1.1.1 三级标题示例", False)
    render_paragraph("（最多三级标题，不得出现四级标题。）")
    render_heading(1, "总　　结", False)
    render_paragraph("（总结研究结果、创新点与不足。）")
    render_heading(1, "致　　谢", False)
    render_paragraph("（对指导教师及提供帮助的组织和个人致谢。）")
    render_heading(1, "参考文献", False)
    render_paragraph("[1] 吴功宜.计算机网络(第三版)[M].北京:清华大学出版社,2011:5-26.")
    render_heading(1, "附　　录", False)
    render_heading(3, "附录A 核心代码", False)
    render_paragraph("（正文中无法呈现的图、表、数据与程序代码，图表需单独编号，如附图1、附表1。）")

    enable_update_fields(doc)
    kill_section_break_paragraphs(doc)
    doc.save(out_path)
    print(f"已生成空白规范骨架：{out_path}")


def main():
    ap = argparse.ArgumentParser(description="按甘肃中医药大学医学信息工程学院本科毕业论文规范排版")
    ap.add_argument("manuscript", nargs="?", help="Markdown 稿件路径")
    ap.add_argument("-o", "--output", help="输出 .docx 路径")
    ap.add_argument("--max-figure-cm", type=float, default=11.0, help="插图最大宽度（厘米），默认 11")
    ap.add_argument("--skeleton", metavar="OUT.docx", help="仅生成空白规范骨架")
    args = ap.parse_args()

    if args.skeleton:
        build_skeleton(args.skeleton)
        return
    if not args.manuscript or not args.output:
        ap.error("需要同时提供稿件路径和 -o 输出路径")
    build(args.manuscript, args.output, args.max_figure_cm)


if __name__ == "__main__":
    main()
