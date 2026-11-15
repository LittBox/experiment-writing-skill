#!/usr/bin/env python3
"""docx_styles.py - 计算机实验报告 DOCX 的样式定义（强制排版规范的唯一来源）。

职责：
  1. 提供字体、间距、缩进等常量（validate_docx.py 与测试共用同一份常量）。
  2. new_document()：创建 A4 文档，并彻底重写样式表：
       - docDefaults / Normal / 标题 / 题注 / 表格文字 / 代码等样式全部显式写出
         ascii、hAnsi、eastAsia、cs 四个字体属性；
       - 清除所有 asciiTheme / hAnsiTheme / eastAsiaTheme / cstheme 主题字体属性；
       - 清除样式中的主题色和蓝色（标题、超链接等）；
       - 删除 python-docx 默认模板自带的彩色表格样式。
  3. 提供 run 级字体辅助函数，供 table_builder / code_block / build_docx 使用。

规范数值见 references/word-format-spec.md。
"""
from __future__ import annotations

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Twips

# --------------------------------------------------------------------------
# 字体常量
# --------------------------------------------------------------------------
FONT_SONG = "宋体"              # 中文正文
FONT_HEI = "黑体"               # 标题、表头
FONT_WEST = "Times New Roman"   # 西文
FONT_CODE = "Courier New"       # 代码（Windows 与 macOS 均自带）

# 允许出现在文档中的字体（含常见别名）。其余一律视为违规。
ALLOWED_FONTS = {
    FONT_SONG, "SimSun", FONT_HEI, "SimHei", FONT_WEST, FONT_CODE,
}

# 明令禁止的字体（统一小写比较，见 is_forbidden_font）。
FORBIDDEN_FONTS = {
    "楷体", "楷体_gb2312", "楷体-简", "楷体-繁", "kaiti", "kaiti sc", "kaiti tc",
    "kaiti_gb2312", "stkaiti", "stkaiti sc", "kai", "dfkai-sb", "标楷体",
    "微软雅黑", "微软雅黑 light", "microsoft yahei", "microsoft yahei ui",
    "microsoftyahei", "microsoft yahei light", "msyh",
}

# 主题字体属性（任何出现都视为违规）。
THEME_FONT_ATTRS = ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme")
FONT_ATTRS = ("ascii", "hAnsi", "eastAsia", "cs")


def is_forbidden_font(name: str | None) -> bool:
    if not name:
        return False
    n = name.strip().lower()
    return n in FORBIDDEN_FONTS or "kaiti" in n or "yahei" in n or "楷体" in n or "雅黑" in n


# --------------------------------------------------------------------------
# 间距、缩进、字号常量
# --------------------------------------------------------------------------
BODY_SIZE_PT = 12                 # 小四
HALF_LINE_PT = 6                  # 0.5 行 = 6pt（以 12pt 正文为基准）
BODY_LINE_MULTIPLE = 1.25         # 正文 1.25 倍行距（w:line=300, lineRule=auto）
BODY_INDENT_CHARS = 2             # 首行缩进 2 字符
CAPTION_GAP_PT = 3                # 题注与图/表之间的距离
CAPTION_SIZE_PT = 10.5            # 五号

PAGE_W_TWIPS, PAGE_H_TWIPS = 11906, 16838   # A4
MARGIN_TB_CM, MARGIN_LR_CM = 2.54, 3.17
HEADER_DIST_TWIPS, FOOTER_DIST_TWIPS = 851, 992

# 样式名（python-docx 由名称推导 styleId：去掉空格）
STYLE_BODY = "Report Body"
STYLE_TABLE_TEXT = "Table Text"
STYLE_TABLE_HEADER = "Table Header"
STYLE_TABLE_CAPTION = "Table Caption"
STYLE_FIGURE = "Figure"
STYLE_FIGURE_CAPTION = "Figure Caption"
STYLE_CODE_LINE = "Code Line"
STYLE_CODE_NUMBER = "Code Line Number"
STYLE_CODE_CAPTION = "Code Caption"
STYLE_COVER_TITLE = "Title"


# --------------------------------------------------------------------------
# 底层 XML 辅助
# --------------------------------------------------------------------------
def set_rfonts(rpr, ascii_: str, east_asia: str, *, h_ansi: str | None = None,
               cs: str | None = None) -> None:
    """在 rPr 上写入显式字体：先清空 rFonts 的全部属性（含主题属性）再设置四个属性。"""
    rfonts = rpr.get_or_add_rFonts()
    for key in list(rfonts.attrib):
        del rfonts.attrib[key]
    rfonts.set(qn("w:ascii"), ascii_)
    rfonts.set(qn("w:hAnsi"), h_ansi or ascii_)
    rfonts.set(qn("w:eastAsia"), east_asia)
    rfonts.set(qn("w:cs"), cs or ascii_)


def _set_sz(rpr, size_pt: float) -> None:
    half_points = str(int(round(size_pt * 2)))
    sz = rpr.find(qn("w:sz"))
    if sz is None:
        # 借助 python-docx 的 sz_val 属性保证子元素顺序正确
        from docx.shared import Pt as _Pt
        rpr.sz_val = _Pt(size_pt)
        sz = rpr.find(qn("w:sz"))
    sz.set(qn("w:val"), half_points)
    szcs = rpr.find(qn("w:szCs"))
    if szcs is None:
        szcs = OxmlElement("w:szCs")
        sz.addnext(szcs)
    szcs.set(qn("w:val"), half_points)


def _set_ind(ppr, *, first_line_chars: int = 0, char_pt: float = BODY_SIZE_PT) -> None:
    """设置首行缩进。chars=0 表示明确的“无缩进”（同时清除 left/hanging）。"""
    ind = ppr.get_or_add_ind()
    for key in list(ind.attrib):
        del ind.attrib[key]
    ind.set(qn("w:left"), "0")
    ind.set(qn("w:right"), "0")
    ind.set(qn("w:firstLineChars"), str(first_line_chars * 100))
    ind.set(qn("w:firstLine"), str(int(first_line_chars * char_pt * 20)))


def _strip_color(rpr) -> None:
    for tag in ("w:color",):
        for el in rpr.findall(qn(tag)):
            rpr.remove(el)


def _make_style(styles, name: str, base: str | None = "Normal"):
    try:
        st = styles[name]
    except KeyError:
        st = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    if base:
        st.base_style = styles[base]
    st.quick_style = True
    return st


def _configure(style, *, east: str, ascii_: str, size: float, bold: bool = False,
               align=None, before: float = 0, after: float = 0,
               line: float | str = 1.0, indent_chars: int = 0,
               keep_next: bool = False, keep_lines: bool = False) -> None:
    """统一配置一个段落样式的字体、字号、对齐、间距、缩进。"""
    rpr = style.element.get_or_add_rPr()
    set_rfonts(rpr, ascii_, east)
    style.font.size = Pt(size)
    _set_sz(rpr, size)
    style.font.bold = bold
    style.font.italic = False
    _strip_color(rpr)
    style.font.color.rgb = RGBColor(0, 0, 0)

    pf = style.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if isinstance(line, str) and line.startswith("atLeast:"):
        pf.line_spacing = Pt(float(line.split(":", 1)[1]))
        from docx.enum.text import WD_LINE_SPACING
        pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
    else:
        pf.line_spacing = float(line)
    pf.alignment = align
    pf.keep_with_next = keep_next
    pf.keep_together = keep_lines
    pf.widow_control = True
    _set_ind(style.element.get_or_add_pPr(), first_line_chars=indent_chars, char_pt=size)


# --------------------------------------------------------------------------
# 样式表重写
# --------------------------------------------------------------------------
def _rewrite_doc_defaults(styles) -> None:
    defaults = styles.element.find(qn("w:docDefaults"))
    if defaults is None:
        defaults = OxmlElement("w:docDefaults")
        styles.element.insert(0, defaults)
    for child in list(defaults):
        defaults.remove(child)
    rpr_default = OxmlElement("w:rPrDefault")
    rpr = OxmlElement("w:rPr")
    rpr_default.append(rpr)
    defaults.append(rpr_default)
    ppr_default = OxmlElement("w:pPrDefault")
    defaults.append(ppr_default)

    # rPr 必须是 CT_RPr 才有 get_or_add_*，因此经由 parse 后的元素重新取得
    rpr = defaults.find(qn("w:rPrDefault")).find(qn("w:rPr"))
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    new_rpr = parse_xml(
        f'<w:rPr {nsdecls("w")}>'
        f'<w:rFonts w:ascii="{FONT_WEST}" w:hAnsi="{FONT_WEST}" '
        f'w:eastAsia="{FONT_SONG}" w:cs="{FONT_WEST}"/>'
        f'<w:sz w:val="{BODY_SIZE_PT * 2}"/><w:szCs w:val="{BODY_SIZE_PT * 2}"/>'
        f'<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>'
        f'</w:rPr>'
    )
    rpr_default = defaults.find(qn("w:rPrDefault"))
    rpr_default.replace(rpr, new_rpr)


def _scrub_all_styles(styles) -> None:
    """清除所有样式里的主题字体、主题色，删除自带的彩色表格样式。"""
    root = styles.element
    for st in list(root.findall(qn("w:style"))):
        st_type = st.get(qn("w:type"))
        st_id = st.get(qn("w:styleId"))
        if st_type == "table":
            if st_id not in ("TableNormal", "TableGrid"):
                root.remove(st)
            continue
        rpr = st.find(qn("w:rPr"))
        if rpr is None:
            continue
        _strip_color(rpr)
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is not None:
            is_heading = bool(st_id) and (st_id.startswith("Heading") or st_id == "Title"
                                          or st_id == "Subtitle")
            ascii_, east = (FONT_HEI, FONT_HEI) if is_heading else (FONT_WEST, FONT_SONG)
            set_rfonts(rpr, ascii_, east)
        # 清除下划线之类的蓝色超链接残留
        for u in rpr.findall(qn("w:u")):
            rpr.remove(u)


def build_styles(doc) -> None:
    styles = doc.styles
    _rewrite_doc_defaults(styles)
    _scrub_all_styles(styles)

    # Normal：作为所有非正文样式的基底，不带缩进、不带段间距
    normal = styles["Normal"]
    _configure(normal, east=FONT_SONG, ascii_=FONT_WEST, size=BODY_SIZE_PT,
               before=0, after=0, line=1.0, indent_chars=0)

    # 正文：1.25 倍行距，段前后 6pt，首行缩进 2 字符
    body = _make_style(styles, STYLE_BODY)
    _configure(body, east=FONT_SONG, ascii_=FONT_WEST, size=BODY_SIZE_PT,
               align=WD_ALIGN_PARAGRAPH.JUSTIFY,
               before=HALF_LINE_PT, after=HALF_LINE_PT,
               line=BODY_LINE_MULTIPLE, indent_chars=BODY_INDENT_CHARS)

    # 标题：黑体、黑色、无缩进、左对齐
    heading_specs = {
        "Heading 1": (16, True, 12, 6),    # 三号
        "Heading 2": (14, True, 12, 6),    # 四号
        "Heading 3": (12, True, 6, 6),     # 小四
    }
    for name, (size, bold, before, after) in heading_specs.items():
        st = styles[name]
        _configure(st, east=FONT_HEI, ascii_=FONT_HEI, size=size, bold=bold,
                   align=WD_ALIGN_PARAGRAPH.LEFT, before=before, after=after,
                   line=BODY_LINE_MULTIPLE, indent_chars=0,
                   keep_next=True, keep_lines=True)

    # 封面标题
    title = styles["Title"]
    _configure(title, east=FONT_HEI, ascii_=FONT_HEI, size=22, bold=True,
               align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=12,
               line=1.25, indent_chars=0)
    # 默认 Title 样式带下边框，去掉
    ppr = title.element.get_or_add_pPr()
    for b in ppr.findall(qn("w:pBdr")):
        ppr.remove(b)

    # 表格内文字：不继承正文缩进，单倍行距
    tt = _make_style(styles, STYLE_TABLE_TEXT)
    _configure(tt, east=FONT_SONG, ascii_=FONT_WEST, size=BODY_SIZE_PT,
               align=WD_ALIGN_PARAGRAPH.LEFT, before=0, after=0, line=1.0)
    th = _make_style(styles, STYLE_TABLE_HEADER)
    _configure(th, east=FONT_HEI, ascii_=FONT_HEI, size=BODY_SIZE_PT,
               align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0, line=1.0)

    # 题注：表题在表上、图题在图下
    tc = _make_style(styles, STYLE_TABLE_CAPTION)
    _configure(tc, east=FONT_SONG, ascii_=FONT_WEST, size=CAPTION_SIZE_PT,
               align=WD_ALIGN_PARAGRAPH.CENTER,
               before=HALF_LINE_PT, after=CAPTION_GAP_PT, line=1.0, keep_next=True)
    fig = _make_style(styles, STYLE_FIGURE)
    _configure(fig, east=FONT_SONG, ascii_=FONT_WEST, size=BODY_SIZE_PT,
               align=WD_ALIGN_PARAGRAPH.CENTER,
               before=HALF_LINE_PT, after=CAPTION_GAP_PT, line=1.0, keep_next=True)
    fc = _make_style(styles, STYLE_FIGURE_CAPTION)
    _configure(fc, east=FONT_SONG, ascii_=FONT_WEST, size=CAPTION_SIZE_PT,
               align=WD_ALIGN_PARAGRAPH.CENTER,
               before=0, after=HALF_LINE_PT, line=1.0)
    cc = _make_style(styles, STYLE_CODE_CAPTION)
    _configure(cc, east=FONT_SONG, ascii_=FONT_WEST, size=CAPTION_SIZE_PT,
               align=WD_ALIGN_PARAGRAPH.CENTER,
               before=HALF_LINE_PT, after=CAPTION_GAP_PT, line=1.0, keep_next=True)

    # 代码：独立样式，无首行缩进（模板中的缩进是误用，不复刻）
    cl = _make_style(styles, STYLE_CODE_LINE)
    _configure(cl, east=FONT_SONG, ascii_=FONT_CODE, size=9,
               align=WD_ALIGN_PARAGRAPH.LEFT, before=0, after=0,
               line="atLeast:18", indent_chars=0)
    cn = _make_style(styles, STYLE_CODE_NUMBER)
    _configure(cn, east=FONT_SONG, ascii_=FONT_CODE, size=9,
               align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0,
               line="atLeast:18", indent_chars=0)

    # 页眉页脚
    for name in ("Header", "Footer"):
        st = _make_style(styles, name)
        _configure(st, east=FONT_SONG, ascii_=FONT_WEST, size=9,
                   align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0, line=1.0)


def apply_page_setup(doc) -> None:
    for sec in doc.sections:
        sec.page_width = Twips(PAGE_W_TWIPS)
        sec.page_height = Twips(PAGE_H_TWIPS)
        sec.top_margin = sec.bottom_margin = Cm(MARGIN_TB_CM)
        sec.left_margin = sec.right_margin = Cm(MARGIN_LR_CM)
        sec.header_distance = Twips(HEADER_DIST_TWIPS)
        sec.footer_distance = Twips(FOOTER_DIST_TWIPS)


def new_document():
    """创建符合规范的空白文档（A4、样式已重写、无主题字体、无蓝色）。"""
    doc = Document()
    # 清空默认模板可能带来的占位段落
    body = doc.element.body
    for p in list(body.findall(qn("w:p"))):
        body.remove(p)
    build_styles(doc)
    apply_page_setup(doc)
    return doc


# --------------------------------------------------------------------------
# run 级辅助
# --------------------------------------------------------------------------
_ROLE_FONTS = {
    "body": (FONT_WEST, FONT_SONG),
    "heading": (FONT_HEI, FONT_HEI),
    "code": (FONT_CODE, FONT_SONG),
}


def set_run_font(run, role: str = "body", size: float | None = None,
                 bold: bool | None = None, color: str | None = None) -> None:
    """给 run 显式写入四个字体属性。所有手工创建的 run 都必须经过它。"""
    ascii_, east = _ROLE_FONTS[role]
    rpr = run._r.get_or_add_rPr()
    set_rfonts(rpr, ascii_, east)
    if size is not None:
        _set_sz(rpr, size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color.upper())


def add_text(paragraph, text: str, role: str = "body", **kw):
    run = paragraph.add_run(text)
    set_run_font(run, role, **kw)
    return run


def add_page_number_field(paragraph) -> None:
    """在页脚段落中插入 PAGE 域，字体显式指定。"""
    def _fld(kind):
        r = paragraph.add_run()
        set_run_font(r, "body", size=9)
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), kind)
        r._r.append(el)

    _fld("begin")
    r = paragraph.add_run()
    set_run_font(r, "body", size=9)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    r._r.append(instr)
    _fld("separate")
    add_text(paragraph, "1", "body", size=9)
    _fld("end")
