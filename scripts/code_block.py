#!/usr/bin/env python3
"""code_block.py - 按 templates/code_template.docx 复刻的代码展示块。

结构：两列表格，左列为 Word 自动编号的行号（01. 02. ...），右列为一行代码。
  - 行号列底色 E7E5DC，宽 709 twips；代码列宽 7587 twips（合计 8296）。
  - 无外框和横线，仅保留列间竖线（颜色/粗细沿用模板）。
  - 代码列偶数行底色 F8F8F8；行高 atLeast 18pt；字体 Courier New 9pt。
  - 代码行使用独立样式 "Code Line"，不继承正文首行缩进（模板中的缩进属误用，不复刻）。
  - 每个代码块新建独立的 w:num 并 startOverride=1，行号从 01 重新开始。
  - 每个 run 都显式写入 ascii/hAnsi/eastAsia/cs 四个字体属性（经 docx_styles.set_run_font）。
  - 语法高亮由 Pygments 完成，配色沿用模板的 GitHub 风格。

用法：
    from code_block import add_code_block
    add_code_block(doc, source_text, language="java", caption="代码 2-1 学生类定义")
"""
from __future__ import annotations

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt

from docx_styles import (
    FONT_CODE, FONT_SONG, HALF_LINE_PT, STYLE_CODE_CAPTION, STYLE_CODE_LINE,
    STYLE_CODE_NUMBER, add_text, set_run_font,
)

# --------------------------------------------------------------------------
# 模板参数
# --------------------------------------------------------------------------
NUM_COL_TWIPS = 709
CODE_COL_TWIPS = 7587
NUM_FILL = "E7E5DC"
ALT_FILL = "F8F8F8"
DIVIDER_COLOR = "65E266"      # 模板中的列间竖线颜色
DIVIDER_SZ = 18               # 1/8 pt，即 2.25pt
CODE_SIZE_PT = 9
CODE_ABSTRACT_ID = 90         # 自建 abstractNum 的 id，避免与已有编号冲突
TAB_SPACES = 4

# Pygments token → 颜色（模板 GitHub 风格）
DEFAULT_COLOR = "202020"
_TOKEN_COLORS = [
    ("Comment", "6E7781"),
    ("Keyword.Type", "267F99"),
    ("Keyword", "CF222E"),
    ("Literal.String", "953800"),
    ("Literal.Number", "098658"),
    ("Name.Class", "267F99"),
    ("Name.Decorator", "8250DF"),
    ("Name.Function", "8250DF"),
    ("Name.Builtin", "0000FF"),
    ("Name.Exception", "267F99"),
    ("Operator.Word", "CF222E"),
]


def _color_for(token_type) -> str:
    name = str(token_type).replace("Token.", "")
    for prefix, color in _TOKEN_COLORS:
        if name == prefix or name.startswith(prefix + "."):
            return color
    return DEFAULT_COLOR


def _tokenize_lines(source: str, language: str | None) -> list[list[tuple[str, str]]]:
    """返回按行分组的 (文本, 颜色) 列表。Pygments 缺失或语言未知时退化为单色。"""
    source = source.replace("\r\n", "\n").replace("\r", "\n").replace("\t", " " * TAB_SPACES)
    lines: list[list[tuple[str, str]]] = [[]]
    try:
        from pygments import lex
        from pygments.lexers import TextLexer, get_lexer_by_name
        try:
            lexer = get_lexer_by_name(language or "text")
        except Exception:
            lexer = TextLexer()
        stream = [(_color_for(tt), val) for tt, val in lex(source, lexer)]
    except ImportError:
        stream = [(DEFAULT_COLOR, source)]
    for color, val in stream:
        parts = val.split("\n")
        for i, part in enumerate(parts):
            if i > 0:
                lines.append([])
            if part:
                lines[-1].append((part, color))
    # Pygments 会在末尾补换行，去掉由此产生的多余空行
    while len(lines) > 1 and not lines[-1]:
        lines.pop()
    return lines


# --------------------------------------------------------------------------
# 编号（行号）
# --------------------------------------------------------------------------
def _numbering_root(doc):
    return doc.part.numbering_part.element


def _ensure_abstract_num(root) -> None:
    for an in root.findall(qn("w:abstractNum")):
        if an.get(qn("w:abstractNumId")) == str(CODE_ABSTRACT_ID):
            return
    xml = (
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{CODE_ABSTRACT_ID}">'
        '<w:multiLevelType w:val="singleLevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimalZero"/>'
        '<w:suff w:val="nothing"/><w:lvlText w:val="%1."/><w:lvlJc w:val="center"/>'
        '<w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr>'
        f'<w:rPr><w:rFonts w:ascii="{FONT_CODE}" w:hAnsi="{FONT_CODE}" '
        f'w:eastAsia="{FONT_SONG}" w:cs="{FONT_CODE}"/>'
        f'<w:color w:val="000000"/><w:sz w:val="{CODE_SIZE_PT * 2}"/>'
        f'<w:szCs w:val="{CODE_SIZE_PT * 2}"/></w:rPr></w:lvl>'
        '</w:abstractNum>'
    )
    el = parse_xml(xml)
    first_num = root.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(el)      # abstractNum 必须位于所有 num 之前
    else:
        root.append(el)


def _new_num_id(root) -> int:
    """为一个代码块新建 w:num（startOverride=1），返回 numId。"""
    _ensure_abstract_num(root)
    ids = [int(n.get(qn("w:numId"))) for n in root.findall(qn("w:num"))]
    num_id = max(ids + [0]) + 1
    xml = (
        f'<w:num {nsdecls("w")} w:numId="{num_id}">'
        f'<w:abstractNumId w:val="{CODE_ABSTRACT_ID}"/>'
        '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>'
        '</w:num>'
    )
    root.append(parse_xml(xml))
    return num_id


# --------------------------------------------------------------------------
# 表格
# --------------------------------------------------------------------------
def _spacer(doc, pt: float) -> None:
    """插入固定高度的空段落，用于表格前后的 0.5 行间距。"""
    p = doc.add_paragraph()
    ppr = p._p.get_or_add_pPr()
    spacing = ppr.get_or_add_spacing()
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), str(int(pt * 20)))
    spacing.set(qn("w:lineRule"), "exact")


def _tbl_pr_xml() -> str:
    total = NUM_COL_TWIPS + CODE_COL_TWIPS
    return (
        f'<w:tblPr {nsdecls("w")}>'
        f'<w:tblW w:w="{total}" w:type="dxa"/><w:jc w:val="center"/>'
        '<w:tblBorders>'
        '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideV w:val="single" w:sz="{DIVIDER_SZ}" w:space="0" w:color="{DIVIDER_COLOR}"/>'
        '</w:tblBorders>'
        '<w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
        '</w:tblPr>'
    )


def _tc_pr(width: int, fill: str | None) -> object:
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ""
    return parse_xml(
        f'<w:tcPr {nsdecls("w")}><w:tcW w:w="{width}" w:type="dxa"/>{shd}'
        '<w:vAlign w:val="center"/></w:tcPr>'
    )


def _replace_tc_pr(cell, width: int, fill: str | None) -> None:
    tc = cell._tc
    old = tc.find(qn("w:tcPr"))
    if old is not None:
        tc.remove(old)
    tc.insert(0, _tc_pr(width, fill))


def add_code_block(doc, source: str, language: str | None = None,
                   caption: str | None = None) -> None:
    """向文档追加一个代码块。caption 若给出，则置于代码块上方（代码题注）。"""
    lines = _tokenize_lines(source, language)
    if not lines:
        lines = [[]]

    if caption:
        cp = doc.add_paragraph(style=STYLE_CODE_CAPTION)
        add_text(cp, caption, "body", size=10.5)
    else:
        _spacer(doc, HALF_LINE_PT)

    root = _numbering_root(doc)
    num_id = _new_num_id(root)

    table = doc.add_table(rows=len(lines), cols=2)
    tbl = table._tbl
    old = tbl.find(qn("w:tblPr"))
    new_pr = parse_xml(_tbl_pr_xml())
    if old is not None:
        tbl.replace(old, new_pr)
    else:
        tbl.insert(0, new_pr)
    grid = tbl.find(qn("w:tblGrid"))
    for col, w in zip(grid.findall(qn("w:gridCol")), (NUM_COL_TWIPS, CODE_COL_TWIPS)):
        col.set(qn("w:w"), str(w))

    styles = doc.styles
    for idx, (row, tokens) in enumerate(zip(table.rows, lines)):
        # 一行代码不跨页拆开
        trpr = row._tr.get_or_add_trPr()
        trpr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        num_cell, code_cell = row.cells
        _replace_tc_pr(num_cell, NUM_COL_TWIPS, NUM_FILL)
        _replace_tc_pr(code_cell, CODE_COL_TWIPS, ALT_FILL if idx % 2 == 1 else None)

        np_ = num_cell.paragraphs[0]
        np_.style = styles[STYLE_CODE_NUMBER]
        num_pr = np_._p.get_or_add_pPr().get_or_add_numPr()
        num_pr.get_or_add_ilvl().val = 0
        num_pr.get_or_add_numId().val = num_id

        cp_ = code_cell.paragraphs[0]
        cp_.style = styles[STYLE_CODE_LINE]
        for text, color in tokens:
            run = cp_.add_run(text)
            set_run_font(run, "code", size=CODE_SIZE_PT, color=color)

    _spacer(doc, HALF_LINE_PT)
