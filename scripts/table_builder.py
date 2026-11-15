#!/usr/bin/env python3
"""table_builder.py - 按 excel_template.docx 生成两类表格。

表格一（非数据类：配置项、环境说明、需求说明等）
  - 全框线 0.5pt 黑色；表头灰底 D7D7D7、黑体加粗居中；单元格边距 80 twips；
  - 表格居中、垂直居中；不使用浮动定位（模板中的 tblpPr 是误用，不复刻）。
表格二（数据类：测试数据、实验结果、性能对比等），三线表
  - 无竖线、无内部横线；表头上边线 3pt、表头下边线 1pt、表尾下边线 3pt；
  - 表头白底黑体不加粗居中；跨页重复表头；表头行最小高 909 twips。

表内段落一律使用独立样式 Table Text / Table Header（无首行缩进），
每个 run 显式写入 ascii/hAnsi/eastAsia/cs 字体。
"""
from __future__ import annotations

from typing import Sequence

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Twips

from docx_styles import (
    STYLE_TABLE_CAPTION, STYLE_TABLE_HEADER, STYLE_TABLE_TEXT,
    add_text,
)

TEXT_WIDTH_TWIPS = 8306            # 版心宽度（A4，左右边距 3.17cm）
HEADER_FILL_TYPE1 = "D7D7D7"       # 表格一表头灰底（中性灰，非彩色）
CELL_MARGIN_TYPE1 = 80
TYPE2_HEADER_MIN_HEIGHT = 909


# --------------------------------------------------------------------------
# XML 辅助
# --------------------------------------------------------------------------
def _child(parent, tag: str):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el


def _set_borders(container, spec: dict[str, tuple[str, int]], tag: str) -> None:
    """spec: {'top': ('single', 4), ...}；写入 tblBorders 或 tcBorders。"""
    borders = _child(container, tag)
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if side not in spec:
            continue
        val, sz = spec[side]
        el = borders.find(qn(f"w:{side}"))
        if el is None:
            el = OxmlElement(f"w:{side}")
            borders.append(el)
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")


def _set_tbl_pr(table, *, widths: Sequence[int], borders: dict, cell_margin: int) -> None:
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    # 去掉可能带入的样式引用，避免继承默认样式的颜色
    for el in tbl_pr.findall(qn("w:tblStyle")):
        tbl_pr.remove(el)
    total = sum(widths)
    tbl_w = _child(tbl_pr, "w:tblW")
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")
    jc = _child(tbl_pr, "w:jc")
    jc.set(qn("w:val"), "center")
    _set_borders(tbl_pr, borders, "w:tblBorders")
    layout = _child(tbl_pr, "w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    mar = _child(tbl_pr, "w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        el = mar.find(qn(f"w:{side}"))
        if el is None:
            el = OxmlElement(f"w:{side}")
            mar.append(el)
        el.set(qn("w:w"), str(cell_margin))
        el.set(qn("w:type"), "dxa")

    grid = tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(w))


def _set_cell(cell, width: int, *, fill: str | None = None,
              margins: tuple[int, int, int, int] | None = None,
              borders: dict | None = None) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tcw = _child(tc_pr, "w:tcW")
    tcw.set(qn("w:w"), str(width))
    tcw.set(qn("w:type"), "dxa")
    if borders:
        _set_borders(tc_pr, borders, "w:tcBorders")
    if fill:
        shd = _child(tc_pr, "w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)
    if margins:
        mar = _child(tc_pr, "w:tcMar")
        for side, w in zip(("top", "left", "bottom", "right"), margins):
            el = mar.find(qn(f"w:{side}"))
            if el is None:
                el = OxmlElement(f"w:{side}")
                mar.append(el)
            el.set(qn("w:w"), str(w))
            el.set(qn("w:type"), "dxa")
    va = _child(tc_pr, "w:vAlign")
    va.set(qn("w:val"), "center")


def _fill_paragraph(cell, text: str, *, header: bool, bold: bool,
                    align=None) -> None:
    p = cell.paragraphs[0]
    p.style = cell.part.document.styles[STYLE_TABLE_HEADER if header else STYLE_TABLE_TEXT]
    if align is not None:
        p.alignment = align
    add_text(p, str(text), "heading" if header else "body",
             bold=bold if header else None)


def _row_flags(row, *, header: bool, min_height: int | None = None) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    cant = OxmlElement("w:cantSplit")
    tr_pr.append(cant)
    if header:
        if min_height:
            h = OxmlElement("w:trHeight")
            h.set(qn("w:val"), str(min_height))
            h.set(qn("w:hRule"), "atLeast")
            tr_pr.append(h)
        tr_pr.append(OxmlElement("w:tblHeader"))
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    tr_pr.append(jc)


def _normalize_widths(ncols: int, widths: Sequence[int] | None) -> list[int]:
    if not widths:
        base = TEXT_WIDTH_TWIPS // ncols
        ws = [base] * ncols
        ws[-1] += TEXT_WIDTH_TWIPS - sum(ws)
        return ws
    if len(widths) != ncols:
        raise ValueError(f"列宽数量 {len(widths)} 与列数 {ncols} 不一致")
    total = sum(widths)
    if total > TEXT_WIDTH_TWIPS:           # 超出版心则等比缩放
        ws = [int(w * TEXT_WIDTH_TWIPS / total) for w in widths]
        ws[-1] += TEXT_WIDTH_TWIPS - sum(ws)
        return ws
    return list(widths)


def add_table_caption(doc, text: str):
    """表题：位于表上方，居中，独立样式。text 形如 '表1 实验环境配置'。"""
    p = doc.add_paragraph(style=STYLE_TABLE_CAPTION)
    add_text(p, text, "body")
    return p


# --------------------------------------------------------------------------
# 对外接口
# --------------------------------------------------------------------------
def add_table_type1(doc, headers: Sequence[str], rows: Sequence[Sequence[str]],
                    widths: Sequence[int] | None = None, caption: str | None = None):
    """表格一：非数据类，全框线 + 灰色表头。"""
    if caption:
        add_table_caption(doc, caption)
    ncols = len(headers)
    ws = _normalize_widths(ncols, widths)
    table = doc.add_table(rows=1 + len(rows), cols=ncols)
    line = ("single", 4)
    _set_tbl_pr(table, widths=ws, cell_margin=CELL_MARGIN_TYPE1,
                borders={s: line for s in ("top", "left", "bottom", "right",
                                           "insideH", "insideV")})
    m = (CELL_MARGIN_TYPE1,) * 4
    for ci, h in enumerate(headers):
        c = table.rows[0].cells[ci]
        _set_cell(c, ws[ci], fill=HEADER_FILL_TYPE1, margins=m)
        _fill_paragraph(c, h, header=True, bold=True)
    _row_flags(table.rows[0], header=True)
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            c = table.rows[ri].cells[ci]
            _set_cell(c, ws[ci], margins=m)
            _fill_paragraph(c, val, header=False, bold=False,
                            align=WD_ALIGN_PARAGRAPH.LEFT)
        _row_flags(table.rows[ri], header=False)
    return table


def add_table_type2(doc, headers: Sequence[str], rows: Sequence[Sequence[str]],
                    widths: Sequence[int] | None = None, caption: str | None = None):
    """表格二：数据类三线表。"""
    if caption:
        add_table_caption(doc, caption)
    ncols = len(headers)
    ws = _normalize_widths(ncols, widths)
    table = doc.add_table(rows=1 + len(rows), cols=ncols)
    none = ("none", 0)
    _set_tbl_pr(table, widths=ws, cell_margin=15,
                borders={s: none for s in ("top", "left", "bottom", "right",
                                           "insideH", "insideV")})
    for ci, h in enumerate(headers):
        c = table.rows[0].cells[ci]
        _set_cell(c, ws[ci], fill="FFFFFF", margins=(150, 120, 150, 120),
                  borders={"top": ("single", 24), "bottom": ("single", 8)})
        _fill_paragraph(c, h, header=True, bold=False)
    _row_flags(table.rows[0], header=True, min_height=TYPE2_HEADER_MIN_HEIGHT)
    last = len(rows)
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            c = table.rows[ri].cells[ci]
            bd = {"bottom": ("single", 24)} if ri == last else None
            _set_cell(c, ws[ci], margins=(60, 120, 60, 120), borders=bd)
            _fill_paragraph(c, val, header=False, bold=False,
                            align=WD_ALIGN_PARAGRAPH.CENTER)
        _row_flags(table.rows[ri], header=False)
    return table
