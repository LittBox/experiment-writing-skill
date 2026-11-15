#!/usr/bin/env python3
"""validate_docx.py - 生成后的 DOCX 强制规范检查（字体、主题字体、缩进、颜色）。

直接解析 DOCX 内部 XML（不依赖 python-docx 的对象模型），检查范围：
  word/document.xml、styles.xml、header*.xml、footer*.xml、numbering.xml、
  footnotes.xml、endnotes.xml、comments.xml。

规则（均为强制）：
  F1  任何 rFonts 出现楷体 / 微软雅黑（含别名）           -> 错误
  F2  任何 rFonts 出现主题字体属性（asciiTheme 等）        -> 错误
  F3  rFonts 写了但四个属性（ascii/hAnsi/eastAsia/cs）不全 -> 错误
  F4  字体不在白名单内                                     -> 错误
  F5  每个 run 的“有效字体”（run -> 字符样式 -> 段落样式链 -> docDefaults）
      四个槽位必须全部能解析到明确字体，且合规             -> 错误
  F6  theme1.xml 不参与检查，但样式/文档中不得引用它
  P1  Report Body 段落有效首行缩进必须为 2 字符（200）
  P2  表格内段落有效首行缩进必须为 0
  P3  标题 / 代码 / 题注样式不得有首行缩进
  C1  标题样式、表格单元格底纹、文字颜色不得为蓝色系

用法：
  python3 -I validate_docx.py report.docx [--json]
退出码：0 通过；1 有错误；2 无法读取文件。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx_styles import (  # noqa: E402
    ALLOWED_FONTS, BODY_INDENT_CHARS, FONT_ATTRS, STYLE_BODY, THEME_FONT_ATTRS,
    is_forbidden_font,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


SLOTS = ("ascii", "hAnsi", "eastAsia", "cs")
SCANNED_PARTS = re.compile(
    r"^word/(document|styles|numbering|footnotes|endnotes|comments|header\d*|footer\d*)\.xml$"
)
_ALLOWED_LOWER = {f.lower() for f in ALLOWED_FONTS}
# 标题 / 代码 / 题注类样式（styleId 为 python-docx 由样式名去空格得到）
NO_INDENT_STYLE_IDS = {
    "Heading1", "Heading2", "Heading3", "Heading4", "Title",
    "TableText", "TableHeader", "TableCaption", "Figure", "FigureCaption",
    "CodeLine", "CodeLineNumber", "CodeCaption", "Header", "Footer",
}


@dataclass
class Issue:
    code: str
    where: str
    message: str

    def __str__(self) -> str:
        return f"[{self.code}] {self.where}: {self.message}"


@dataclass
class Report:
    errors: list[Issue] = field(default_factory=list)
    stats: dict[str, int] = field(default_factory=dict)

    def add(self, code: str, where: str, message: str) -> None:
        self.errors.append(Issue(code, where, message))

    @property
    def ok(self) -> bool:
        return not self.errors


# --------------------------------------------------------------------------
# 样式解析
# --------------------------------------------------------------------------
class StyleTable:
    def __init__(self, styles_root):
        self.by_id: dict[str, etree._Element] = {}
        self.defaults_rpr = None
        self.defaults_ppr = None
        if styles_root is None:
            return
        for st in styles_root.findall("w:style", NS):
            sid = st.get(q("styleId"))
            if sid:
                self.by_id[sid] = st
        d = styles_root.find("w:docDefaults", NS)
        if d is not None:
            self.defaults_rpr = d.find("w:rPrDefault/w:rPr", NS)
            self.defaults_ppr = d.find("w:pPrDefault/w:pPr", NS)

    def chain(self, style_id: str | None) -> list[etree._Element]:
        out, seen = [], set()
        while style_id and style_id not in seen and style_id in self.by_id:
            seen.add(style_id)
            st = self.by_id[style_id]
            out.append(st)
            based = st.find("w:basedOn", NS)
            style_id = based.get(q("val")) if based is not None else None
        return out

    def default_style_id(self, kind: str) -> str | None:
        for sid, st in self.by_id.items():
            if st.get(q("type")) == kind and st.get(q("default")) in ("1", "true"):
                return sid
        return None

    def effective_fonts(self, run_rpr, para_style_id) -> dict[str, str | None]:
        """按 run -> 字符样式 -> 段落样式链 -> docDefaults 解析四个槽位。"""
        layers: list[etree._Element] = []
        if run_rpr is not None:
            layers.append(run_rpr)
            rs = run_rpr.find("w:rStyle", NS)
            if rs is not None:
                layers += [s.find("w:rPr", NS) for s in self.chain(rs.get(q("val")))]
        layers += [s.find("w:rPr", NS) for s in self.chain(para_style_id)]
        if self.defaults_rpr is not None:
            layers.append(self.defaults_rpr)
        result: dict[str, str | None] = {s: None for s in SLOTS}
        for rpr in layers:
            if rpr is None:
                continue
            rf = rpr.find("w:rFonts", NS)
            if rf is None:
                continue
            for s in SLOTS:
                if result[s] is None and rf.get(q(s)):
                    result[s] = rf.get(q(s))
        return result

    def effective_first_line(self, para_ppr, para_style_id) -> tuple[int, int]:
        """返回 (firstLineChars, firstLine twips)，未设置视为 0。"""
        layers = [para_ppr] + [s.find("w:pPr", NS) for s in self.chain(para_style_id)]
        if self.defaults_ppr is not None:
            layers.append(self.defaults_ppr)
        chars = twips = None
        for ppr in layers:
            if ppr is None:
                continue
            ind = ppr.find("w:ind", NS)
            if ind is None:
                continue
            if chars is None and ind.get(q("firstLineChars")) is not None:
                chars = int(ind.get(q("firstLineChars")))
            if twips is None and ind.get(q("firstLine")) is not None:
                twips = int(ind.get(q("firstLine")))
            if chars is not None and twips is not None:
                break
        return chars or 0, twips or 0


# --------------------------------------------------------------------------
# 检查项
# --------------------------------------------------------------------------
def _check_rfonts_element(rf, where: str, rep: Report) -> None:
    for attr in THEME_FONT_ATTRS:
        if rf.get(q(attr)) is not None:
            rep.add("F2", where, f"使用了主题字体属性 {attr}={rf.get(q(attr))}")
    present = {s: rf.get(q(s)) for s in SLOTS}
    for s, name in present.items():
        if name is None:
            continue
        if is_forbidden_font(name):
            rep.add("F1", where, f"禁用字体 {name!r}（{s}）")
        elif name.strip().lower() not in _ALLOWED_LOWER:
            rep.add("F4", where, f"字体 {name!r}（{s}）不在白名单内")
    if any(v is not None for v in present.values()) and any(v is None for v in present.values()):
        missing = [s for s, v in present.items() if v is None]
        rep.add("F3", where, f"rFonts 缺少属性：{', '.join(missing)}")


def _scan_part_fonts(name: str, root, rep: Report) -> int:
    n = 0
    for rf in root.iter(q("rFonts")):
        n += 1
        parent = rf.getparent()
        owner = parent.getparent() if parent is not None else None
        label = name
        if owner is not None and owner.tag == q("style"):
            label = f"{name}#style:{owner.get(q('styleId'))}"
        _check_rfonts_element(rf, label, rep)
    return n


def _style_id_of(p) -> str | None:
    ps = p.find("w:pPr/w:pStyle", NS)
    return ps.get(q("val")) if ps is not None else None


def _is_blue(hex_: str | None) -> bool:
    if not hex_ or hex_.lower() in ("auto", "none") or not re.fullmatch(r"[0-9A-Fa-f]{6}", hex_):
        return False
    r, g, b = (int(hex_[i:i + 2], 16) for i in (0, 2, 4))
    return b > 120 and b - r > 50 and b - g > 30


def _check_run_effective_fonts(part: str, root, styles: StyleTable, rep: Report) -> int:
    default_p = styles.default_style_id("paragraph")
    count = 0
    for p in root.iter(q("p")):
        sid = _style_id_of(p) or default_p
        for r in p.findall("w:r", NS):
            # 只检查含文字、制表符或域指令的 run
            if r.find("w:t", NS) is None and r.find("w:instrText", NS) is None \
                    and r.find("w:tab", NS) is None:
                continue
            count += 1
            eff = styles.effective_fonts(r.find("w:rPr", NS), sid)
            text = "".join(t.text or "" for t in r.findall("w:t", NS))[:20]
            where = f"{part} 段落(样式={sid}) run {text!r}"
            for s in SLOTS:
                if eff[s] is None:
                    rep.add("F5", where, f"槽位 {s} 没有明确字体（将由 Word 自动替换）")
                elif is_forbidden_font(eff[s]):
                    rep.add("F1", where, f"有效字体为禁用字体 {eff[s]!r}（{s}）")
                elif eff[s].strip().lower() not in _ALLOWED_LOWER:
                    rep.add("F4", where, f"有效字体 {eff[s]!r}（{s}）不在白名单内")
    return count


def _check_indents(part: str, root, styles: StyleTable, rep: Report) -> None:
    default_p = styles.default_style_id("paragraph")
    for p in root.iter(q("p")):
        sid = _style_id_of(p) or default_p
        ppr = p.find("w:pPr", NS)
        chars, twips = styles.effective_first_line(ppr, sid)
        in_table = any(a.tag == q("tc") for a in p.iterancestors())
        text = "".join(t.text or "" for t in p.iter(q("t")))[:20]
        where = f"{part} 段落(样式={sid}) {text!r}"
        if in_table and (chars or twips):
            rep.add("P2", where, f"表格内段落带首行缩进 (chars={chars}, twips={twips})")
        elif sid == STYLE_BODY.replace(" ", "") and not in_table:
            if chars != BODY_INDENT_CHARS * 100:
                rep.add("P1", where, f"正文首行缩进应为 {BODY_INDENT_CHARS} 字符，实际 {chars / 100:g}")
        elif sid in NO_INDENT_STYLE_IDS and (chars or twips):
            rep.add("P3", where, f"该样式不应有首行缩进 (chars={chars}, twips={twips})")


def _check_colors(part: str, root, styles: StyleTable, rep: Report) -> None:
    for el in root.iter(q("color")):
        v = el.get(q("val"))
        if _is_blue(v) or el.get(q("themeColor")) is not None:
            owner = el.getparent().getparent()
            label = owner.get(q("styleId")) if owner is not None and owner.tag == q("style") else part
            rep.add("C1", f"{part}#{label}", f"文字颜色 {v}/主题色 {el.get(q('themeColor'))} 不允许")
    for el in root.iter(q("shd")):
        for attr in ("fill", "color"):
            if _is_blue(el.get(q(attr))) or el.get(q("themeFill")) is not None and \
                    el.get(q("themeFill")) not in ("background1", "light1"):
                rep.add("C1", part, f"底纹 {attr}={el.get(q(attr))} 为蓝色/主题色")
                break


# --------------------------------------------------------------------------
# 入口
# --------------------------------------------------------------------------
def validate(path: str | Path) -> Report:
    rep = Report()
    try:
        zf = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile) as e:
        raise RuntimeError(f"无法读取 {path}: {e}") from e
    with zf:
        names = zf.namelist()
        parsed = {}
        for n in names:
            if SCANNED_PARTS.match(n):
                parsed[n] = etree.fromstring(zf.read(n))
        styles = StyleTable(parsed.get("word/styles.xml"))
        rfonts_total = run_total = 0
        for n, root in parsed.items():
            rfonts_total += _scan_part_fonts(n, root, rep)
            _check_colors(n, root, styles, rep)
            if n != "word/styles.xml" and n != "word/numbering.xml":
                run_total += _check_run_effective_fonts(n, root, styles, rep)
                _check_indents(n, root, styles, rep)
        # 编号级别里的字体也走 rFonts 扫描；此处额外校验 numbering 的 lvl 字体槽位
        if "word/document.xml" not in parsed:
            rep.add("F5", "word/document.xml", "缺少 document.xml")
        rep.stats = {"rFonts 元素": rfonts_total, "检查的 run": run_total,
                     "扫描的部件": len(parsed)}
    return rep


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="校验 DOCX 是否符合强制排版规范")
    ap.add_argument("docx")
    ap.add_argument("--json", action="store_true", help="以 JSON 输出报告")
    args = ap.parse_args(argv)
    try:
        rep = validate(args.docx)
    except RuntimeError as e:
        print(e, file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps({"ok": rep.ok, "stats": rep.stats,
                          "errors": [i.__dict__ for i in rep.errors]},
                         ensure_ascii=False, indent=2))
    else:
        print("统计：", rep.stats)
        if rep.ok:
            print("通过：未发现违规字体、主题字体、缩进或颜色问题。")
        else:
            print(f"发现 {len(rep.errors)} 个错误：")
            for i in rep.errors:
                print("  ", i)
    return 0 if rep.ok else 1


if __name__ == "__main__":
    sys.exit(main())
