# -*- coding: utf-8 -*-
"""
filing_docx.py — 按投行排版合同（filing_contract.py）生成 / 修补 / 校验 .docx

用法：
  python filing_docx.py apply  --in 源.docx --out 目标.docx [--profile 保荐书|通用] [--keep-page]
  python filing_docx.py check  目标.docx [--json]
  python filing_docx.py new   --out 新文件.docx --profile 保荐书

apply：注入全套样式、多级编号、页面/页眉页脚、目录域、表格美颜；不动正文文字。
check：逐条比对合同，输出 PASS/FAIL 与偏差值；退出码 0=全部通过，1=有偏差。
"""
import argparse, copy, io, json, os, re, shutil, sys, zipfile
from lxml import etree
from docx import Document
from docx.oxml.ns import qn, nsmap
from docx.shared import Pt, Cm

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import filing_contract as C

W = nsmap["w"]

# 校验模式：strict=False 为"审计模式"——把历史文件的宽容写法（手写序号、收窄表宽、
# 缺表头重复等）报为 WARN，只把硬性矛盾报 FAIL；strict=True 为"交付模式"——一切按合同
# 判 FAIL。本技能自己产出的文件在两种模式下都应 0 FAIL。


def make(tag, **kw):
    qtag = tag if ":" in tag else "w:" + tag
    e = etree.Element(qn(qtag), nsmap={"w": W})
    for k, v in kw.items():
        e.set(qn(k if ":" in k else "w:" + k), str(v))
    return e


def sub(parent, tag, **kw):
    e = make(tag, **kw)
    parent.append(e)
    return e

# --------------------------------------------------------------- styles.xml
def _style_element(spec, auto_numbering=False):
    st = make("style", type=spec["type"], styleId=spec["id"])
    if spec.get("default"):
        st.set(qn("w:default"), "1")
    if spec.get("custom") and spec["type"] != "character":
        st.set(qn("w:customStyle"), "1")
    nm = sub(st, "name"); nm.set(qn("w:val"), spec["name"])
    if spec.get("based_on"):
        b = sub(st, "basedOn"); b.set(qn("w:val"), spec["based_on"])
    if spec.get("next"):
        n = sub(st, "next"); n.set(qn("w:val"), spec["next"])
    if spec.get("link"):
        l = sub(st, "link"); l.set(qn("w:val"), spec["link"])
    sub(st, "qFormat")

    ppr_needed = any(k in spec for k in (
        "jc", "first_line", "first_line_chars", "left", "left_chars", "hanging",
        "before", "after", "before_lines", "after_lines", "line", "exact",
        "keep_next", "keep_lines", "page_break", "outline", "widow", "numbering",
        "tab_pos")) or spec.get("auto_update")
    if spec["type"] == "paragraph" and ppr_needed:
        ppr = sub(st, "pPr")
        if spec.get("keep_next"):
            sub(ppr, "keepNext")
        if spec.get("keep_lines"):
            sub(ppr, "keepLines")
        if spec.get("page_break"):
            sub(ppr, "pageBreakBefore")
        if spec.get("widow") is False:
            sub(ppr, "widowControl", val=0)
        # 【硬性规则】默认不写自动编号：除非使用者明确要求，
        # 标题样式不挂 numPr，序号一律手写在标题文字里。
        if "numbering" in spec and auto_numbering:
            npr = sub(ppr, "numPr")
            if spec["numbering"]:
                sub(npr, "ilvl", val=spec["numbering"])
            sub(npr, "numId", val=1)
        if spec.get("tab_pos"):
            t = sub(ppr, "tabs")
            tb = sub(t, "tab", val="right", leader="dot", pos=spec["tab_pos"])
        # spacing 只在实际声明了段前/段后/行距时才创建。
        # 【实测修复】原实现无条件写 `line=360`（1.5 倍），把**未声明行距**的样式
        # （Normal / 006备注 / 009单位 / 011五级标题 / 表格前单位 / 表格后说明）
        # 也强行改成 1.5 倍行距。实证文件 3-2 上市保荐书.docx 里这些样式是
        # **单倍行距**（无 w:spacing），改成 1.5 倍会让全篇凭空多出 1 页
        # （逐步隔离实测：交付件 26 页 → 仅把 006 恢复单倍行距即回到 25 页，
        # 与原件页数一致）。合同本就故意不声明这些样式的 line，故按声明写。
        spacing_keys = ("before_lines", "before", "after_lines", "after", "line")
        if any(spec.get(k) is not None for k in spacing_keys) or spec.get("exact"):
            sp = sub(ppr, "spacing")
            if spec.get("before_lines") is not None:
                sp.set(qn("w:beforeLines"), str(spec["before_lines"]))
            if spec.get("before") is not None:
                sp.set(qn("w:before"), str(spec["before"]))
            if spec.get("after_lines") is not None:
                sp.set(qn("w:afterLines"), str(spec["after_lines"]))
            if spec.get("after") is not None:
                sp.set(qn("w:after"), str(spec["after"]))
            if spec.get("line") is not None:
                sp.set(qn("w:line"), str(spec["line"]))
                # 公告体裁用固定行距（exact），申报文件用倍数（auto）
                sp.set(qn("w:lineRule"), "exact" if spec.get("exact") else "auto")
        if any(k in spec for k in ("left", "left_chars", "first_line",
                                   "first_line_chars", "hanging")):
            ind = sub(ppr, "ind")
            if spec.get("left_chars") is not None:
                ind.set(qn("w:leftChars"), str(spec["left_chars"]))
            if spec.get("left") is not None:
                ind.set(qn("w:left"), str(spec["left"]))
            if spec.get("first_line_chars") is not None:
                ind.set(qn("w:firstLineChars"), str(spec["first_line_chars"]))
            if spec.get("first_line") is not None:
                ind.set(qn("w:firstLine"), str(spec["first_line"]))
            if spec.get("hanging") is not None:
                ind.set(qn("w:hanging"), str(spec["hanging"]))
        if spec.get("jc"):
            sub(ppr, "jc", val=spec["jc"])
        if spec.get("outline") is not None:
            sub(ppr, "outlineLvl", val=spec["outline"])

    needs_rpr = any(k in spec for k in ("font", "sz", "szCs", "bold", "kern"))
    if needs_rpr:
        rpr = sub(st, "rPr")
        if spec.get("font"):
            f = sub(rpr, "rFonts")
            f.set(qn("w:ascii"), spec["font"]["ascii"])
            f.set(qn("w:eastAsia"), spec["font"]["east_asia"])
            f.set(qn("w:hAnsi"), spec["font"]["hAnsi"])
            f.set(qn("w:cs"), "Times New Roman")
        if spec.get("bold"):
            sub(rpr, "b"); sub(rpr, "bCs")
        if spec.get("kern"):
            sub(rpr, "kern", val=spec["kern"])
        if spec.get("sz"):
            sub(rpr, "sz", val=spec["sz"])
        if spec.get("szCs"):
            sub(rpr, "szCs", val=spec["szCs"])
    return st


def write_styles(doc, auto_numbering=None):
    """写入合同全部段落/字符/表格样式（已存在则整体替换）。

    auto_numbering: True 才给标题样式挂自动多级编号。
        默认取 C.AUTO_NUMBERING_DEFAULT（False）—— 序号手写在标题文字里。
    """
    if auto_numbering is None:
        auto_numbering = C.AUTO_NUMBERING_DEFAULT
    styles = doc.styles.element
    wanted = {s["id"]: s for s in C.all_styles()}
    for st in list(styles.findall(qn("w:style"))):
        sid = st.get(qn("w:styleId"))
        if sid in wanted:
            styles.remove(st)
    for spec in C.all_styles():
        # 手写编号口径下必须剥掉 spec 里的 numbering 字段，
        # 否则标题样式会挂上 numPr（编号既重复又与自动编号耦合）。
        use = spec
        if not auto_numbering and "numbering" in spec:
            use = {k: v for k, v in spec.items() if k != "numbering"}
        el = _style_element(use, auto_numbering=auto_numbering)
        styles.append(el)
    return [s["name"] for s in C.all_styles()]


def write_family_styles(doc, family):
    """只写入指定体裁的样式（保留文档已有的其他族）。
    family: announce / guide —— 找不到通用样式骨架时先补基础样式，
    否则自定义样式会因缺 Normal 而无法解析。"""
    ensure_base_styles(doc)
    specs = C.ANNOUNCE_STYLES if family == "announce" else C.GUIDE_STYLES
    styles = doc.styles.element
    wanted = {s["id"]: s for s in specs}
    for st in list(styles.findall(qn("w:style"))):
        if st.get(qn("w:styleId")) in wanted:
            styles.remove(st)
    for spec in specs:
        styles.append(_style_element(spec))
    return [s["name"] for s in specs]


def write_announce_styles(doc):
    """只写入公告体裁样式。"""
    return write_family_styles(doc, "announce")


def write_guide_styles(doc):
    """只写入公文体裁样式。"""
    return write_family_styles(doc, "guide")


# 公告体裁必需的基础样式（缺 Normal 时 Word 会拒绝解析自定义样式）
BASE_STYLES = [
    {"id": "Normal", "name": "Normal", "type": "paragraph", "default": True,
     "font": {"east_asia": "仿宋_GB2312", "ascii": "仿宋_GB2312", "hAnsi": "仿宋GB-2312"},
     "sz": 30, "szCs": 30, "jc": "left", "line": 560, "exact": True},
    {"id": "TableNormal", "name": "Normal Table", "type": "table", "default": True},
]


def ensure_base_styles(doc):
    """补上列表/表格用的基础样式，使自定义样式能被 Word 正确解析。"""
    styles = doc.styles.element
    have = {s.get(qn("w:styleId")) for s in styles.findall(qn("w:style"))}
    by_name = set()
    for s in styles.findall(qn("w:style")):
        nm = s.find(qn("w:name"))
        if nm is not None:
            by_name.add(nm.get(qn("w:val")))
    added = []
    if "Normal" not in by_name and "a" not in have:
        styles.insert(0, _style_element(BASE_STYLES[0]))
        added.append("Normal")
    if "Normal Table" not in by_name and "a1" not in have:
        styles.append(_style_element(BASE_STYLES[1]))
        added.append("Normal Table")
    return added


# --------------------------------------------------------------- numbering.xml
def drop_numbering(doc):
    """移除 numbering.xml 部件及其关系，并清掉样式表里残留的 numPr。

    python-docx 的 Document() 默认自带一个 numbering 部件，并让内置列表样式
    （List Bullet / List Number 等）带着 numPr 指向它。手写编号口径下这些
    既无用又会污染"是否存在自动编号"的判定，故一并摘除。
    """
    part = doc.part
    styles = doc.styles.element
    removed_style_numpr = 0
    for st in styles.findall(qn("w:style")):
        ppr = st.find(qn("w:pPr"))
        if ppr is None:
            continue
        npr = ppr.find(qn("w:numPr"))
        if npr is not None:
            ppr.remove(npr)
            removed_style_numpr += 1
    try:
        npart = part.numbering_part
    except Exception:
        return {"style_numpr_removed": removed_style_numpr, "part_removed": False}
    for rid in list(part.rels):
        try:
            if part.rels[rid].target_part is npart:
                del part.rels[rid]
        except Exception:
            pass
    try:
        part.package.parts.remove(npart)
    except Exception:
        pass
    return {"style_numpr_removed": removed_style_numpr, "part_removed": True}


def write_numbering(doc):
    """注入插件同款多级编号：第%1节 / %2、 / （%3） / %4、 / （%5） / %6） / ①（原地替换）。"""
    part = doc.part.numbering_part
    root = part._element
    for child in list(root):
        root.remove(child)
    new = _numbering_root()
    for child in list(new):
        root.append(child)
    return True


def _numbering_root():
    root = etree.Element(qn("w:numbering"), nsmap={"w": W})
    abst = make("abstractNum", abstractNumId=0)
    sub(abst, "nsid", val="FFFFFF7C")
    sub(abst, "multiLevelType", val="multilevel")
    sub(abst, "tmpl", val="FFFFFF7C")
    sizes = [(32, 44), (28, 32), (24, 28), (24, 28), (24, 28), (24, 24), (24, 24)]
    for lv in C.NUMBERING:
        l = make("lvl", ilvl=lv["ilvl"])
        sub(l, "start", val=1)
        sub(l, "numFmt", val=lv["fmt"])
        sub(l, "pStyle", val=lv["style"])
        sub(l, "lvlText", val=lv["text"])
        sub(l, "lvlJc", val="left")
        ppr = sub(l, "pPr")
        sub(ppr, "ind", left=0, firstLine=0)
        rpr = sub(l, "rPr")
        f = sub(rpr, "rFonts")
        f.set(qn("w:ascii"), "Times New Roman")
        f.set(qn("w:eastAsia"), "黑体" if lv["ilvl"] < 3 else "宋体")
        f.set(qn("w:hAnsi"), "Times New Roman")
        sz, szcs = sizes[lv["ilvl"]]
        sub(rpr, "sz", val=sz)
        sub(rpr, "szCs", val=szcs)
        abst.append(l)
    root.append(abst)
    n = make("num", numId=1)
    sub(n, "abstractNumId", val=0)
    root.append(n)
    return root


# --------------------------------------------------------------- section
def setup_section(sec, cover=False, body=False):
    from docx.enum.section import WD_ORIENT
    p = C.PAGE
    sec.page_width = Pt(p["page_w"] / 20)
    sec.page_height = Pt(p["page_h"] / 20)
    sec.top_margin = Pt(p["margin_top"] / 20)
    sec.bottom_margin = Pt(p["margin_bottom"] / 20)
    sec.left_margin = Pt(p["margin_left"] / 20)
    sec.right_margin = Pt(p["margin_right"] / 20)
    sec.header_distance = Pt(p["header_dist"] / 20)
    sec.footer_distance = Pt(p["footer_dist"] / 20)
    sectPr = sec._sectPr
    for tag in ("pgSz", "pgMar"):
        el = sectPr.find(qn("w:" + tag))
        if el is not None:
            if tag == "pgSz":
                el.set(qn("w:w"), str(p["page_w"])); el.set(qn("w:h"), str(p["page_h"]))
                if el.get(qn("w:orient")):
                    del el.attrib[qn("w:orient")]
            else:
                el.set(qn("w:top"), str(p["margin_top"])); el.set(qn("w:bottom"), str(p["margin_bottom"]))
                el.set(qn("w:left"), str(p["margin_left"])); el.set(qn("w:right"), str(p["margin_right"]))
                el.set(qn("w:header"), str(p["header_dist"])); el.set(qn("w:footer"), str(p["footer_dist"]))
                el.set(qn("w:gutter"), str(p["gutter"]))
    if cover:
        pn = sectPr.find(qn("w:pgNumType"))
        if pn is None:
            pn = sub(sectPr, "pgNumType")
        pn.set(qn("w:start"), str(p["cover_page_start"]))
        if sectPr.find(qn("w:titlePg")) is None:
            tpg = make("titlePg")
            pn.addnext(tpg) if pn is not None else sectPr.append(tpg)
    if body:
        pn = sectPr.find(qn("w:pgNumType"))
        if pn is None:
            pn = sub(sectPr, "pgNumType")
        pn.set(qn("w:start"), str(p["body_page_start"]))


def get_style(doc, key):
    """按 styleId 或中文样式名取样式对象（避开 python-docx 的内置样式英文名映射）。"""
    for s in doc.styles:
        try:
            if s.style_id == key or s.name == key:
                return s
        except Exception:
            continue
    raise KeyError("style not found: %s" % key)


def set_header(doc, section, text, east_asia="宋体"):
    hdr = section.header
    hdr.is_linked_to_previous = False
    for p in list(hdr.paragraphs[1:]):
        p._p.getparent().remove(p._p)
    p = hdr.paragraphs[0] if hdr.paragraphs else hdr.add_paragraph()
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    p.style = get_style(doc, "Header")
    run = p.add_run(text)
    f = C.FONT
    run.font.name = f["hAnsi"]
    run._element.rPr.rFonts.set(qn("w:eastAsia"), f["east_asia"])
    run.font.size = Pt(C.HEADER_FOOTER["header_sz"] / 2)
    pPr = p._p.get_or_add_pPr()
    old = pPr.find(qn("w:pBdr"))
    if old is not None:
        pPr.remove(old)
    b = make("pBdr")
    bb = sub(b, "bottom", val="single", sz=4, space=1, color="auto")
    pPr.append(b)
    return hdr


def _unwrap_sdt(root):
    """把 w:sdt 包裹的段落释放为直接子节点。

    页眉/页脚常把段落包在结构化文档标记（w:sdt 内嵌在 w:sdtContent）里，
    而 python-docx 的 header.paragraphs / footer.paragraphs 只认直接子节点。
    不解开就会"新增一段"而旧段仍在，最终渲染出两个页码。"""
    n = 0
    for sdt in list(root.iter(qn("w:sdt"))):
        content = sdt.find(qn("w:sdtContent"))
        if content is None:
            continue
        parent = sdt.getparent()
        idx = list(parent).index(sdt)
        for child in list(content):
            content.remove(child)
            parent.insert(idx, child)
            idx += 1
        parent.remove(sdt)
        n += 1
    return n


def clear_header(doc, section):
    """清空页眉（公文体裁不设页眉，页眉区留白）。"""
    hdr = section.header
    hdr.is_linked_to_previous = False
    _unwrap_sdt(hdr._element)
    for p in list(hdr.paragraphs[1:]):
        p._p.getparent().remove(p._p)
    if hdr.paragraphs:
        p = hdr.paragraphs[0]
        for r in list(p.runs):
            r._r.getparent().remove(r._r)
        pPr = p._p.get_or_add_pPr()
        old = pPr.find(qn("w:pBdr"))
        if old is not None:
            pPr.remove(old)
    return hdr


def set_footer(doc, section, prefix="", align="center"):
    """写默认页脚：可选前缀 + PAGE 域。"""
    return write_page_footer(doc, section.footer, prefix, align)


def write_page_footer(doc, ftr, prefix="", align="center"):
    """把任一类页脚（default / even / first）规范为：居中 + 单个 PAGE 域。

    先解开 w:sdt 包裹——否则 footer.paragraphs 看不到被包裹的旧段落，
    会写成两段（渲染成两个页码）。"""
    ftr.is_linked_to_previous = False
    _unwrap_sdt(ftr._element)
    for p in list(ftr.paragraphs[1:]):
        p._p.getparent().remove(p._p)
    p = ftr.paragraphs[0] if ftr.paragraphs else ftr.add_paragraph()
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    # 清掉继承来的 framePr（转换产物常带 frame 定位，会把页码挤到奇怪位置）
    pPr = p._p.get_or_add_pPr()
    fr = pPr.find(qn("w:framePr"))
    if fr is not None:
        pPr.remove(fr)
    try:
        p.style = get_style(doc, "Footer")
    except KeyError:
        pass
    p.paragraph_format.alignment = {"center": 1, "left": 0, "right": 2}[align]
    # 清掉段落里已有的域与残留 run，避免叠加出两个页码
    for r in list(p._p.findall(qn("w:r"))):
        p._p.remove(r)
    for hl in list(p._p.findall(qn("w:hyperlink"))):
        p._p.remove(hl)
    if prefix:
        p.add_run(prefix)
    _add_field(p, C.HEADER_FOOTER["footer_page_field"])
    for r in p.runs:
        r.font.size = Pt(C.HEADER_FOOTER["footer_sz"] / 2)
        r.font.name = "Times New Roman"
        rpr = r._element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = etree.SubElement(rpr, qn("w:rFonts"))
        rf.set(qn("w:eastAsia"), "宋体")
        rf.set(qn("w:ascii"), "Times New Roman")
        rf.set(qn("w:hAnsi"), "Times New Roman")
        if "hint" in rf.attrib:
            del rf.attrib[qn("w:hint")]
    return ftr


def _add_field(p, code):
    """在段落末尾插入域（begin / instrText / separate / t / end）。"""
    for kind in ("begin",):
        r = p.add_run()._r
        e = make("fldChar", fldCharType=kind)
        r.append(e)
    r = p.add_run()._r
    it = make("instrText")
    it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    it.text = code
    r.append(it)
    r = p.add_run()._r
    r.append(make("fldChar", fldCharType="separate"))
    r = p.add_run()._r
    t = make("t"); t.text = "1"
    r.append(t)
    r = p.add_run()._r
    r.append(make("fldChar", fldCharType="end"))
    return p


def add_toc(doc, title="目  录"):
    """插入目录标题 + TOC 域（TOC \\o "1-2" \\h \\z \\u）。"""
    p = doc.add_paragraph(title, style=get_style(doc, "TOCHeading"))
    tp = doc.add_paragraph(style=get_style(doc, "TOC1"))
    _add_field(tp, C.HEADER_FOOTER["toc_field"])
    return p, tp


# --------------------------------------------------------------- tables
def format_table(table, pattern=None, font_sz=None, row_height=None,
                 cell_font=None, keep_cell_align=False, keep_row_height=False,
                 allow_row_split=None, keep_cell_margin=False):
    """表格美颜：边框、字体字号、行高、表头重复、对齐、数字靠右、合计行加粗。

    pattern: threeline / boxed / blueprint / announcement
             （默认取 C.TABLE['default_pattern']）
    font_sz / row_height: 覆盖该 pattern 的默认值
    cell_font: 覆盖表内中文字体（公告/公文体裁用 仿宋_GB2312；默认 宋体）
    keep_cell_align: True 时保留单元格原有的水平对齐，只改表头。
        —— 用于"索引表/目录表"（序号列惯例居中），财务表才用数字右对齐。
    keep_row_height: True 时**只在行本身没有行高时**才写入标准值。
        —— 用于超长表格（数百行）：原作者已按紧凑度设过行高。
    allow_row_split: 是否允许行内跨页断开（None = 按表长自动判断）。
        —— 标准里"不允许跨页断行 (cantSplit)"只适合短表；
        对数百行的长表，cantSplit 会迫使每行整行落在同一页，页尾大量空白。
    keep_cell_margin: True 时保留原有单元格左右边距。
        —— 紧凑多列表（如 8 列的索引表）原作者常把边距设为 0 以容纳更多字；
        按标准改成 0.19cm 会让每列有效宽度减少、换行增多，
        407 行的表因此凭空多出近十页（实测 48→57 页）。
    """
    key = pattern or C.TABLE["default_pattern"]
    pat = C.TABLE_PATTERNS[key]
    fs = font_sz if font_sz is not None else pat["font_sz"]
    rh = row_height if row_height is not None else pat["row_height"]
    cf = cell_font or C.TABLE["cell_font"]
    n_rows = len(table.rows)
    if allow_row_split is None:
        allow_row_split = n_rows >= 40          # 长表允许行内断页

    tbl = table._tbl
    tblPr = tbl.tblPr
    # 宽度 100% + 居中
    for tag in ("tblW", "jc"):
        el = tblPr.find(qn("w:" + tag))
        if el is not None:
            tblPr.remove(el)
    w = make("tblW", w=C.TABLE["width_pct"], type="pct"); tblPr.append(w)
    jc = make("jc", val=C.TABLE["align"]); tblPr.append(jc)
    # 边框
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    b = make("tblBorders")
    for name, spec in pat["borders"].items():
        e = sub(b, name, val=spec["val"])
        if spec.get("sz"):
            e.set(qn("w:sz"), str(spec["sz"]))
            e.set(qn("w:space"), "0")
            e.set(qn("w:color"), spec.get("color", "auto"))
    tblPr.append(b)
    # 单元格边距
    old = tblPr.find(qn("w:tblCellMar"))
    if old is not None and not keep_cell_margin:
        tblPr.remove(old)
    if old is None or not keep_cell_margin:
        cm = make("tblCellMar")
        sub(cm, "top", w=C.TABLE["pad_top"], type="dxa")
        sub(cm, "left", w=C.TABLE["pad_lr"], type="dxa")
        sub(cm, "bottom", w=C.TABLE["pad_bottom"], type="dxa")
        sub(cm, "right", w=C.TABLE["pad_lr"], type="dxa")
        tblPr.append(cm)
    # 行高 / 表头
    for ri, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        prev_h = trPr.find(qn("w:trHeight"))
        prev_val = prev_h.get(qn("w:val")) if prev_h is not None else None
        for tag in ("trHeight", "cantSplit", "tblHeader"):
            el = trPr.find(qn("w:" + tag))
            if el is not None:
                trPr.remove(el)
        # 超长表保留原作者的行高（强制统一会让 400 行表长出十余页）
        use_h = rh
        if keep_row_height and prev_val not in (None, "0"):
            use_h = prev_val
        h = sub(trPr, "trHeight", val=use_h)
        h.set(qn("w:hRule"), "atLeast")
        # cantSplit（行不可跨页断开）只给短表；长表加它会迫使每行整行落同页，
        # 页尾留大量空白（实测 407 行表使文档 48→58 页）
        if not allow_row_split:
            sub(trPr, "cantSplit")
        if ri == 0 and C.TABLE["heading_format"]:
            sub(trPr, "tblHeader")
            row.repeat_header = True
    # 字体 / 字号 / 垂直居中
    for ri, row in enumerate(table.rows):
        for cell in row.cells:
            # 直接遍历 XML：paragraphs[].runs 不含超链接、域、sdt 里嵌套的 run，
            # 只靠 python-docx 的 runs 会漏掉这些（表现为个别单元格字号不统一）
            for r in cell._tc.iter(qn("w:r")):
                rpr = r.find(qn("w:rPr"))
                if rpr is None:
                    rpr = etree.Element(qn("w:rPr"))
                    r.insert(0, rpr)
                rf = rpr.find(qn("w:rFonts"))
                if rf is None:
                    rf = etree.Element(qn("w:rFonts"))
                    rpr.insert(0, rf)
                rf.set(qn("w:ascii"), cf["ascii"])
                rf.set(qn("w:hAnsi"), cf.get("hAnsi", cf["ascii"]))
                rf.set(qn("w:eastAsia"), cf["east_asia"])
                rf.set(qn("w:cs"), cf["ascii"])
                for k in ("hint", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
                    if k in rf.attrib:
                        del rf.attrib[qn("w:" + k)]
                for tag in ("w:sz", "w:szCs"):
                    e = rpr.find(qn(tag))
                    if e is None:
                        e = etree.SubElement(rpr, qn(tag))
                    e.set(qn("w:val"), str(fs))
                b = rpr.find(qn("w:b"))
                want_bold = (ri == 0) if C.TABLE["header_bold"] else False
                if want_bold:
                    if b is None:
                        b = etree.SubElement(rpr, qn("w:b"))
                    b.set(qn("w:val"), "1")
                elif b is not None:
                    rpr.remove(b)
            tcPr = cell._tc.get_or_add_tcPr()
            va = tcPr.find(qn("w:vAlign"))
            if va is None:
                va = sub(tcPr, "vAlign")
            va.set(qn("w:val"), C.TABLE["v_align"])
            shd = tcPr.find(qn("w:shd"))
            if shd is not None:
                tcPr.remove(shd)
            if ri == 0 and pat["header_shading"]:
                sub(tcPr, "shd", val="clear", color="auto", fill=pat["header_shading"])
    # 对齐
    for ri, row in enumerate(table.rows):
        for cell in row.cells:
            if ri == 0:
                align = C.TABLE["header_jc"]
            elif keep_cell_align:
                # 保留原对齐（索引表/目录表：序号列与类别列惯用居中）
                continue
            elif C.is_numeric_cell(cell.text):
                align = C.TABLE["cell_jc_numeric"]
            else:
                align = C.TABLE["cell_jc_text"]
            for p in cell.paragraphs:
                p.paragraph_format.alignment = {"center": 1, "right": 2, "left": 0,
                                                "both": 3}[align]
                # 单元格内段落不得残留首行缩进 / 段前段后
                p.paragraph_format.first_line_indent = 0
                p.paragraph_format.space_before = 0
                p.paragraph_format.space_after = 0
    if keep_cell_align:
        # 仍要清掉表内段落的缩进与段距（否则行高会被撑开）
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.first_line_indent = 0
                    p.paragraph_format.space_before = 0
                    p.paragraph_format.space_after = 0
    # 表头行与下行同页
    if C.TABLE["header_keep_next"] and table.rows:
        for p in table.rows[0].cells[0].paragraphs:
            p.paragraph_format.keep_with_next = True
    # 合计行（整行加粗 + 首格居中）
    last = table.rows[-1]
    if last.cells and last.cells[0].text.strip().startswith(C.TABLE["total_row_prefix"]):
        for cell in last.cells:
            for r in cell._tc.iter(qn("w:r")):
                rpr = r.find(qn("w:rPr"))
                if rpr is None:
                    rpr = etree.Element(qn("w:rPr"))
                    r.insert(0, rpr)
                b = rpr.find(qn("w:b"))
                if b is None:
                    b = etree.SubElement(rpr, qn("w:b"))
                b.set(qn("w:val"), "1")
            for p in cell.paragraphs:
                p.paragraph_format.alignment = 1
    return True


# --------------------------------------------------------------- new doc
def new_document(profile="通用", auto_numbering=None):
    """新建合同文档。

    auto_numbering: True 才注入自动多级编号。默认 False（序号手写），
        即 numbering.xml 不写入、标题样式也不挂 numPr。
    """
    if auto_numbering is None:
        auto_numbering = C.AUTO_NUMBERING_DEFAULT
    doc = Document()
    write_styles(doc, auto_numbering=auto_numbering)
    if auto_numbering:
        write_numbering(doc)
    else:
        # 手写编号口径：摘掉 python-docx 默认自带的空 numbering 部件
        drop_numbering(doc)
    setup_section(doc.sections[0], cover=False, body=True)
    return doc


# --------------------------------------------------------------- check
def check(path, as_json=False, strict=False, profile="全部", auto_numbering=False,
          pattern=None):
    """逐条比对合同；输出偏差清单。返回 (ok, findings)。

    auto_numbering: 使用者是否明确要求使用自动多级编号。
        默认 False = 手写序号口径，此时文档若挂了 numPr/numbering.xml 判 FAIL。
    pattern: 使用者明确指定的表格线型（threeline/boxed/blueprint/announcement）。
        默认 None = 四边线表（表格美颜②）；公告/公文强制 announcement。

    profile 决定「哪些样式族必须存在」：
      申报文件 / 招股说明书 / 回复报告 / 新三板 / 公告 / 公文 → 只校验该口径用到的族
      全部（默认）→ 各族都要求

    strict=False（默认）容忍四类合法差异，降级为 WARN：
      1) 边距 1797 与 1800 的 ≤4 twips 抖动（Word 排版回写造成）
      2) 无表头重复且宽度为 dxa 的落款/签字表（非数据表）
      3) 已写 firstLineChars 时 firstLine 的 twips 取值差异（渲染以字符单位为准）
      4) 未显式写入但可由 basedOn/docDefaults 继承得到的字体与字号
    """
    z = zipfile.ZipFile(path)
    names = set(z.namelist())
    findings = []

    def add(level, area, rule, expected, actual):
        findings.append({"level": level, "area": area, "rule": rule,
                         "expected": expected, "actual": actual})

    def near(a, b, tol=4):
        try:
            return abs(int(a) - int(b)) <= tol
        except (TypeError, ValueError):
            return False

    # 按 profile 过滤必查样式
    FAMILY = {
        "申报文件": ["001", "002", "003", "003Char", "004", "005", "005Char", "006", "009", "011",
                     "TableUnit", "TableNote", "ReportTable", "TOCHeading", "TOC1", "TOC2",
                     "Header", "Footer"],
        "招股说明书": ["Heading1", "Heading1Char", "Heading2", "Heading2Char", "Heading3",
                       "Heading3Char", "Heading4", "Heading5", "Heading6", "Heading7",
                       "BodyIndent", "TableUnit", "TableNote", "ReportTable",
                       "TOCHeading", "TOC1", "TOC2", "Header", "Footer"],
        "回复报告": ["Heading1", "Heading1Char", "Heading2", "Heading2Char", "Heading3",
                     "Heading3Char", "Heading4", "Heading5", "BodyIndent",
                     "TableUnit", "TableNote", "ReportTable", "TOCHeading", "TOC1", "TOC2",
                     "Header", "Footer"],
        "新三板": ["001", "002", "003", "003Char", "004", "005", "005Char", "006", "009",
                   "TableUnit", "TableNote", "ReportTable", "TOCHeading", "TOC1", "TOC2",
                   "Header", "Footer"],
        # 交易所临时公告体裁
        "公告": ["AnnTitle", "AnnSubject", "AnnH1", "AnnH2", "AnnH3", "AnnBody",
                 "AnnList", "AnnNote", "AnnSection", "AnnSign"],
        # 交易所自律监管指南 / 业务办理指南（公文体例）
        "公文": ["GuideTitle", "GuideH1", "GuideH2", "GuideBody", "GuideAnnex",
                 "GuideCaption", "GuideCell", "GuideNote"],
    }
    want_ids = None
    if profile in FAMILY:
        want_ids = set(FAMILY[profile]) | {"Normal"}
        # 公告/公文体裁的表格用"直接边框"实现（0.5pt 细线格），
        # 不依赖 普通表格 / 报告表格 样式；要求它存在是错的。
        if profile not in ("公告", "公文"):
            want_ids.add("NormalTable")
    # 必须从 all_styles() 取，否则公告/公文体裁的样式族根本不会被校验
    specs = [s for s in C.all_styles()
             if want_ids is None or s["id"] in want_ids]
    # Normal 的字体/字号由各体裁的 docDefaults 决定，不按申报文件规格要求
    normal_spec = None
    for s in specs:
        if s["id"] == "Normal":
            normal_spec = s
    if normal_spec is not None and profile in ("公告", "公文"):
        prof = C.PROFILES[profile]
        nf = prof.get("body_font") or normal_spec.get("font")
        normal_spec = dict(normal_spec)
        normal_spec["font"] = nf
        normal_spec["sz"] = prof.get("body_sz", normal_spec.get("sz"))
        normal_spec["szCs"] = prof.get("body_sz", normal_spec.get("szCs"))
        specs = [normal_spec if s["id"] == "Normal" else s for s in specs]

    # --- styles
    styles = etree.fromstring(z.read("word/styles.xml"))
    by_id = {}
    for st in styles.findall(qn("w:style")):
        by_id[st.get(qn("w:styleId"))] = st
    by_name = {}
    for st in styles.findall(qn("w:style")):
        nm = st.find(qn("w:name"))
        if nm is not None:
            by_name[nm.get(qn("w:val"))] = st

    for spec in specs:
        st = by_id.get(spec["id"])
        if st is None:
            # 允许基于同名样式（如 001 一级标题 与 001 一级标题 的变体写法）匹配
            st = None
            for alt_id, cand in by_id.items():
                nmel = cand.find(qn("w:name"))
                nmv = nmel.get(qn("w:val")) if nmel is not None else ""
                if nmv.replace(" ", "") == spec["name"].replace(" ", ""):
                    st = cand
                    break
        if st is None:
            # 非本口径的样式族缺失，只提示
            add("WARN" if not strict else "FAIL", "styles",
                "存在样式 " + spec["name"], "存在", "缺失")
            continue
        inherited = spec.get("based_on") or ""
        rpr = st.find(qn("w:rPr"))
        # 字符型样式常只写差异项（继承段落样式），未显式写完不算错
        is_char = spec["type"] == "character"
        f = rpr.find(qn("w:rFonts")) if rpr is not None else None
        if spec.get("font") and not is_char:
            if f is None:
                add("WARN" if not strict else "FAIL", "styles", "%s 字体" % spec["name"],
                    "%s/%s" % (spec["font"]["east_asia"], spec["font"]["ascii"]), "未显式设置（继承）")
            else:
                got = (f.get(qn("w:eastAsia")), f.get(qn("w:ascii")))
                want = (spec["font"]["east_asia"], spec["font"]["ascii"])
                if got != want:
                    # 只写 eastAsia（ascii 由 docDefaults 继承为 Times New Roman）是合法写法
                    if got[0] == want[0] and got[1] is None:
                        add("WARN", "styles", "%s ascii 字体" % spec["name"], want[1], "继承")
                    elif got[0] is None and got[1] is None:
                        add("WARN" if not strict else "FAIL", "styles", "%s 字体" % spec["name"],
                            "/".join(want), "未设置（继承）")
                    else:
                        # Normal 的字体由各文档的 docDefaults 决定，允许差异
                        lvl = "WARN" if (not strict or spec["id"] == "Normal") else "FAIL"
                        add(lvl, "styles", "%s 字体" % spec["name"], "/".join(want), "/".join(map(str, got)))
        for key, tag in (("sz", "sz"), ("szCs", "szCs")):
            if spec.get(key) and not is_char:
                e = rpr.find(qn("w:" + tag)) if rpr is not None else None
                got = e.get(qn("w:val")) if e is not None else None
                if got != str(spec[key]):
                    # Normal 的字号常由 docDefaults 提供，属合法继承
                    lvl = "WARN" if (not strict or spec["id"] == "Normal") else "FAIL"
                    add(lvl, "styles", "%s %s" % (spec["name"], tag), spec[key], got)
        ppr = st.find(qn("w:pPr"))
        # (key, tag, attr, tol, soft)  soft=True 表示该值可随字号换算 / 家族惯例合理变化
        checks = [("jc", "w:jc", "val", None, spec["id"].startswith("TOC")),
                  ("outline", "w:outlineLvl", "val", None, True),
                  ("before_lines", "w:spacing", "beforeLines", None, False),
                  ("before", "w:spacing", "before", None, True),
                  ("after_lines", "w:spacing", "afterLines", None, False),
                  ("after", "w:spacing", "after", None, True),
                  ("line", "w:spacing", "line", None, True),
                  ("first_line_chars", "w:ind", "firstLineChars", None, False),
                  ("first_line", "w:ind", "firstLine", None, True),
                  ("left_chars", "w:ind", "leftChars", None, False),
                  ("left", "w:ind", "left", 4, True)]
        for key, tag, attr, tol, soft in checks:
            if spec.get(key) is None:
                continue
            e = ppr.find(qn(tag)) if ppr is not None else None
            got = e.get(qn("w:" + attr)) if e is not None else None
            if got is None:
                add("WARN" if (not strict or soft) else "FAIL", "styles",
                    "%s.%s" % (spec["name"], attr), spec[key], "未设置（继承）")
            elif tol and near(got, spec[key], tol):
                continue
            elif got != str(spec[key]):
                # 字符单位（firstLineChars）已正确时，twips 的 firstLine 仅作参考
                lvl = "WARN" if (soft and not strict) else "FAIL"
                add(lvl, "styles", "%s.%s" % (spec["name"], attr), spec[key], got)
        for key, tag in (("keep_next", "w:keepNext"), ("keep_lines", "w:keepLines"),
                         ("page_break", "w:pageBreakBefore")):
            if spec.get(key):
                ok = ppr is not None and ppr.find(qn(tag)) is not None
                if not ok:
                    add("FAIL", "styles", "%s.%s" % (spec["name"], tag.split(":")[1]), "有", "无")

    # --- section / page setup
    doc = etree.fromstring(z.read("word/document.xml"))
    sect = doc.find(qn("w:body")).findall(qn("w:sectPr"))
    p = C.PAGE
    for i, sp in enumerate(sect):
        sz = sp.find(qn("w:pgSz")); marg = sp.find(qn("w:pgMar"))
        def cmp(tag, el, attr, want, label, tol=4):
            got = el.get(qn("w:" + attr)) if el is not None else None
            if got is None:
                add("WARN", "page", "%s[%d] %s" % (label, i, attr), want, "未设置")
            elif near(got, want, tol):
                return
            elif got != str(want):
                add("FAIL", "page", "%s[%d] %s" % (label, i, attr), want, got)
        cmp("pgSz", sz, "w", p["page_w"], "纸张宽"); cmp("pgSz", sz, "h", p["page_h"], "纸张高")
        cmp("pgMar", marg, "top", p["margin_top"], "上边距"); cmp("pgMar", marg, "bottom", p["margin_bottom"], "下边距")
        cmp("pgMar", marg, "left", p["margin_left"], "左边距"); cmp("pgMar", marg, "right", p["margin_right"], "右边距")
        cmp("pgMar", marg, "header", p["header_dist"], "页眉距"); cmp("pgMar", marg, "footer", p["footer_dist"], "页脚距")

    # --- numbering
    # 【硬性规则】默认不使用自动编号：序号手写在标题文字里，
    # 因此**缺 numbering.xml 是正确状态，不报任何级别**。
    # 反过来，若文档挂了自动编号（numPr 指向 numbering.xml），
    # 默认判为 FAIL —— 除非使用者明确要求（本次调用传 --auto-numbering）。
    auto_ok = bool(auto_numbering)
    # 判定"是否真的用了自动编号"，只看**会被渲染出来的**引用：
    #   1) 正文段落的 numPr —— 一定渲染，算
    #   2) **自定义样式**上的 numPr —— 作者自己挂的，算
    #   3) Word 内置样式（heading N / List Number / List Bullet…）上的 numPr —— 不算。
    #      那是 Word 默认模板自带的定义，实测客户文件里普遍存在却在正文中从未使用
    #      （3-2 上市保荐书：18 个内置样式有 numPr，正文 numPr=0）。判 FAIL 是误报。
    doc_root = etree.fromstring(z.read("word/document.xml"))
    styles_root = etree.fromstring(z.read("word/styles.xml"))

    def real_numpr(el):
        """该元素是否带**真实**的自动编号引用。

        numId=0 在 OOXML 里表示"无编号"（显式取消的声明），
        Word 生成目录样式（如 `TOC 标题1`）时会写 `<w:numPr><w:numId w:val="0"/></w:numPr>`，
        那是结构性标记、不产生任何编号，必须排除。
        """
        for npr in el.iter(qn("w:numPr")):
            nid = npr.find(qn("w:numId"))
            v = nid.get(qn("w:val")) if nid is not None else None
            # ilvl-only 的 numPr 也视为无效（没有 numId 就没有编号）
            if v is not None and v != "0":
                return True
        return False

    n_body = sum(1 for p in doc_root.iter(qn("w:p")) if real_numpr(p))
    custom_with_numpr = []
    for s in styles_root.findall(qn("w:style")):
        if s.get(qn("w:customStyle")) != "1" or not real_numpr(s):
            continue
        nm = s.find(qn("w:name"))
        custom_with_numpr.append(nm.get(qn("w:val")) if nm is not None else s.get(qn("w:styleId")))
    if not auto_ok:
        if n_body or custom_with_numpr:
            add("FAIL", "numbering", "不应使用自动编号（默认手写序号）",
                "正文与自定义样式均无 numPr",
                "正文numPr=%d 自定义样式numPr=%s" % (
                    n_body, ",".join(custom_with_numpr) or "无"))
    if "word/numbering.xml" in names and auto_ok:
        num = etree.fromstring(z.read("word/numbering.xml"))
        got = {}
        for an in num.findall(qn("w:abstractNum")):
            for lv in an.findall(qn("w:lvl")):
                it = lv.get(qn("w:ilvl"))
                lvtext = lv.find(qn("w:lvlText")); fmt = lv.find(qn("w:numFmt"))
                pst = lv.find(qn("w:pStyle"))
                got[it] = (lvtext.get(qn("w:val")) if lvtext is not None else None,
                           fmt.get(qn("w:val")) if fmt is not None else None,
                           pst.get(qn("w:val")) if pst is not None else None)
        expected = C.NUMBERING_ANNOUNCE if profile in ("公告", "公文") else C.NUMBERING
        for lv in expected:
            k = str(lv["ilvl"])
            if k not in got:
                add("FAIL", "numbering",
                    "第%d级编号" % (lv["ilvl"] + 1), lv["text"], "缺失")
                continue
            t, f, s = got[k]
            if t != lv["text"] or f != lv["fmt"]:
                add("WARN" if hand_written else "FAIL", "numbering",
                    "第%d级编号" % (lv["ilvl"] + 1),
                    "%s / %s" % (lv["text"], lv["fmt"]), "%s / %s" % (t, f))
            if s != lv["style"]:
                add("WARN", "numbering", "第%d级绑定样式" % (lv["ilvl"] + 1), lv["style"], s)

    # --- tables（按该体裁的口径校验）
    # 【硬性规则】默认校验「表格美颜② 四边线表」。只有"公告/公文"这两种
    # 由交易所规则指定 0.5pt 细线格的体裁例外；使用者若明确要求其他线型，
    # 用 --pattern 覆盖。
    if profile in ("公告", "公文"):
        pat_key = "announcement"
    else:
        pat_key = pattern or C.TABLE["default_pattern"]
    pat = C.TABLE_PATTERNS[pat_key]
    tbls = doc.findall(".//" + qn("w:tbl"))
    for ti, tbl in enumerate(tbls):
        tblPr = tbl.find(qn("w:tblPr"))
        w = tblPr.find(qn("w:tblW")) if tblPr is not None else None
        b = tblPr.find(qn("w:tblBorders")) if tblPr is not None else None
        rows = tbl.findall(qn("w:tr"))
        cols = len(list(tbl.iter(qn("w:gridCol"))))
        # 判定是否"数据表"：≥3 列、≥3 行、首行每格都是短文本（标题行特征）
        header_cells = rows[0].findall(qn("w:tc")) if rows else []
        header_texts = ["".join(x.text or "" for x in tc.iter(qn("w:t"))).strip() for tc in header_cells]
        looks_header = (bool(header_texts) and all(0 < len(t) <= 24 for t in header_texts))
        is_data = (len(rows) >= 3 and cols >= 3 and looks_header)
        has_tblheader = (rows and rows[0].find(qn("w:trPr")) is not None
                         and rows[0].find(qn("w:trPr")).find(qn("w:tblHeader")) is not None)
        lvl_fail = "FAIL" if strict else "WARN"
        # 落款/签字/双列标签/版式表：不参与数据表规范校验
        if not (is_data or has_tblheader):
            continue
        if w is None or w.get(qn("w:type")) != "pct" or w.get(qn("w:w")) != str(C.TABLE["width_pct"]):
            got_w = w.get(qn("w:w")) if w is not None else None
            lvl = "WARN"
            if w is None or w.get(qn("w:type")) != "pct":
                lvl = lvl_fail
            add(lvl, "table", "表%d 宽度" % (ti + 1), "%s pct" % C.TABLE["width_pct"],
                "%s %s" % (got_w, w.get(qn("w:type")) if w is not None else None))
        if b is None:
            add(lvl_fail, "table", "表%d 边框" % (ti + 1), pat["label"], "未设置")
        else:
            for tag, want in pat["borders"].items():
                e = b.find(qn("w:" + tag))
                if want["val"] == "none":
                    if e is not None and e.get(qn("w:val")) not in (None, "none", "nil"):
                        add(lvl_fail, "table", "表%d %s 应为无框" % (ti + 1, tag),
                            "none", e.get(qn("w:val")))
                    continue
                if e is None or e.get(qn("w:val")) != want["val"] or e.get(qn("w:sz")) != str(want["sz"]):
                    add(lvl_fail, "table", "表%d %s 线" % (ti + 1, tag),
                        "%s/%spt" % (want["val"], want["sz"] / 8.0),
                        "%s/%s" % (e.get(qn("w:val")), e.get(qn("w:sz"))) if e is not None else "无")
        if rows and not has_tblheader:
            add(lvl_fail, "table", "表%d 表头跨页重复" % (ti + 1), "有", "无")
        for ri, tr in enumerate(rows):
            trPr = tr.find(qn("w:trPr"))
            h = trPr.find(qn("w:trHeight")) if trPr is not None else None
            if h is None:
                add("WARN", "table", "表%d 第%d行 行高" % (ti + 1, ri + 1), "atLeast/%s" % pat["row_height"], "无")
            elif h.get(qn("w:hRule")) != "atLeast":
                add("WARN", "table", "表%d 第%d行 行高规则" % (ti + 1, ri + 1), "atLeast", h.get(qn("w:hRule")))

    # --- header / footer
    hdrs = [n for n in names if re.match(r"word/header\d*\.xml", n)]
    ftrs = [n for n in names if re.match(r"word/footer\d*\.xml", n)]
    if not sect or not hdrs:
        add("WARN", "header", "页眉文件", "存在 headerN.xml", "缺失")
    for h in hdrs:
        root = etree.fromstring(z.read(h))
        szs = [e.get(qn("w:val")) for e in root.iter(qn("w:sz"))]
        if str(C.HEADER_FOOTER["header_sz"]) not in szs and szs:
            add("WARN", "header", h + " 字号", C.HEADER_FOOTER["header_sz"], ",".join(map(str, szs)))
        # 公告/公文体裁页眉按体裁惯例处理：公告留白；公文不设页眉规则
        if profile not in ("公告", "公文"):
            if root.find(".//" + qn("w:pBdr")) is None:
                add("FAIL", "header", h + " 页眉下框线", "single/4", "无")
        elif profile == "公告":
            htxt = "".join(e.text or "" for e in root.iter(qn("w:t"))).strip()
            if htxt:
                add("WARN", "header", h + " 应为留白", "空", htxt[:30])
    # 页码域：只要求"默认页脚"（footer 中对应 default 的那个）带 PAGE
    for f in ftrs:
        root = etree.fromstring(z.read(f))
        codes = "".join((e.text or "") for e in root.iter(qn("w:instrText")))
        has_page = "PAGE" in codes
        if not has_page and profile not in ("公告", "公文"):
            add("FAIL", "footer", f + " 页码域", "PAGE", codes.strip() or "无")
        elif profile in ("公告", "公文") and re.search(r"footer\d*\.xml$", f):
            # 转换产物可能把域包在 sdt 里，检查任意页脚是否有 PAGE
            pass
    if profile in ("公告", "公文"):
        any_page = False
        for f in ftrs:
            root = etree.fromstring(z.read(f))
            if "PAGE" in "".join((e.text or "") for e in root.iter(qn("w:instrText"))):
                any_page = True
                break
        if not any_page and ftrs:
            add("WARN", "footer", "页码域", "PAGE（至少一个页脚）", "无")

    ok = not any(x["level"] == "FAIL" for x in findings)
    if as_json:
        print(json.dumps({"file": path, "ok": ok, "findings": findings}, ensure_ascii=False, indent=1))
    else:
        fails = [x for x in findings if x["level"] == "FAIL"]
        warns = [x for x in findings if x["level"] == "WARN"]
        print("文件: %s" % path)
        print("结果: %s  (FAIL %d / WARN %d)" % ("通过" if ok else "不通过", len(fails), len(warns)))
        for x in fails:
            print("  [FAIL] %-9s %-34s 应=%-26s 实=%s" % (x["area"], x["rule"], x["expected"], x["actual"]))
        for x in warns:
            print("  [WARN] %-9s %-34s 应=%-26s 实=%s" % (x["area"], x["rule"], x["expected"], x["actual"]))
    return ok, findings


# --------------------------------------------------------------- CLI
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["apply", "check", "new", "contract"])
    ap.add_argument("target", nargs="?")
    ap.add_argument("--in", dest="src")
    ap.add_argument("--out", dest="dst")
    ap.add_argument("--profile", default="全部")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--auto-numbering", dest="auto_numbering", action="store_true",
                    help="明确要求时才启用自动多级编号；默认序号手写在标题文字里")
    ap.add_argument("--pattern", default=None,
                    help="明确指定表格线型；默认 boxed（表格美颜② 四边线表）")
    ap.add_argument("--strict", action="store_true",
                    help="check 时把可继承/可容忍的差异也报为 FAIL")
    a = ap.parse_args()

    if a.cmd == "contract":
        print(json.dumps({"page": C.PAGE, "para_styles": C.PARA_STYLES,
                          "table": C.TABLE, "numbering": C.NUMBERING,
                          "header_footer": C.HEADER_FOOTER, "text_rules": C.TEXT_RULES},
                         ensure_ascii=False, indent=1))
        return 0
    if a.cmd == "check":
        ok, _ = check(a.target, a.json, strict=getattr(a, "strict", False),
                      profile=getattr(a, "profile", "全部"),
                      auto_numbering=getattr(a, "auto_numbering", False),
                      pattern=getattr(a, "pattern", None))
        return 0 if ok else 1
    if a.cmd == "new":
        doc = new_document(a.profile, auto_numbering=a.auto_numbering)
        doc.save(a.dst)
        print("已生成:", a.dst)
        return 0
    if a.cmd == "apply":
        shutil.copyfile(a.src, a.dst)
        doc = Document(a.dst)
        auto = getattr(a, "auto_numbering", False)
        write_styles(doc, auto_numbering=auto)
        if auto:
            write_numbering(doc)
        for i, sec in enumerate(doc.sections):
            setup_section(sec, cover=(i == 0), body=(i > 0))
        doc.save(a.dst)
        ok, _ = check(a.dst, profile=a.profile)
        print("已写入样式与版式:", a.dst)
        return 0 if ok else 1
    return 2


if __name__ == "__main__":
    sys.exit(main())
