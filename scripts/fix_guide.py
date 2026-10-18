# -*- coding: utf-8 -*-
"""
fix_guide.py — 把交易所自律监管指南 / 业务办理指南 / 格式指引整改为标准格式

体裁：公文体例（GB/T 9704 系）。与"公告"体裁的差异（混用会串体）：
  · 正文 仿宋_GB2312 14pt（sz=28），固定行距 30 磅（line=600 exact）
  · 一级标题（一、二、三、）是**加粗仿宋同字号**——公文体例典型写法，不是黑体
  · 全文只有文件标题用黑体（20pt 居中）
  · 表内文字 仿宋 10pt（sz=20），表头加粗
  · 全文左对齐（公文惯例）

整改内容（只动格式，不动文字）：
  1) 注入公文体裁样式集
  2) 按文本特征归类每一段（文件标题/公文一级/公文二级/附件标签/附件标题/公文正文/公文注释），
     清掉直接缩进与行距残留
  3) 补齐行距：转换产物常只有少数段落设了行距，其余塌成单倍——一律按体裁补齐
  4) 表格：套公告口径 0.5pt 细线 + **表头跨页重复** + 行高 + 垂直居中 + 数字右对齐 +
     合计行加粗；表内文字 10pt；跨页长表必须带 tblHeader
  5) 统一字体：docDefaults 与 Normal 切到仿宋_GB2312，清掉串体
  6) 页脚页码统一居中（公文页码居中）

用法：
  python fix_guide.py 输入.docx 输出.docx [--dry-run] [--keep-fonts]
"""
import argparse, os, re, shutil, sys
from lxml import etree
from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import filing_docx as F
import filing_contract as C

# ---------------------------------------------------------------- 段落角色识别
RE_TITLE = re.compile(r"^第[一二三四五六七八九十百]+号\s")           # "第一号 信息披露业务办理"
RE_H1 = re.compile(r"^[一二三四五六七八九十]+、")                     # "一、一般规定"
RE_H2 = re.compile(r"^\d+[、.][^0-9]")                                # "1、公告标题" / "1.上市公司…"
RE_ANNEX = re.compile(r"^附件\s*\d+\s*[：:]?$")                        # "附件1：" / "附件2："
RE_NOTE = re.compile(r"^注\s*\d*[：:]")                                # "注：" / "注1："


def classify(text):
    """返回公文体裁样式名；None 表示留空段由调用方单独处理。

    加粗规则按实证文件定：
      · 文件标题 —— 黑体 40 半磅
      · 一级标题（一、）—— 加粗仿宋
      · 附件标签 —— 加粗仿宋
      · 二级标题（1、）与正文 —— 不加粗
    """
    t = (text or "").strip()
    if not t:
        return None
    if RE_TITLE.match(t) and len(t) <= 40:
        return "文件标题"
    if RE_ANNEX.match(t):
        return "附件标签"
    if RE_H1.match(t) and len(t) <= 60:
        return "公文一级"
    if RE_NOTE.match(t):
        return "公文注释"
    if RE_H2.match(t) and len(t) <= 40:
        return "公文二级"
    return "公文正文"


def is_annex_caption(text, prev_text):
    """附件名（"上海证券交易所信息披露公告类别索引"）：紧跟在"附件N："之后、
    居中、短句、不含句号 —— 判为附件标题。"""
    t = (text or "").strip()
    p = (prev_text or "").strip()
    if not t or not RE_ANNEX.match(p):
        return False
    return len(t) <= 40 and not re.search(r"[。；，]", t)


# ---------------------------------------------------------------- 字体统一
def set_doc_defaults(doc, spec):
    styles = doc.styles.element
    dd = styles.find(qn("w:docDefaults"))
    if dd is None:
        dd = etree.Element(qn("w:docDefaults"))
        styles.insert(0, dd)
    rprd = dd.find(qn("w:rPrDefault"))
    if rprd is None:
        rprd = etree.SubElement(dd, qn("w:rPrDefault"))
    rpr = rprd.find(qn("w:rPr"))
    if rpr is None:
        rpr = etree.SubElement(rprd, qn("w:rPr"))
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = etree.Element(qn("w:rFonts"))
        rpr.insert(0, rf)
    f = spec["font"]
    rf.set(qn("w:ascii"), f["ascii"])
    rf.set(qn("w:hAnsi"), f.get("hAnsi", f["ascii"]))
    rf.set(qn("w:eastAsia"), f["east_asia"])
    rf.set(qn("w:cs"), f["ascii"])
    # 清掉主题字体引用，否则会回落到 等线/主题字体
    for k in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        if k in rf.attrib:
            del rf.attrib[qn("w:" + k)]
    return True


def set_normal_font(doc, spec):
    f = spec["font"]
    done = False
    for s in doc.styles.element.findall(qn("w:style")):
        if s.get(qn("w:type")) != "paragraph":
            continue
        nm = s.find(qn("w:name"))
        if nm is None or nm.get(qn("w:val")) not in ("Normal", "正文", "Body Text"):
            continue
        rpr = s.find(qn("w:rPr"))
        if rpr is None:
            rpr = etree.SubElement(s, qn("w:rPr"))
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = etree.Element(qn("w:rFonts"))
            rpr.insert(0, rf)
        rf.set(qn("w:ascii"), f["ascii"])
        rf.set(qn("w:hAnsi"), f.get("hAnsi", f["ascii"]))
        rf.set(qn("w:eastAsia"), f["east_asia"])
        rf.set(qn("w:cs"), f["ascii"])
        for k in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            if k in rf.attrib:
                del rf.attrib[qn("w:" + k)]
        sz = rpr.find(qn("w:sz"))
        if sz is None:
            sz = etree.SubElement(rpr, qn("w:sz"))
        sz.set(qn("w:val"), str(spec["sz"]))
        # szCs 必须与 sz 同值，否则中文/西文的行高按不同字号计算
        szcs = rpr.find(qn("w:szCs"))
        if szcs is None:
            szcs = etree.SubElement(rpr, qn("w:szCs"))
        szcs.set(qn("w:val"), str(spec["sz"]))
        done = True
    return done


def unify_run_font(run, spec):
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = etree.Element(qn("w:rFonts"))
        rpr.insert(0, rf)
    f = spec["font"]
    rf.set(qn("w:ascii"), f["ascii"])
    rf.set(qn("w:hAnsi"), f.get("hAnsi", f["ascii"]))
    rf.set(qn("w:eastAsia"), f["east_asia"])
    rf.set(qn("w:cs"), f["ascii"])
    for k in ("hint", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        if k in rf.attrib:
            del rf.attrib[qn("w:" + k)]
    for tag in ("w:sz", "w:szCs"):
        e = rpr.find(qn(tag))
        if e is None:
            e = etree.SubElement(rpr, qn(tag))
        e.set(qn("w:val"), str(spec["sz"]))
    b = rpr.find(qn("w:b"))
    if spec.get("bold"):
        if b is None:
            b = etree.SubElement(rpr, qn("w:b"))
        b.set(qn("w:val"), "1")
    elif b is not None:
        rpr.remove(b)


def strip_direct(p, keep_num=True):
    """清掉段落直接缩进/行距/对齐（保留 numPr：编号是内容）。"""
    ppr = p._p.find(qn("w:pPr"))
    if ppr is None:
        return
    has_num = ppr.find(qn("w:numPr")) is not None
    tags = ["w:spacing", "w:jc"]
    if not has_num:
        tags.append("w:ind")
    for tag in tags:
        e = ppr.find(qn(tag))
        if e is not None:
            ppr.remove(e)


# ---------------------------------------------------------------- 长表处理
def clean_table_paragraph_marks(table, font_sz):
    """清掉单元格段落标记 (pPr/rPr) 上的字号残留。

    这些是"段落标记"的字号（无文字），不影响可见内容，但会让字号统计出现
    杂值；统一为表内字号，便于校验与后续维护。"""
    n = 0
    for pPr in table._tbl.iter(qn("w:pPr")):
        rpr = pPr.find(qn("w:rPr"))
        if rpr is None:
            continue
        for tag in ("w:sz", "w:szCs"):
            e = rpr.find(qn(tag))
            if e is None:
                continue
            if e.get(qn("w:val")) != str(font_sz):
                e.set(qn("w:val"), str(font_sz))
                n += 1
    return n


def format_long_table(table, font_sz, long_rows=60):
    """公告口径细线格 + 表头跨页重复 + 行高 + 垂直居中 + 合计行加粗。

    两项针对超长表的保守处理（实测踩过）：
      · keep_row_height：≥60 行的表，原作者已按紧凑度设过行高；
        强制统一成 0.7cm 会让 400 行表凭空长出十余页（实测 48→58 页）。
      · keep_cell_align：公告类别索引表这类"索引/目录表"序号列惯用居中，
        不能按财务表的"数字右对齐"规则改。
    """
    rows = len(table.rows)
    is_long = rows >= long_rows
    F.format_table(table, "announcement", font_sz=font_sz,
                   cell_font={"east_asia": "仿宋_GB2312",
                              "ascii": "Times New Roman",
                              "hAnsi": "Times New Roman"},
                   keep_cell_align=True,
                   keep_row_height=is_long,
                   keep_cell_margin=is_long)
    # 表头行：居中（加粗由 format_table 处理）
    for cell in table.rows[0].cells:
        for p in cell.paragraphs:
            p.paragraph_format.alignment = 1
    # 表内段落：单倍行距、无段前段后、无首行缩进
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.first_line_indent = 0
                p.paragraph_format.space_before = 0
                p.paragraph_format.space_after = 0
                p.paragraph_format.line_spacing = 1.0
    return rows


# ---------------------------------------------------------------- 主流程
def fix(src, dst, dry_run=False, unify_fonts=True):
    shutil.copyfile(src, dst)
    doc = Document(dst)

    F.write_guide_styles(doc)
    body_spec = C.style_by_name("公文正文")
    set_doc_defaults(doc, body_spec)
    set_normal_font(doc, body_spec)

    stats = {"styled": {}, "fonts": 0, "tables": 0, "empty_fixed": 0, "captions": 0}

    # ---- 逐段归类
    paras = list(doc.paragraphs)
    prev_text = ""
    for p in paras:
        name = classify(p.text)
        # 附件名特判：紧跟"附件N："且居中短句
        if name == "公文正文" and is_annex_caption(p.text, prev_text):
            name = "附件标题"
            stats["captions"] += 1
        if name:
            try:
                p.style = F.get_style(doc, name)
                stats["styled"][name] = stats["styled"].get(name, 0) + 1
            except KeyError:
                pass
            strip_direct(p)
            spec = C.style_by_name(name)
        else:
            # 空段：补回体裁的缩进与固定行距，避免行距跳变
            stats["empty_fixed"] += 1
            spec = body_spec
            pPr = p._p.get_or_add_pPr()
            for tag in ("w:spacing", "w:jc"):
                e = pPr.find(qn(tag))
                if e is not None:
                    pPr.remove(e)
            if pPr.find(qn("w:numPr")) is None:
                ind = pPr.find(qn("w:ind"))
                if ind is not None:
                    pPr.remove(ind)
                ind = etree.SubElement(pPr, qn("w:ind"))
                ind.set(qn("w:firstLineChars"), "200")
                ind.set(qn("w:firstLine"), "560")
            sp = etree.SubElement(pPr, qn("w:spacing"))
            sp.set(qn("w:line"), "600")
            sp.set(qn("w:lineRule"), "exact")
        if not unify_fonts:
            prev_text = p.text
            continue
        for r in p.runs:
            if r.text:
                unify_run_font(r, spec)
                stats["fonts"] += 1
        prev_text = p.text

    # ---- 表格
    for t in doc.tables:
        format_long_table(t, C.PROFILES["公文"]["table_font_sz"])
        stats["mark_clean"] = stats.get("mark_clean", 0) + clean_table_paragraph_marks(
            t, C.PROFILES["公文"]["table_font_sz"])
        stats["tables"] += 1

    # ---- 规则 1：公文体裁一律手写编号，清掉任何自动编号引用与部件。
    # 文体的序号是"一、/1."手写在文字里的；若来源文件带了 numPr，一并去除。
    n_numpr = len(list(doc.element.body.iter(qn("w:numPr"))))
    if n_numpr:
        for npr in list(doc.element.body.iter(qn("w:numPr"))):
            npr.getparent().remove(npr)
        stats["numpr_removed"] = n_numpr
    F.drop_numbering(doc)

    # ---- 页面、页眉留白、页脚页码居中
    for sec in doc.sections:
        F.setup_section(sec, cover=False, body=True)
        F.clear_header(doc, sec)
        F.write_page_footer(doc, sec.footer)
        for ftr in (sec.even_page_footer, sec.first_page_footer):
            F.write_page_footer(doc, ftr)

    if not dry_run:
        doc.save(dst)
    return stats


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--keep-fonts", action="store_true")
    a = ap.parse_args()

    stats = fix(a.src, a.dst, a.dry_run, not a.keep_fonts)
    print("段落样式统计:", stats["styled"])
    print("附件标题识别: %d | 空段修补: %d | 统一字体 run: %d | 表格: %d" % (
        stats["captions"], stats["empty_fixed"], stats["fonts"], stats["tables"]))
    if not a.dry_run:
        ok, _ = F.check(a.dst, profile="公文")
        print("自检:", "PASS" if ok else "FAIL（详见 filing_docx.py check）")
    return 0 if not a.dry_run else 0


if __name__ == "__main__":
    sys.exit(main())
