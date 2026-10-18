# -*- coding: utf-8 -*-
"""
format_existing.py — 把一份既有 .docx 整体誊清为投行标准格式

与 filing_docx.py apply 的区别：
  apply            只写样式/编号/页面（不动正文与表格实例格式）
  format_existing  在 apply 之上，再逐段套用样式、逐表美颜、清理楷体与冗余空格、刷新页码域

用法：
  python format_existing.py 输入.docx 输出.docx [--profile 申报文件] [--table boxed]
  python format_existing.py 输入.docx 输出.docx --dry-run      # 只报告将要改动什么

段落归类规则（可按需在 map_paragraph 中调整）：
  以"一、""二、…"开头            → 002 二级标题
  以"（一）""（二）…"开头          → 003 三级标题
  以"1、""2、…"开头               → 004 四级标题
  以"（1）""（2）…"开头            → 011 五级标题
  纯"目  录"/"释  义"/"第X章"      → 001 一级标题
  "单位：…"                       → 009 单位
  "注：…"                         → 006 备注
  其余非空段                       → 005 正文
已带本技能样式的段落不动，避免覆盖作者的人工判断。
"""
import argparse, os, re, shutil, sys, io, zipfile
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

RE_LEVEL1 = re.compile(r"^(目\s*录|释\s*义|第[一二三四五六七八九十百]+[章节]|摘\s*要|前\s*言)$")
RE_LEVEL2 = re.compile(r"^[一二三四五六七八九十]+、")
RE_LEVEL3 = re.compile(r"^（[一二三四五六七八九十]+）")
RE_LEVEL4 = re.compile(r"^\d+、")
RE_LEVEL5 = re.compile(r"^（\d+）")
RE_UNIT = re.compile(r"^单位[：:]")
RE_NOTE = re.compile(r"^注[：:]")

COVER_SZ_MIN = 36        # 18pt；小于此值不视为封面大字


def is_cover_paragraph(p):
    """封面/扉页大字：整段黑体且字号 ≥ 18pt（或居中且字号 ≥ 14pt 的短标题）。
    这类段落由作者手工排版（发行人全名、文件名称、"二〇二五年十月"），
    **不得**被批量归一化为正文字号，否则封面会整体坍缩。"""
    runs = [r for r in p.runs if r.text.strip()]
    if not runs:
        return False
    sizes = []
    for r in runs:
        sz = r._element.rPr.find(qn("w:sz")) if r._element.rPr is not None else None
        val = sz.get(qn("w:val")) if sz is not None else None
        if val:
            sizes.append(int(val))
        elif r.font.size is not None:
            sizes.append(int(r.font.size.pt * 2))
    if not sizes:
        return False
    max_sz = max(sizes)
    if max_sz >= COVER_SZ_MIN:
        return True
    # 居中且 ≥ 14pt 的短句（"保荐机构（主承销商）"）
    centered = p.paragraph_format.alignment is not None and int(p.paragraph_format.alignment) == 1
    return centered and max_sz >= 28 and len(p.text.strip()) <= 24


def map_paragraph(text, in_table=False):
    t = (text or "").strip()
    if not t:
        return None
    if in_table:
        return None
    if RE_UNIT.match(t):
        return "009"
    if RE_NOTE.match(t):
        return "006"
    if RE_LEVEL1.match(t):
        return "001"
    if RE_LEVEL2.match(t):
        return "002"
    if RE_LEVEL3.match(t):
        return "003"
    if RE_LEVEL4.match(t):
        return "004"
    if RE_LEVEL5.match(t):
        return "011"
    if len(t) <= 30 and (t.endswith("风险") or t.endswith("情况") or t.endswith("说明")):
        return "004"
    return "005"


def known_style_ids(doc):
    return {s.style_id for s in doc.styles}


def clean_paragraph(p, keep_first_line=True):
    """清掉段内的手工格式残留：全角空格外的多余空格、制表符、楷体。"""
    for r in p.runs:
        if not r.text:
            continue
        txt = r.text
        if "楷体" in (r.font.name or "") or "楷体" in (
                r._element.rPr is not None and r._element.rPr.rFonts is not None
                and (r._element.rPr.rFonts.get(qn("w:eastAsia")) or "") or ""):
            r.font.name = "Times New Roman"
            r._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
        # 行首行尾的多余半角空格（不触碰中段的词间空格）
        if not keep_first_line:
            txt = txt.strip(" ")
        if txt != r.text:
            r.text = txt


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--profile", default="申报文件")
    ap.add_argument("--table", default="boxed", choices=list(C.TABLE_PATTERNS))
    ap.add_argument("--auto-numbering", dest="auto_numbering", action="store_true",
                    help="明确要求时才启用自动多级编号；默认序号手写在标题文字里")
    ap.add_argument("--keep-cover", dest="keep_cover", action="store_true", default=True,
                    help="保留封面/扉页的手工大字排版（默认开启）")
    ap.add_argument("--no-cover", dest="keep_cover", action="store_false",
                    help="连同封面一起归一化为标准样式（慎用，封面会整体变小）")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    shutil.copyfile(a.src, a.dst)
    doc = Document(a.dst)
    ids = known_style_ids(doc)

    applied = {}
    skipped_cover = 0
    skipped_toc = 0
    skipped_kept = 0
    for p in doc.paragraphs:
        sid = map_paragraph(p.text)
        if sid is None or sid not in ids:
            continue
        # 封面/扉页大字：跳过，保留作者手工排版
        if a.keep_cover and is_cover_paragraph(p):
            skipped_cover += 1
            continue
        # 目录条目：保留 TOC 样式，不要套正文样式（会丢掉点线制表位）
        cur = p.style.style_id if p.style is not None else None
        cur_name = (p.style.name if p.style is not None else "") or ""
        if cur_name.lower().startswith("toc") or cur_name.startswith("目录"):
            skipped_toc += 1
            continue
        if cur in ids and cur in {s["id"] for s in C.PARA_STYLES}:
            skipped_kept += 1             # 已有标准样式，尊重作者选择
            continue
        p.style = F.get_style(doc, sid)
        applied[sid] = applied.get(sid, 0) + 1
        clean_paragraph(p)

    # 默认手写编号：不注入 numbering.xml，样式也不挂 numPr（规则 1）。
    # 只有使用者明确要求（--auto-numbering）才启用自动编号。
    auto = getattr(a, "auto_numbering", False)
    F.write_styles(doc, auto_numbering=auto)
    if auto:
        F.write_numbering(doc)
    else:
        F.drop_numbering(doc)
    for i, sec in enumerate(doc.sections):
        F.setup_section(sec, cover=(i == 0), body=(i > 0))
    for t in doc.tables:
        # 默认四边线表（表格美颜②）；公告/公文由各自脚本走 announcement（规则 2）
        F.format_table(t, a.table)

    # 纯排版不推断数据单位，也不向空白段补写文字。

    if not a.dry_run:
        doc.save(a.dst)
    print("段落样式套用统计:", applied)
    print("跳过封面大字 %d 段 / 跳过目录条目 %d 段 / 保留已有标准样式 %d 段"
          % (skipped_cover, skipped_toc, skipped_kept))
    print("表格美颜:", len(doc.tables), "张 →", a.table)
    print("输出:", a.dst, "(dry-run，未写盘)" if a.dry_run else "")
    ok, _ = F.check(a.dst, profile=a.profile)
    print("自检:", "PASS" if ok else "FAIL（详见 filing_docx.py check）")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
