# -*- coding: utf-8 -*-
"""
fix_announcement.py — 把交易所临时公告（公告格式指引模板/公司公告）整改为标准格式

体裁：上交所/深交所临时公告。与申报文件的差异（不要混用）：
  · 正文 仿宋_GB2312 小三（15pt / sz=30），固定行距 28 磅（line=560 exact）
  · 章节标题 黑体 15pt；要点/注释 12pt（sz=24）
  · 全文左对齐（公文惯例，不用两端对齐）
  · 表格 0.5pt 四边细线格，行高最小值 0.7cm，表头跨页重复
  · 页码居中，页眉留白（公告不设页眉）

整改内容（只动格式，不动文字）：
  1) 注入公告体裁样式集（AnnTitle/AnnH1/AnnBody/AnnNote/AnnSign…）
  2) 按文本特征把每段归到标准样式，清掉直接格式残留
  3) 所有表格套公告表格线 + 表头重复 + 行高 + 数字右对齐 + 合计行加粗
  4) 统一字体：清掉 等线/宋体/仿宋 混用，中文一律 仿宋_GB2312（标题黑体），
     西文统一 Times New Roman 或与中文一致
  5) 统一行距：exact 560（正文/标题）、360 auto（注释）
  6) 修正编号悬空引用（numId 指向不存在的编号定义时补上定义或清除引用）

用法：
  python fix_announcement.py 输入.docx 输出.docx [--dry-run] [--no-numbering]
"""
import argparse, copy, os, re, shutil, sys, zipfile
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

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def q(t):
    """命名空间查询名。必须传裸局部名（"abstractNum"），传 "w:..." 会静默查不到。"""
    return "{%s}%s" % (W, t)

# ---------------------------------------------------------------- 段落角色识别
RE_TITLE = re.compile(r"^第[一二三四五六七八九十百]+号\s")
RE_SUBJECT = re.compile(r"^[X×]*[^。；，]{0,40}(股份有限公司|有限公司).{0,30}公告$")
RE_H1 = re.compile(r"^[一二三四五六七八九十]+、")
RE_H2 = re.compile(r"^（[一二三四五六七八九十]+）")
RE_H3 = re.compile(r"^\d+[.、]")
RE_H4 = re.compile(r"^（\d+）")
RE_NOTE = re.compile(r"^注\s*\d+[：:]")
RE_SIGN = re.compile(r"^(特此公告。|[X×]*[^。]{0,20}(董事会|监事会|股份有限公司)|年\s+月\s+日)$")
RE_META = re.compile(r"^证券代码[：:]")
RE_BOILERPLATE = re.compile(r"本公司董事会及全体董事保证")
RE_SECTION = re.compile(r"^(上网公告文件|报备文件|重要内容提示[：:]|适用情形[：:])")
RE_LABEL = re.compile(r"^(适用情形[：:]|重要内容提示[：:]|上网公告文件|报备文件)$")


def classify(text):
    """返回公告体裁样式名；None 表示留空段不动。

    实证公告的加粗分布（据此定加粗，不凭感觉）：
      · 标题行（黑体 15pt）— 不额外加粗（黑体本身即粗体视觉）
      · 标签行（适用情形：/重要内容提示：/上网公告文件/报备文件）— 仿宋加粗
      · 正文与其他各级 — 不加粗
    """
    t = (text or "").strip()
    if not t:
        return None
    if RE_LABEL.match(t):
        return "公告标签"
    if RE_TITLE.match(t):
        return "公告标题"
    if RE_H1.match(t):
        return "公告一级"
    if RE_H2.match(t):
        return "公告二级"
    if RE_H3.match(t):
        return "公告三级"
    if RE_H4.match(t):
        return "公告三级"
    if RE_NOTE.match(t):
        return "公告注释"
    if RE_SUBJECT.search(t) and len(t) <= 60:
        return "公告文种标题"
    if RE_SIGN.match(t):
        return "公告落款"
    if len(t) <= 40 and (t.endswith("：") or t.endswith(":")):
        return "公告标签"
    if len(t) <= 50 and (t.endswith("表：") or t.endswith("清单") or t.endswith("结构表")):
        return "公告三级"
    return "公告正文"


# ---------------------------------------------------------------- 字体统一
def set_normal_font(doc, spec):
    """把 Normal 的中文字体**与字号**切到体裁正文值。

    未挂样式的段落（空段、转换残留）会落到 Normal，不改就会串成宋体；
    字号同理：源文件 Normal 常带 sz=21 / szCs=24 这种小字号且两个值自相矛盾
    （21 与 24 指同一字体却不同大小），会让兜底文字比正文小一档。
    """
    f = spec["font"]
    for s in doc.styles.element.findall(qn("w:style")):
        if s.get(qn("w:type")) != "paragraph":
            continue
        nm = s.find(qn("w:name"))
        if nm is None or nm.get(qn("w:val")) != "Normal":
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
        # sz 与 szCs 必须同值且等于体裁正文字号
        for tag in ("w:sz", "w:szCs"):
            e = rpr.find(qn(tag))
            if e is None:
                e = etree.SubElement(rpr, qn(tag))
            e.set(qn("w:val"), str(spec["sz"]))
        return True
    return False


def set_doc_defaults(doc, spec):
    """把 docDefaults 的中文字体设为该体裁的正文字体。
    否则未显式设字体的文本会回落到宋体，与仿宋正文串体。"""
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
    return True


def unify_run_font(run, spec):
    """把 run 的字体/字号压到合同值，清掉等线/宋体/仿宋混用。"""
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = etree.Element(qn("w:rFonts"))
        rpr.insert(0, rf)
    font = spec["font"]
    rf.set(qn("w:ascii"), font["ascii"])
    rf.set(qn("w:hAnsi"), font.get("hAnsi", font["ascii"]))
    rf.set(qn("w:eastAsia"), font["east_asia"])
    rf.set(qn("w:cs"), font["ascii"])
    if "hint" in rf.attrib:
        del rf.attrib[qn("w:hint")]
    # 字号
    for tag in ("w:sz", "w:szCs"):
        e = rpr.find(qn(tag))
        if e is None:
            e = etree.SubElement(rpr, qn(tag))
        e.set(qn("w:val"), str(spec["sz"]))
    # 加粗：标题族显式，正文一律不加粗
    b = rpr.find(qn("w:b"))
    want_bold = bool(spec.get("bold"))
    if want_bold and b is None:
        b = etree.SubElement(rpr, qn("w:b"))
        b.set(qn("w:val"), "1")
    elif not want_bold and b is not None:
        rpr.remove(b)


def strip_direct_paragraph_format(p, drop_numbering=False):
    """清掉段落的直接缩进/行距/对齐，交由样式决定。

    默认保留 numPr：公告模板里的"⚫ 要点"与"七、股本变动结构表"等
    都是**有效的自动编号**（.doc 转换后编号定义仍在 numbering.xml 里），
    删掉会丢内容。只有确认 numId 无对应定义时才用 drop_numbering=True 清除。
    """
    ppr = p._p.find(qn("w:pPr"))
    if ppr is None:
        return
    tags = ["w:spacing", "w:jc"]
    has_num = ppr.find(qn("w:numPr")) is not None
    # 带编号的段落不删 ind：缩进由编号定义提供，直接删会改变悬挂位置
    if not has_num:
        tags.append("w:ind")
    if drop_numbering:
        tags.append("w:numPr")
    for tag in tags:
        e = ppr.find(qn(tag))
        if e is not None:
            ppr.remove(e)


# ---------------------------------------------------------------- 编号修复
def fix_numbering(doc, add_announce_list=True):
    """把手写编号口径落到公告上：**去掉自动编号，序号码字面化**（规则 1）。

    公告模板里的编号有三类，处理方式不同：
      · `⚫` 项目符号（Wingdings bullet）—— 去掉符号，保留纯文字。
        实测 `⚫`(U+26AB) 在仿宋/黑体里并不存在，原文件靠 Wingdings 字体渲染；
        若字面化成 `●` 会与原文观感不符，使用者明确要求改为纯文字。
      · 中文序号（如模板漏写、靠自动编号补出的 `七、`）—— **字面化为序号文字**。
        同一层其余条目都是手写（一、…六、），只有它靠自动编号，属原文的不一致，
        故补齐为字面 `七、` 与其余对齐（这是修复不一致，不是删内容）。
    返回处理统计。
    """
    root = doc.part.numbering_part._element if _has_numbering(doc) else None
    # numId -> 该编号定义的格式（用于判断是符号还是序号）
    kinds = {}
    if root is not None:
        abs_of = {n.get(qn("w:numId")): (n.find(qn("w:abstractNumId")).get(qn("w:val"))
                                        if n.find(qn("w:abstractNumId")) is not None else None)
                  for n in root.findall(qn("w:num"))}
        for an in root.findall(qn("w:abstractNum")):
            lv = an.find(qn("w:lvl"))
            if lv is None:
                continue
            fmt = lv.find(qn("w:numFmt"))
            t = lv.find(qn("w:lvlText"))
            kinds[an.get(qn("w:abstractNumId"))] = (
                fmt.get(qn("w:val")) if fmt is not None else None,
                t.get(qn("w:val")) if t is not None else None)
        for nid, aid in abs_of.items():
            kinds[("numId", nid)] = kinds.get(aid)

    stats = {"bullets_removed": 0, "literals": 0}
    for p in doc.paragraphs:
        ppr = p._p.find(qn("w:pPr"))
        if ppr is None:
            continue
        npr = ppr.find(qn("w:numPr"))
        if npr is None:
            continue
        nid = npr.find(qn("w:numId"))
        nval = nid.get(qn("w:val")) if nid is not None else None
        fmt, lvltext = kinds.get(("numId", nval), (None, None))
        # 去掉编号引用
        ppr.remove(npr)
        # 序号型（非 bullet）→ 字面化
        if fmt and fmt != "bullet":
            text = p.text or ""
            if not re.match(r"^[一二三四五六七八九十百]+、", text.strip()):
                # 用编号定义的 start 推出该条序号；模板里 start=7 -> 七、
                n = _start_of(doc, nval)
                lit = "%s、" % C.cn_num(n)
                if p.runs:
                    p.runs[0].text = lit + p.runs[0].text
                else:
                    p.add_run(lit)
                stats["literals"] += 1
        else:
            stats["bullets_removed"] += 1
    # 编号引用清空后，把整个 numbering 部件也摘掉（规则 1）
    F.drop_numbering(doc)
    return stats


def _has_numbering(doc):
    try:
        doc.part.numbering_part
        return True
    except Exception:
        return False


def _start_of(doc, numId):
    """读该编号定义的 start 值（模板 `七、` 的 start=7）。失败返回 7。"""
    try:
        root = doc.part.numbering_part._element
        for n in root.findall(qn("w:num")):
            if n.get(qn("w:numId")) != numId:
                continue
            aid = n.find(qn("w:abstractNumId")).get(qn("w:val"))
            for an in root.findall(qn("w:abstractNum")):
                if an.get(qn("w:abstractNumId")) != aid:
                    continue
                lv = an.find(qn("w:lvl"))
                st = lv.find(qn("w:start")) if lv is not None else None
                if st is not None:
                    return int(st.get(qn("w:val")))
    except Exception:
        pass
    return 7


# ---------------------------------------------------------------- 主流程
def fix(src, dst, dry_run=False, unify_fonts=True):
    shutil.copyfile(src, dst)
    doc = Document(dst)

    # 1) 注入公告样式集（保留文档原有样式，只增补）
    F.write_announce_styles(doc)

    stats = {"styled": {}, "fonts": 0, "tables": 0, "cleared_blank": 0}

    body_spec = C.style_by_name("公告正文")
    set_doc_defaults(doc, body_spec)
    set_normal_font(doc, body_spec)

    # 2) 逐段归类 + 统一字体
    for p in doc.paragraphs:
        name = classify(p.text)
        if name:
            try:
                p.style = F.get_style(doc, name)
                stats["styled"][name] = stats["styled"].get(name, 0) + 1
            except KeyError:
                continue
            strip_direct_paragraph_format(p)
            spec = C.style_by_name(name)
        else:
            # 空段：不套样式（避免污染样式应用统计），但要显式补回
            # 首行缩进与固定行距，否则空段会塌成单倍行距、正文之间出现间距跳变
            stats["cleared_blank"] += 1
            spec = body_spec
            pPr = p._p.get_or_add_pPr()
            # 空段没有编号，补回首行缩进与固定行距；有编号则只补行距
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
                ind.set(qn("w:firstLine"), "600")
            sp = etree.SubElement(pPr, qn("w:spacing"))
            sp.set(qn("w:line"), "560")
            sp.set(qn("w:lineRule"), "exact")
        if not unify_fonts:
            continue
        for r in p.runs:
            # 空 run（只有段落标记）不设字体，避免往无文字处写冗余 rPr
            if r.text:
                unify_run_font(r, spec)
                stats["fonts"] += 1

    # 3) 表格（公告体裁：0.5pt 细线 + 12pt 仿宋表内文字）
    for t in doc.tables:
        F.format_table(t, "announcement", font_sz=24,
                       cell_font={"east_asia": "仿宋_GB2312", "ascii": "Times New Roman",
                                  "hAnsi": "Times New Roman"})
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        if "hint" in (r._element.rPr.rFonts.attrib if
                                      r._element.rPr is not None and
                                      r._element.rPr.rFonts is not None else {}):
                            del r._element.rPr.rFonts.attrib[qn("w:hint")]
        stats["tables"] += 1

    # 4) 编号
    num_info = fix_numbering(doc)

    # 5) 页面、页眉留白、页脚页码（三种页脚统一）
    for i, sec in enumerate(doc.sections):
        F.setup_section(sec, cover=False, body=True)
        F.clear_header(doc, sec)
        for ftr in (sec.footer, sec.even_page_footer, sec.first_page_footer):
            F.write_page_footer(doc, ftr)

    if not dry_run:
        doc.save(dst)
    return stats, num_info


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--keep-fonts", action="store_true", help="不统一字体（只套样式与表格）")
    a = ap.parse_args()

    stats, num_info = fix(a.src, a.dst, a.dry_run, not a.keep_fonts)
    print("段落样式统计:", stats["styled"])
    print("统一字体的 run 数:", stats["fonts"])
    print("处理表格:", stats["tables"])
    print("编号修复:", num_info)
    if not a.dry_run:
        ok, _ = F.check(a.dst, profile="公告")
        print("自检:", "PASS" if ok else "FAIL（详见 filing_docx.py check）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
