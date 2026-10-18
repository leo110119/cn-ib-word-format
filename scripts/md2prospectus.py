# -*- coding: utf-8 -*-
"""
md2prospectus.py — 把 Markdown 文档转成「招股说明书」体裁的 Word 申报文件

招股说明书口径（对应插件"设置为招股说明书常用样式"）：
  · 标题族用「标题 1..7」，序号**手写在标题文字里**（第X节 / X、 / （X） / X、）
    —— 默认不使用自动多级编号，除非使用者明确要求（--auto-numbering）
  · 正文用「正文缩进」：宋体 小四 12pt、1.5 倍行距、首行缩进 2 字符、两端对齐
  · 表格默认用「表格美颜② 四边线表」（boxed）；除非使用者明确指定其他线型
  · 表头跨页重复
  · 页眉：发行人全称 + 空格 + 文件全名；页脚：居中页码
  · 一/二级标题带分页控制，一级标题分页符前

Markdown 映射规则：
  #   H1  -> 标题 1        （分页 + 自动编号「第X节」）
  ##  H2  -> 标题 2        （自动编号「X、」）
  ### H3  -> 标题 3        （自动编号「（X）」）
  #### H4 -> 标题 4
  段落             -> 正文缩进
  - /*  无序列表    -> 正文缩进（转为「· 」前缀，保持内容不丢）
  1.    有序列表    -> 正文缩进（保留原序号文字）
  | 表格 |         -> 三线表
  > 引用            -> 006备注（小一档字号）
  ```代码块```      -> 正文缩进（保留原文与缩进）
  --- 分隔线        -> 忽略（改由标题分页控制）

用法：
  python md2prospectus.py 输入.md 输出.docx [--title 主标题] [--org 发行人]
"""
import argparse, os, re, sys
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import filing_docx as F
import filing_contract as C

CODE_FONT = {"east_asia": "宋体", "ascii": "Courier New", "hAnsi": "Courier New"}

# Markdown 到招股书标题层级映射。
# 招股书的层级：标题1 = 「第X节」，标题2 = 「X、」，标题3 = 「（X）」，标题4 = 「X、」。
# Markdown 文档通常只有一个 # 作为文档标题，真正的章节是 ##，所以：
#   #    -> 封面/文档标题（不进编号体系）
#   ##   -> 标题 1「第X节」   ← 正文里「第 N 节」的引用正好对上
#   ###  -> 标题 2「X、」
#   #### -> 标题 3「（X）」
# 若直接按 # -> 标题1 映射，# 会被编号成"第一节"，而 ## 变成"一、"，
# 正文引用「第 1 节」就会指向错误的位置（实测踩过）。
#
# 【编号规则】默认**手写编号**（不使用自动多级编号）。除非使用者明确要求
# 用自动编号，否则序号直接拼进标题文字，见 build_heading_text()。
LEVEL_MAP = {1: None, 2: "标题 1", 3: "标题 2", 4: "标题 3", 5: "标题 4", 6: "标题 4"}
LEVEL_NUM = {2: 1, 3: 2, 4: 3, 5: 4, 6: 4}      # Markdown 级别 -> 编号层级（1 起）
HEADING_LEVEL_REV = {1: 2, 2: 3, 3: 4, 4: 5}     # 编号层级 -> Markdown 级别

# 自动多级编号会把序号加在标题文字前（第%1节 / %2、 / （%3） / %4、）。
# Markdown 标题里若已有手写序号（"1. " / "一、" / "（一）" / "第一节"），
# 必须剥掉，否则会渲染成"一、1. 可复制简报"这种重复编号。
MANUAL_NUM_PATTERNS = [
    r"^第[一二三四五六七八九十百]+[节章部分]\s*[、.．:：]?\s*",
    r"^[一二三四五六七八九十]+[、.．]\s*",
    r"^（[一二三四五六七八九十]+）\s*",
    r"^\([一二三四五六七八九十]+\)\s*",
    r"^（\d+）\s*",
    r"^\(\d+\)\s*",
    r"^\d+[、.．)]\s*",
]


def strip_manual_number(text):
    """去掉标题开头的手写序号（自动编号会补上）。"""
    t = (text or "").strip()
    for pat in MANUAL_NUM_PATTERNS:
        t2 = re.sub(pat, "", t, count=1)
        if t2 != t and t2.strip():
            return t2.strip()
    return t


CN_NUM = "一二三四五六七八九十"


def cn_seq(n):
    """1..99 -> 一/二/…/十一/二十/二十一"""
    if n <= 10:
        return CN_NUM[n - 1]
    if n < 20:
        return "十" + CN_NUM[n - 11]
    tens, ones = divmod(n, 10)
    return CN_NUM[tens - 1] + "十" + (CN_NUM[ones - 1] if ones else "")


def remap_section_refs(text, heading_ordinals):
    """把「第 N 节」按 heading_ordinals（原序号 -> 新序号）改写，
    使正文交叉引用与新编号一致。"""
    def repl(m):
        n = int(m.group(1))
        new = heading_ordinals.get(n)
        if new is None:
            return m.group(0)
        return "第%s节" % cn_seq(new)
    return re.sub(r"第\s*(\d+)\s*节", repl, text)


def build_heading_text(raw, level, counters):
    """把手写编号拼进标题文字（默认不使用自动多级编号）。

    counters: {编号层级: 已出现次数}，本函数就地递增并清掉下级计数。
    返回 (最终显示文本, 该级序号)
    """
    lvl = LEVEL_NUM.get(level)
    text = strip_manual_number(raw)
    if lvl is None:
        return text, None
    counters[lvl] = counters.get(lvl, 0) + 1
    for k in list(counters):
        if k > lvl:
            del counters[k]
    n = counters[lvl]
    prefix = C.manual_number("doc_prospectus", lvl, n)
    return (prefix + text) if prefix else text, n


# ---------------------------------------------------------------- markdown 解析
def parse_md(text):
    """把 Markdown 拆成块序列 [(kind, payload)]。
    kind: h1..h4 / p / bullet / ordered / table / quote / code / hr
    """
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    blocks = []
    i = 0
    while i < len(lines):
        raw = lines[i]
        s = raw.strip()

        # 代码块
        if s.startswith("```"):
            i += 1
            buf = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                buf.append(lines[i])
                i += 1
            i += 1
            blocks.append(("code", buf))
            continue

        # 空行
        if not s:
            i += 1
            continue

        # 分隔线（---/***/___ 三个以上）
        if re.match(r"^(-{3,}|\*{3,}|_{3,})$", s):
            blocks.append(("hr", None))
            i += 1
            continue

        # 标题
        m = re.match(r"^(#{1,6})\s*(.*)$", s)
        if m:
            lvl = min(len(m.group(1)), 4)
            blocks.append(("h%d" % lvl, m.group(2).strip()))
            i += 1
            continue

        # 表格（连续的 | 开头行，且含分隔行）
        if s.startswith("|") and i + 1 < len(lines) and re.match(r"^\|[\s:\-|]+\|$", lines[i + 1].strip()):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append(lines[i].strip())
                i += 1
            blocks.append(("table", rows))
            continue

        # 引用
        if s.startswith(">"):
            buf = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                buf.append(re.sub(r"^>\s?", "", lines[i].strip()))
                i += 1
            blocks.append(("quote", buf))
            continue

        # 无序列表
        if re.match(r"^[-*+]\s+", s):
            buf = []
            while i < len(lines) and re.match(r"^[-*+]\s+", lines[i].strip()):
                buf.append(re.sub(r"^[-*+]\s+", "", lines[i].strip()))
                i += 1
            blocks.append(("bullet", buf))
            continue

        # 有序列表
        if re.match(r"^\d+[.)]\s+", s):
            buf = []
            while i < len(lines) and re.match(r"^\d+[.)]\s+", lines[i].strip()):
                buf.append(lines[i].strip())
                i += 1
            blocks.append(("ordered", buf))
            continue

        # 普通段落（吸收后续非空、非结构行）
        buf = [s]
        i += 1
        while i < len(lines):
            nxt = lines[i].strip()
            if (not nxt or nxt.startswith(("|", ">", "#", "```"))
                    or re.match(r"^[-*+]\s+", nxt) or re.match(r"^\d+[.)]\s+", nxt)
                    or re.match(r"^(-{3,}|\*{3,}|_{3,})$", nxt)):
                break
            buf.append(nxt)
            i += 1
        blocks.append(("p", "".join(buf)))
    return blocks


# ---------------------------------------------------------------- 行内格式
def add_runs(p, text, font=None, sz=None, bold_default=False):
    """写入行内文本，处理 **粗体** `代码` 与 [链接](url)。
    不丢任何文字：链接保留「显示文字（url）」形式。"""
    text = re.sub(r"`([^`]+)`", lambda m: m.group(1), text)
    text = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", lambda m: "%s（%s）" % (m.group(1), m.group(2)), text)
    parts = re.split(r"(\*\*[^*]+\*\*)", text)
    for part in parts:
        if not part:
            continue
        bold = bold_default
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            part = part[2:-2]
            bold = True
        r = p.add_run(part)
        if font:
            r.font.name = font["hAnsi"]
            r._element.rPr.rFonts.set(qn("w:eastAsia"), font["east_asia"])
        if sz:
            r.font.size = Pt(sz / 2)
        r.font.bold = bold
    return p


def add_body(doc, text, style="正文缩进", font=None, sz=None):
    p = doc.add_paragraph(style=F.get_style(doc, style))
    add_runs(p, text, font=font, sz=sz)
    return p


# ---------------------------------------------------------------- 表格
def add_table(doc, md_rows, pattern="boxed", keep_grid=False):
    """Markdown 表格 -> Word 表格。

    默认 pattern="boxed"（IBD Tools「表格美颜② 四边线表」）：
    四边 `single sz=12`（1.5pt）+ 内部 `single sz=4`（0.5pt）。
    只有使用者明确要求时才换成 threeline（三线表）/ blueprint（蓝线密集表）。
    """
    def cells(line):
        s = line.strip()
        if s.startswith("|"):
            s = s[1:]
        if s.endswith("|"):
            s = s[:-1]
        return [c.strip() for c in s.split("|")]

    header = cells(md_rows[0])
    body = [cells(r) for r in md_rows[2:]]
    ncol = max([len(header)] + [len(r) for r in body]) if body else len(header)
    t = doc.add_table(rows=1 + len(body), cols=ncol)
    try:
        t.style = F.get_style(doc, "报告表格")
    except KeyError:
        pass
    for c, txt in enumerate(header + [""] * (ncol - len(header))):
        cell = t.rows[0].cells[c]
        cell.text = ""
        add_runs(cell.paragraphs[0], txt)
    for ri, row in enumerate(body, start=1):
        for c in range(ncol):
            txt = row[c] if c < len(row) else ""
            cell = t.rows[ri].cells[c]
            cell.text = ""
            add_runs(cell.paragraphs[0], txt)
    F.format_table(t, pattern, font_sz=21)
    if keep_grid and ncol > 1:
        tblPr = t._tbl.tblPr
        b = tblPr.find(qn("w:tblBorders"))
        if b is not None:
            iv = b.find(qn("w:insideV"))
            if iv is None:
                iv = OxmlElement("w:insideV")
                b.append(iv)
            iv.set(qn("w:val"), "single")
            iv.set(qn("w:sz"), "4")
            iv.set(qn("w:space"), "0")
            iv.set(qn("w:color"), "auto")
            ih = b.find(qn("w:insideH"))
            if ih is None:
                ih = OxmlElement("w:insideH")
                b.append(ih)
            ih.set(qn("w:val"), "single")
            ih.set(qn("w:sz"), "4")
            ih.set(qn("w:space"), "0")
            ih.set(qn("w:color"), "auto")
    return t


# ---------------------------------------------------------------- 主流程
def convert(src, dst, title=None, org="XX股份有限公司", doc_name=None,
            pattern="boxed", with_cover=True, with_toc=True,
            auto_numbering=False):
    md = open(src, encoding="utf-8").read()
    blocks = parse_md(md)

    # 默认手写编号：不注入 numbering.xml，标题样式也不挂 numPr。
    # 只有使用者明确要求（auto_numbering=True）才启用自动多级编号。
    doc = F.new_document(profile="招股说明书", auto_numbering=auto_numbering)
    # 标题优先级：--title 显式指定 > 原文 H1 > 文件名。
    # 不要擅自改写原文标题（实测踩过：把原文长标题换成短标题即为改动内容）。
    src_h1 = None
    for kind, payload in blocks:
        if kind == "h1" and (payload or "").strip():
            src_h1 = payload.strip()
            break
    doc_name = title or src_h1 or os.path.splitext(os.path.basename(src))[0]

    # 预扫描 H2 的手写序号，建立「原序号 -> 新序号」映射，
    # 供正文交叉引用（"第 N 节"）改写使用，避免引用与编号不一致。
    heading_ordinals = {}
    seq = 0
    for kind, payload in blocks:
        if kind == "h2":
            seq += 1
            m = re.match(r"^(\d+)[、.．)]", (payload or "").strip())
            if m:
                heading_ordinals[int(m.group(1))] = seq
    # 手写编号用到的计数器（emit 阶段就地递增）
    counters = {}
    # ---- 封面节（titlePg：封面不显示页眉页脚；正文节页码重新从 1 起）
    if with_cover:
        # 封面内容
        for _ in range(4):
            doc.add_paragraph("")
        tp = doc.add_paragraph()
        r = tp.add_run(title or doc_name)
        r.font.size = Pt(22); r.font.bold = True
        r.font.name = C.FONT_HEAD["hAnsi"]
        r._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
        tp.alignment = 1
        for _ in range(4):
            doc.add_paragraph("")
        for ln in (org, "（工作区交付文件）"):
            p = doc.add_paragraph()
            rr = p.add_run(ln)
            rr.font.size = Pt(14)
            rr.font.name = C.FONT["hAnsi"]
            rr._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
            p.alignment = 1
        # 封面节：套页面参数 + titlePg（首页无页眉页脚）
        F.setup_section(doc.sections[0], cover=True, body=False)
        # 新增正文节，页码从 1 起
        body_sec = doc.add_section()
        F.setup_section(body_sec, cover=False, body=True)
    else:
        F.setup_section(doc.sections[0], cover=False, body=True)

    # ---- 目录
    if with_toc:
        F.add_toc(doc, title="目  录")

    # ---- 正文
    stats = {}
    def bump(k):
        stats[k] = stats.get(k, 0) + 1

    for kind, payload in blocks:
        if kind == "hr":
            continue
        if kind == "p":
            add_body(doc, payload); bump("正文缩进")
        elif kind == "bullet":
            for item in payload:
                add_body(doc, "· " + item); bump("正文缩进")
        elif kind == "ordered":
            for item in payload:
                add_body(doc, item); bump("正文缩进")
        elif kind == "quote":
            qtext = remap_section_refs("".join(payload), heading_ordinals)
            add_body(doc, qtext, style="006备注", font=C.FONT, sz=21); bump("006备注")
        elif kind == "code":
            for ln in payload:
                p = doc.add_paragraph(style=F.get_style(doc, "正文缩进"))
                r = p.add_run(ln if ln.strip() else "")
                r.font.size = Pt(21)
                r.font.name = CODE_FONT["ascii"]
                r._element.rPr.rFonts.set(qn("w:eastAsia"), CODE_FONT["east_asia"])
                bump("代码行")
        elif kind.startswith("h"):
            lvl = int(kind[1])
            style = LEVEL_MAP.get(lvl)
            # H1 是文档标题：已在封面写出，正文里不再重复（也不进编号体系）。
            # 只有没做封面时才用它兜底成「标题 1」。
            if style is None:
                if with_cover:
                    continue
                style = "标题 1"
            # 默认手写编号：编号直接拼进标题文字（第X节 / X、 / （X））。
            text, ordinal = build_heading_text(payload, lvl, counters)
            p = doc.add_paragraph(style=F.get_style(doc, style))
            add_runs(p, text)
            bump(style)
        elif kind == "table":
            add_table(doc, payload, pattern=pattern); bump("表格")

    # ---- 节与页眉页脚
    # 封面节：titlePg + 起始页码 0，**不写页眉页脚**（封面不显示页码与页眉）。
    # 注意：给 titlePg 的封面节写 first_page_footer 的 PAGE 域会让封面印出"0"，
    # 所以封面节一个 header/footer 部件都不建。
    # 正文节：起始页码 1，页眉 + 三种页脚（含页码）。
    secs = doc.sections
    if with_cover:
        F.setup_section(secs[0], cover=True, body=False)

    body_secs = secs[1:] if with_cover else secs
    for sec in body_secs:
        F.setup_section(sec, cover=False, body=True)
        sp = sec._sectPr
        # 清掉从封面节复制来的 titlePg：正文节首页也要显示页码
        tpg = sp.find(qn("w:titlePg"))
        if tpg is not None:
            sp.remove(tpg)
        for ftr in (sec.footer, sec.even_page_footer, sec.first_page_footer):
            F.write_page_footer(doc, ftr, prefix="")
        F.set_header(doc, sec, "%s  %s" % (org, doc_name))

    doc.save(dst)
    return stats


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--title")
    ap.add_argument("--org", default="XX股份有限公司")
    ap.add_argument("--pattern", default="boxed",
                    choices=["boxed", "threeline", "blueprint"],
                    help="表格线型；默认 boxed（表格美颜② 四边线表）")
    ap.add_argument("--auto-numbering", dest="auto_numbering", action="store_true",
                    help="明确要求时才用自动多级编号；默认序号手写在标题文字里")
    ap.add_argument("--no-cover", action="store_true")
    ap.add_argument("--no-toc", action="store_true")
    a = ap.parse_args()

    stats = convert(a.src, a.dst, title=a.title, org=a.org, pattern=a.pattern,
                    with_cover=not a.no_cover, with_toc=not a.no_toc,
                    auto_numbering=a.auto_numbering)
    print("转换统计:", stats)
    ok, _ = F.check(a.dst, profile="招股说明书",
                    auto_numbering=a.auto_numbering, pattern=a.pattern)
    print("自检:", "PASS" if ok else "FAIL（详见 filing_docx.py check）")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
