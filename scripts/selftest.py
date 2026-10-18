# -*- coding: utf-8 -*-
"""End-to-end self test: build a filing-style docx from the contract, then check it.

输出写入临时目录（可由 --out 指定），不污染工作区。
"""
import os, sys, zipfile, subprocess, json, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx import Document
from docx.shared import Pt
import filing_docx as F
import filing_contract as C

OUT = os.environ.get("FILING_SELFTEST_OUT") or os.path.join(
    tempfile.mkdtemp(prefix="filing-selftest-"), "e2e.docx")


def build():
    if os.path.exists(OUT):
        os.remove(OUT)
    doc = Document()
    # 默认手写编号（规则 1）：不注入 numbering.xml，样式不挂 numPr。
    F.write_styles(doc, auto_numbering=C.AUTO_NUMBERING_DEFAULT)
    if C.AUTO_NUMBERING_DEFAULT:
        F.write_numbering(doc)
    else:
        F.drop_numbering(doc)
    F.setup_section(doc.sections[0], cover=False, body=True)
    F.set_header(doc, doc.sections[0], "示例公司                                    2025年度向特定对象发行股票之上市保荐书")
    F.set_footer(doc, doc.sections[0], "3-2-")
    doc.add_paragraph("释  义", style=F.get_style(doc, "001"))
    doc.add_paragraph("在本上市保荐书中，除非另有说明，下列词语具有如下特定含义：",
                      style=F.get_style(doc, "005"))
    doc.add_paragraph("一、发行人基本情况", style=F.get_style(doc, "002"))
    doc.add_paragraph("（一）发行人概况", style=F.get_style(doc, "003"))
    doc.add_paragraph("1、合并资产负债表主要数据", style=F.get_style(doc, "004"))
    doc.add_paragraph("单位：万元", style=F.get_style(doc, "009"))
    t = doc.add_table(rows=4, cols=3)
    t.style = F.get_style(doc, "ReportTable")
    data = [["项目", "2025.06.30", "2024.12.31"],
            ["资产总额", "813,242.82", "772,096.53"],
            ["负债总额", "-4,484.36", "462,123.10"],
            ["合计", "808,758.46", "1,234,219.63"]]
    for r, row in enumerate(data):
        for c, v in enumerate(row):
            t.cell(r, c).text = v
    F.format_table(t, "boxed")
    doc.add_paragraph("注：本表数据经审计。", style=F.get_style(doc, "006"))
    doc.save(OUT)
    return OUT


ok_build = os.path.exists(build())
ok, findings = F.check(OUT)
print("build:", ok_build)
print("check:", "PASS" if ok else "FAIL")
for f in findings:
    print("  ", f["level"], f["area"], f["rule"], "|", f["expected"], "->", f["actual"])

# structural assertions
z = zipfile.ZipFile(OUT)
names = z.namelist()
need = ["word/document.xml", "word/styles.xml",
        "word/header1.xml", "word/footer1.xml",
        "[Content_Types].xml", "word/_rels/document.xml.rels"]
missing = [n for n in need if n not in names]

# 规则 1：手写编号口径下**不应**出现 numbering.xml 或任何 numPr
if not C.AUTO_NUMBERING_DEFAULT:
    if "word/numbering.xml" in names:
        missing.append("手写编号口径不应有 word/numbering.xml")
    if "<w:numPr>" in z.read("word/document.xml").decode("utf-8"):
        missing.append("手写编号口径不应有正文 numPr")
    if "<w:numPr>" in z.read("word/styles.xml").decode("utf-8"):
        missing.append("手写编号口径不应有样式 numPr")

# 内容类型与关系声明完整性（Word 打开"文件损坏"的常见根因）
ct = z.read("[Content_Types].xml").decode("utf-8")
for part in ("word/document.xml", "word/styles.xml"):
    if "/" + part not in ct:
        missing.append("Content_Types 缺少 " + part)
rel = z.read("word/_rels/document.xml.rels").decode("utf-8")
if "styles.xml" not in rel:
    missing.append("document.xml.rels 缺少 styles.xml")
print("missing parts:", missing)
print("artifact:", OUT)
sys.exit(0 if (ok and ok_build and not missing) else 1)
