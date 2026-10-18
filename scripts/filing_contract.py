# -*- coding: utf-8 -*-
"""
filing_contract.py — 投行（IBD Tools by SyH）Word 排版标准合同（单一可信源）

来源（已逐项核验）：
  1) IBD Tools 插件模板 normal.dotm（作者：帅远华，Ver20240129）
     - word/vbaProject.bin 中的宏：NewTableStyle / newtablestyle2 / newtablestyle3 /
       formatstyle（"设置为招股说明书常用样式"）/ formatfeedbackstyle / 标题1..7 / 正文 /
       表格前单位 / 表格后说明 / NEEQ正文 / NEEQ表格字体 / 一键清理 / 表格重整
     - word/numbering.xml：多级列表 第%1节 / %2、 / （%3） / %4、 / （%5） / %6） / ①
     - word/styles.xml + word/glossary/document.xml（构建基块 表格前单位 / 表格后说明 / 报告目录）
  2) 实例文件（应用该插件的真实申报文件）：
     - 3-2 上市保荐书.docx（001/002/003/004/005/006/009/011 + TOC1/TOC2 家族）
     - 3-1 证券发行保荐书.docx（标题 1..7 + 001..011 + toc 1/2 家族）
  插件与实例冲突时：以**插件宏**为准（用户声明"实例偶有几处不一致，但插件总体一定是对的"）。

单位换算（OOXML）：
  sz/半磅：24=12pt(小四) 21=10.5pt(五号) 18=9pt(小五) 28=14pt(四号) 32=16pt(三号) 44=22pt(二号)
  dxa/二十分之一磅：1440=2.54cm 1800=3.17cm 108=0.19cm 397=0.7cm 284=0.5cm 480=2字符(12pt)
  线宽 sz/八分之一磅：4=0.5pt 12=1.5pt 24=3pt
  底纹 fill：RRGGBB。插件 表格美颜① 表头底纹 WdColor=-603923969 → #FFD900
"""

# ---------------------------------------------------------------- 版式常量
PAGE = {
    "page_w": 11906, "page_h": 16838,          # A4 纵向
    "margin_top": 1440, "margin_bottom": 1440,  # 2.54cm
    "margin_left": 1800, "margin_right": 1800,  # 3.17cm（实例亦有 1797）
    "header_dist": 851, "footer_dist": 992,     # 1.5cm / 1.75cm
    "gutter": 0,
    "doc_grid_line_pitch": 312,                 # 每页行网格（12pt 正文）
    "cover_page_start": 0,                      # 封面节起始页码 0
    "body_page_start": 1,                       # 正文节起始页码 1
    "cover_title_page": True,                   # 封面节 titlePg（首页无页眉）
}
LANDSCAPE_PAGE = dict(PAGE, page_w=16838, page_h=11906, margin_top=1800,
                      margin_bottom=1800, margin_left=1440, margin_right=1440,
                      orient="landscape")

FONT = {"east_asia": "宋体", "ascii": "Times New Roman", "hAnsi": "Times New Roman"}
FONT_HEAD = {"east_asia": "黑体", "ascii": "Times New Roman", "hAnsi": "Times New Roman"}

# ---------------------------------------------------------------- 段落样式合同
# key = styleId；name = 中文样式名（Word 样式窗格与插件按钮实际使用的名字）
# 字段：type/name/based_on/next/link/font/sz/szCs/bold/jc/first_line_chars/first_line/
#      left_chars/left/before_lines/after_lines/line/keep_next/keep_lines/page_break/
#      outline/widow/auto_update
PARA_STYLES = [
    # --- 基础
    {"id": "Normal", "name": "Normal", "type": "paragraph", "default": True,
     "font": FONT, "sz": 24, "szCs": 21, "kern": 2, "jc": "both", "widow": False,
     "next": None, "based_on": None},
    # --- 申报文件正文族（3-2 实际使用）
    # 【实测修复】大纲级别 0/0/1/2/3。
    # 原值 0/1/2/3 与本技能自己的说明文档（references/plugin-and-samples.md
    # “001 大纲级别”一行：「合同按 0/0/1/2 层次的合理值（002=0 与插件一致，
    # 003=1、004=2）」）以及实证文件 3-2 上市保荐书.docx 的实际值（0/0/1/2/3）
    # 相矛盾。目录域是 `TOC \o "1-2" \h \z \u`（只收大纲 0 与 1），按 0/1/2/3
    # 会让 003 落到 outlineLvl=2、三级标题**全部被目录排除**：
    # 实测 3-2 的目录只剩 12 条（实证文件为 12+24=36 条），漏掉全部 24 个三级标题。
    {"id": "001", "name": "001一级标题", "type": "paragraph", "font": FONT_HEAD, "sz": 32,
     "szCs": 28, "bold": True, "jc": "center", "before_lines": 50, "after_lines": 50,
     "line": 360, "keep_next": True, "keep_lines": True, "page_break": True,
     "outline": 0, "based_on": "Normal", "custom": True},
    {"id": "002", "name": "002二级标题", "type": "paragraph", "font": FONT_HEAD, "sz": 28,
     "szCs": 28, "bold": True, "jc": "both", "first_line_chars": 200, "first_line": 480,
     "before_lines": 50, "after_lines": 50, "line": 360, "keep_next": True,
     "keep_lines": True, "outline": 0, "custom": True},
    {"id": "003", "name": "003三级标题", "type": "paragraph", "font": FONT_HEAD, "sz": 24,
     "szCs": 28, "bold": True, "jc": "both", "first_line_chars": 200, "first_line": 480,
     "before_lines": 50, "after_lines": 50, "line": 360, "keep_next": True,
     "keep_lines": True, "outline": 1, "link": "003Char", "custom": True},
    {"id": "003Char", "name": "003三级标题 Char", "type": "character", "font": FONT_HEAD,
     "szCs": 28, "bold": True, "link": "003", "custom": True},
    {"id": "004", "name": "004四级标题", "type": "paragraph", "font": FONT, "sz": 24,
     "szCs": 22, "bold": True, "jc": "both", "first_line_chars": 200, "first_line": 480,
     "before_lines": 50, "after_lines": 50, "line": 360, "keep_next": True,
     "keep_lines": True, "outline": 2, "based_on": "Normal", "custom": True},
    {"id": "005", "name": "005正文", "type": "paragraph", "font": FONT, "sz": 24,
     "szCs": 22, "jc": "both", "first_line_chars": 200, "first_line": 480,
     "before_lines": 50, "before": 120, "line": 360, "based_on": "Normal",
     "link": "005Char", "custom": True},
    {"id": "005Char", "name": "005正文 Char", "type": "character", "szCs": 22,
     "link": "005", "custom": True},
    {"id": "006", "name": "006备注", "type": "paragraph", "font": FONT, "sz": 21,
     "szCs": 22, "jc": "both", "first_line_chars": 200, "first_line": 420,
     "based_on": "Normal", "custom": True},
    {"id": "009", "name": "009单位", "type": "paragraph", "font": FONT, "sz": 21,
     "szCs": 22, "kern": 2, "jc": "right", "keep_next": True, "custom": True},
    {"id": "011", "name": "011五级标题", "type": "paragraph", "sz": 24, "szCs": 22,
     "jc": "both", "first_line_chars": 200, "first_line": 480,
     "before_lines": 50, "after_lines": 50, "line": 360, "outline": 3,
     "based_on": "005", "custom": True},
    # --- 标准标题族（插件"设置为招股说明书常用样式"写入，带多级编号；3-1 使用）
    {"id": "Heading1", "name": "标题 1", "type": "paragraph", "font": FONT_HEAD, "sz": 32,
     "szCs": 44, "jc": "center", "before_lines": 50, "after_lines": 50, "line": 360,
     "keep_next": True, "keep_lines": True, "page_break": True, "outline": 0,
     "numbering": 0, "next": "正文缩进", "link": "Heading1Char"},
    {"id": "Heading1Char", "name": "标题 1 Char", "type": "character", "font": FONT_HEAD,
     "sz": 32, "szCs": 44, "link": "Heading1"},
    {"id": "Heading2", "name": "标题 2", "type": "paragraph", "font": FONT_HEAD, "sz": 28,
     "szCs": 32, "jc": "both", "before_lines": 50, "after_lines": 0, "line": 360,
     "keep_next": True, "keep_lines": True, "outline": 1, "numbering": 1,
     "next": "正文缩进", "link": "Heading2Char"},
    {"id": "Heading2Char", "name": "标题 2 Char", "type": "character", "font": FONT_HEAD,
     "sz": 28, "szCs": 32, "link": "Heading2"},
    {"id": "Heading3", "name": "标题 3", "type": "paragraph", "font": FONT_HEAD, "sz": 24,
     "szCs": 24, "jc": "both", "first_line_chars": 200, "first_line": 480,
     "before_lines": 50, "after_lines": 0, "line": 360, "keep_next": True,
     "keep_lines": True, "outline": 2, "numbering": 2, "next": "正文缩进",
     "link": "Heading3Char"},
    {"id": "Heading3Char", "name": "标题 3 Char", "type": "character", "font": FONT_HEAD,
     "sz": 24, "szCs": 24, "link": "Heading3"},
    {"id": "Heading4", "name": "标题 4", "type": "paragraph", "font": FONT, "sz": 24,
     "szCs": 28, "bold": True, "jc": "both", "first_line_chars": 200,
     "first_line": 480, "before_lines": 50, "after_lines": 0, "line": 360,
     "keep_next": True, "keep_lines": True, "outline": 3, "numbering": 3,
     "next": "正文缩进"},
    {"id": "Heading5", "name": "标题 5", "type": "paragraph", "font": FONT, "sz": 24,
     "szCs": 28, "jc": "left", "first_line_chars": 200, "first_line": 480,
     "before_lines": 50, "after_lines": 50, "line": 360, "keep_next": True,
     "keep_lines": True, "outline": 4, "numbering": 4, "next": "正文缩进"},
    {"id": "Heading6", "name": "标题 6", "type": "paragraph", "font": FONT, "sz": 24,
     "szCs": 24, "jc": "both", "first_line_chars": 200, "first_line": 480,
     "before_lines": 50, "after_lines": 50, "line": 360, "outline": 5,
     "numbering": 5, "next": "正文缩进"},
    {"id": "Heading7", "name": "标题 7", "type": "paragraph", "font": FONT, "sz": 24,
     "szCs": 24, "jc": "left", "first_line_chars": 200, "first_line": 480,
     "before_lines": 50, "after_lines": 50, "line": 360, "keep_next": True,
     "keep_lines": True, "outline": 6, "numbering": 6, "next": "正文缩进"},
    # --- 通用正文（插件 formatstyle 写入"正文缩进"）
    {"id": "BodyIndent", "name": "正文缩进", "type": "paragraph", "font": FONT, "sz": 24,
     "szCs": 22, "jc": "both", "first_line_chars": 200, "first_line": 480,
     "before_lines": 50, "before": 120, "after_lines": 50, "after": 120, "line": 360,
     "auto_update": True, "next": "正文缩进"},
    # --- 表格附件（构建基块）
    {"id": "TableUnit", "name": "表格前单位", "type": "paragraph", "font": FONT, "sz": 21,
     "szCs": 24, "kern": 2, "jc": "right", "keep_next": True, "custom": True,
     "next": "正文缩进"},
    {"id": "TableNote", "name": "表格后说明", "type": "paragraph", "font": FONT, "sz": 21,
     "szCs": 24, "kern": 2, "jc": "both", "custom": True, "next": "正文缩进"},
    # --- 目录
    {"id": "TOCHeading", "name": "TOC 标题1", "type": "paragraph", "font": FONT_HEAD,
     "sz": 32, "szCs": 32, "jc": "center", "before": 240, "after": 0, "line": 259,
     "outline": 9, "custom": True},
    {"id": "TOC1", "name": "toc 1", "type": "paragraph", "font": FONT, "sz": 24,
     "szCs": 24, "bold": True, "jc": "both", "line": 360, "tab_pos": 8296},
    {"id": "TOC2", "name": "toc 2", "type": "paragraph", "font": FONT, "sz": 24,
     "szCs": 24, "jc": "both", "left_chars": 200, "left": 420, "line": 360},
    # --- 页眉页脚
    {"id": "Header", "name": "页眉", "type": "paragraph", "font": FONT, "sz": 18,
     "szCs": 18, "kern": 2, "jc": "both"},
    {"id": "Footer", "name": "页脚", "type": "paragraph", "font": FONT, "sz": 21,
     "szCs": 21, "kern": 2, "jc": "center"},
]

# 表格样式（三线表族）
TABLE_STYLES = [
    {"id": "NormalTable", "name": "普通表格", "type": "table", "default": True,
     "cell_mar_lr": 108},
    {"id": "ReportTable", "name": "报告表格", "type": "table", "based_on": "NormalTable",
     "sz": 18, "custom": True},
]

# ---------------------------------------------------------------- 表格合同
# 插件提供三套"表格美颜"。**默认一律用「表格美颜② 四边线表」（boxed）**——
# 除非使用者明确指定别的格式（三线表 / 蓝线密集表 / 公告细线格）。
# 这是使用者的固定要求，不要凭审美自行改回三线表。
TABLE_PATTERNS = {
    # 插件 NewTableStyle（表格美颜①）——标准三线表
    "threeline": {
        "label": "表格美颜① 三线表",
        "row_height": 397, "font_sz": 18, "header_shading": None,
        "borders": {
            "top":     {"val": "thinThickSmallGap", "sz": 24, "color": "auto"},
            "bottom":  {"val": "thickThinSmallGap", "sz": 24, "color": "auto"},
            "insideH": {"val": "single", "sz": 4, "color": "auto"},
            "left":    {"val": "none"},
            "right":   {"val": "none"},
            "insideV": {"val": "none"},
        },
    },
    # 插件 newtablestyle2（表格美颜②）——四边框线表；3-1/3-2 实例实际使用
    "boxed": {
        "label": "表格美颜② 四边线表",
        "row_height": 397, "font_sz": 21, "header_shading": None,
        "borders": {
            "top":     {"val": "single", "sz": 12, "color": "auto"},
            "left":    {"val": "single", "sz": 12, "color": "auto"},
            "bottom":  {"val": "single", "sz": 12, "color": "auto"},
            "right":   {"val": "single", "sz": 12, "color": "auto"},
            "insideH": {"val": "single", "sz": 4, "color": "auto"},
            "insideV": {"val": "single", "sz": 4, "color": "auto"},
        },
    },
    # 插件 newtablestyle3（表格美颜·密集）——蓝线框，行高更小
    "blueprint": {
        "label": "表格美颜③ 蓝线密集表",
        "row_height": 284, "font_sz": 21, "header_shading": "FFD900",
        "borders": {
            "top":     {"val": "single", "sz": 4, "color": "4881BD"},
            "left":    {"val": "single", "sz": 4, "color": "4881BD"},
            "bottom":  {"val": "single", "sz": 4, "color": "4881BD"},
            "right":   {"val": "single", "sz": 4, "color": "4881BD"},
            "insideH": {"val": "single", "sz": 4, "color": "4881BD"},
            "insideV": {"val": "single", "sz": 4, "color": "4881BD"},
        },
    },
    # 公告体裁：0.5pt 四边细线格（上交所/深交所临时公告表格惯例）
    "announcement": {
        "label": "公告表格 0.5pt 细线格",
        "row_height": 397, "font_sz": 24, "header_shading": None,
        "borders": {
            "top":     {"val": "single", "sz": 4, "color": "auto"},
            "left":    {"val": "single", "sz": 4, "color": "auto"},
            "bottom":  {"val": "single", "sz": 4, "color": "auto"},
            "right":   {"val": "single", "sz": 4, "color": "auto"},
            "insideH": {"val": "single", "sz": 4, "color": "auto"},
            "insideV": {"val": "single", "sz": 4, "color": "auto"},
        },
    },
}

TABLE = {
    "width_pct": 5000,        # 相对窗口 100%
    "align": "center",        # 表格居中
    "pad_top": 0, "pad_bottom": 0, "pad_lr": 108,   # 上下 0 / 左右 0.19cm
    "allow_break_across_pages": False,              # 不允许跨页断行（短表）
    "heading_format": True,   # 表头行重复（tblHeader）
    # 【硬性规则】默认用 IBD Tools「表格美颜② 四边线表」。
    # 只有使用者明确要求时才换成 threeline（三线表）/ blueprint（蓝线密集表）。
    "default_pattern": "boxed",
    "header_bold": True,
    "header_jc": "center",
    "header_keep_next": True,
    "v_align": "center",
    "cell_jc_numeric": "right",   # 数字、"-"、带%数字 → 右对齐
    "cell_jc_text": "left",       # 其余 → 左对齐
    "total_row_prefix": "合计",   # 首列以此开头 → 该格居中 + 整行加粗
    "auto_fit": ["content", "window"],
    "cell_font": {"east_asia": "宋体", "ascii": "Times New Roman"},
    "cell_pmargin_lr": 0, "cell_pmargin_tb": 0,
}

# 向后兼容：TABLE["borders"] 指向默认口径
TABLE["borders"] = TABLE_PATTERNS[TABLE["default_pattern"]]["borders"]
TABLE["height_presets"] = {
    p["label"]: {"row_height": p["row_height"], "font_sz": p["font_sz"],
                 "header_shading": p["header_shading"]}
    for p in TABLE_PATTERNS.values()
}
TABLE["default_preset"] = TABLE_PATTERNS[TABLE["default_pattern"]]["label"]

# ---------------------------------------------------------------- 交付口径 Profile
PROFILES = {
    # 3-2 上市保荐书 / 3-1 证券发行保荐书 / 3-4 尽职调查报告 实测口径
    "申报文件": {
        "body_style": "005正文", "comment_style": "006备注", "unit_style": "009单位",
        "headings": ["001一级标题", "002二级标题", "003三级标题", "004四级标题", "011五级标题"],
        "table_pattern": "boxed", "header_sz": 18, "footer_sz": 21,
        "footer_prefix": "{doc_code}-", "heading_scheme": "中文序号手写（一、/（一）/1、）",
        "manual_scheme": "doc_filing", "auto_numbering": False,
    },
    # 通用/招股说明书口径（插件"设置为招股说明书常用样式"）
    # 默认表格用「表格美颜② 四边线表」；编号默认手写（第X节 / X、 / （X））。
    "招股说明书": {
        "body_style": "正文缩进", "comment_style": "正文缩进", "unit_style": "表格前单位",
        "headings": ["标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "标题 6", "标题 7"],
        "table_pattern": "boxed", "header_sz": 18, "footer_sz": 21,
        "heading_scheme": "中文序号手写（第X节 / X、 / （X） / X、）",
        "manual_scheme": "doc_prospectus", "auto_numbering": False,
    },
    # 回复报告口径（插件"设置为回复报告常用样式"）
    "回复报告": {
        "body_style": "正文缩进", "comment_style": "正文缩进", "unit_style": "表格前单位",
        "headings": ["标题 1", "标题 2", "标题 3", "标题 4", "标题 5"],
        "table_pattern": "boxed", "header_sz": 18, "footer_sz": 21,
        "heading_scheme": "中文序号手写（问题X / X、 / （X））",
        "manual_scheme": "doc_reply", "auto_numbering": False,
    },
    # 新三板申报口径（插件 NEEQ 正文 / NEEQ 表格字体）
    "新三板": {
        "body_style": "005正文", "comment_style": "006备注", "unit_style": "009单位",
        "headings": ["001一级标题", "002二级标题", "003三级标题", "004四级标题"],
        "table_pattern": "boxed", "body_sz": 21, "table_font_sz": 21,
        "header_sz": 18, "footer_sz": 21, "heading_scheme": "中文序号手写",
        "manual_scheme": "doc_filing", "auto_numbering": False,
    },
    # 上交所/深交所临时公告体裁（公告格式指引模板）
    # 与申报文件的关键差异：正文用仿宋_GB2312 小三(15pt)、固定行距 28 磅、
    # 全文左对齐（公文惯例，不用两端对齐）。
    # 公告是格式指引模板，表格线由交易所规定为 0.5pt 细线格，属"使用者/规则指定格式"，
    # 不走默认的四边线表。
    "公告": {
        "body_style": "公告正文", "comment_style": "公告注释", "unit_style": "公告注释",
        "headings": ["公告标题", "公告一级", "公告二级", "公告三级"],
        "table_pattern": "announcement", "header_sz": 21, "footer_sz": 21,
        "heading_scheme": "中文序号手写（一、/（一）/1.）",
        "body_sz": 30, "body_font": {"east_asia": "仿宋_GB2312", "ascii": "仿宋_GB2312",
                                     "hAnsi": "仿宋GB-2312"},
        "heading_font": {"east_asia": "黑体", "ascii": "黑体", "hAnsi": "黑体"},
        "table_font_sz": 24, "line_exact": 560, "align": "left",
        "note_sz": 24,
        "manual_scheme": "doc_notice", "auto_numbering": False,
    },
    # 交易所自律监管指南 / 业务办理指南 / 格式指引（GB/T 9704 公文体例）
    # 与"公告"的关键差异：正文 14pt（公告 15pt）、行距固定 30 磅（公告 28 磅）、
    # 一级标题为加粗仿宋（公告为黑体）、表内文字 10pt（公告 12pt）。
    "公文": {        "body_style": "公文正文", "comment_style": "公文注释", "unit_style": "公文注释",
        "headings": ["文件标题", "公文一级", "公文二级"],
        "table_pattern": "announcement", "header_sz": 21, "footer_sz": 21,
        "heading_scheme": "中文序号手写（一、/1.）",
        "body_sz": 28, "body_font": {"east_asia": "仿宋_GB2312", "ascii": "仿宋_GB2312",
                                     "hAnsi": "仿宋GB-2312"},
        "heading_font": {"east_asia": "仿宋_GB2312", "ascii": "仿宋_GB2312",
                         "hAnsi": "仿宋GB-2312"},
        "table_font_sz": 20, "line_exact": 600, "align": "left",
        "note_sz": 24,
        "manual_scheme": "doc_notice", "auto_numbering": False,
    },
}

# ---------------------------------------------------------------- 公告体裁样式（上交所/深交所临时公告）
# 数值取自实证公告模板（第十一号 上市公司非股改限售股上市流通公告）。
# 与 PARAGRAPH_STYLES 分开维护：公告体裁的字体/字号/行距/对齐与申报文件不同，
# 混用会把公告正文压成 12pt 宋体。
ANNOUNCE_STYLES = [
    {"id": "AnnTitle", "name": "公告标题", "type": "paragraph", "custom": True,
     "font": {"east_asia": "黑体", "ascii": "黑体", "hAnsi": "黑体"},
     "sz": 36, "bold": True, "jc": "center", "line": 560, "exact": True,
     "keep_next": True, "keep_lines": True, "outline": 0, "next": "公告正文"},
    {"id": "AnnSubject", "name": "公告文种标题", "type": "paragraph", "custom": True,
     "font": {"east_asia": "仿宋_GB2312", "ascii": "仿宋_GB2312", "hAnsi": "仿宋GB-2312"},
     "sz": 30, "jc": "center", "line": 560, "exact": True,
     "keep_next": True, "keep_lines": True, "outline": 0, "next": "公告正文"},
    {"id": "AnnH1", "name": "公告一级", "type": "paragraph", "custom": True,
     "font": {"east_asia": "黑体", "ascii": "黑体", "hAnsi": "黑体"},
     "sz": 30, "jc": "left", "first_line_chars": 200, "first_line": 600,
     "line": 560, "exact": True, "keep_next": True, "keep_lines": True,
     "outline": 1, "next": "公告正文"},
    {"id": "AnnH2", "name": "公告二级", "type": "paragraph", "custom": True,
     "font": {"east_asia": "仿宋_GB2312", "ascii": "仿宋_GB2312", "hAnsi": "仿宋GB-2312"},
     "sz": 30, "jc": "left", "first_line_chars": 200, "first_line": 600,
     "line": 560, "exact": True, "keep_next": True, "keep_lines": True,
     "outline": 2, "next": "公告正文"},
    {"id": "AnnH3", "name": "公告三级", "type": "paragraph", "custom": True,
     "font": {"east_asia": "仿宋_GB2312", "ascii": "仿宋_GB2312", "hAnsi": "仿宋GB-2312"},
     "sz": 30, "jc": "left", "first_line_chars": 200, "first_line": 600,
     "line": 560, "exact": True, "keep_next": True, "keep_lines": True,
     "outline": 3, "next": "公告正文"},
    # 标签行（"适用情形：""重要内容提示：""上网公告文件""报备文件"）——
    # 实证公告中这些行是**加粗仿宋**，不是黑体，也不是普通正文
    {"id": "AnnLabel", "name": "公告标签", "type": "paragraph", "custom": True,
     "font": {"east_asia": "仿宋_GB2312", "ascii": "仿宋_GB2312", "hAnsi": "仿宋GB-2312"},
     "sz": 30, "bold": True, "jc": "left", "first_line_chars": 200, "first_line": 600,
     "line": 560, "exact": True, "keep_next": True, "next": "公告正文"},
    {"id": "AnnBody", "name": "公告正文", "type": "paragraph", "custom": True,
     "font": {"east_asia": "仿宋_GB2312", "ascii": "仿宋_GB2312", "hAnsi": "仿宋GB-2312"},
     "sz": 30, "jc": "left", "first_line_chars": 200, "first_line": 600,
     "line": 560, "exact": True, "next": "公告正文"},
    {"id": "AnnList", "name": "公告要点", "type": "paragraph", "custom": True,
     "font": {"east_asia": "仿宋_GB2312", "ascii": "仿宋_GB2312", "hAnsi": "仿宋GB-2312"},
     "sz": 30, "jc": "left", "left": 1020, "hanging": 420,
     "line": 560, "exact": True, "next": "公告正文"},
    {"id": "AnnNote", "name": "公告注释", "type": "paragraph", "custom": True,
     "font": {"east_asia": "仿宋_GB2312", "ascii": "仿宋_GB2312", "hAnsi": "仿宋GB-2312"},
     "sz": 24, "jc": "left", "first_line_chars": 200, "first_line": 480,
     "line": 360, "exact": False, "next": "公告注释"},
    {"id": "AnnSection", "name": "公告分项", "type": "paragraph", "custom": True,
     "font": {"east_asia": "黑体", "ascii": "黑体", "hAnsi": "黑体"},
     "sz": 30, "bold": True, "jc": "left", "first_line_chars": 200, "first_line": 600,
     "line": 560, "exact": True, "keep_next": True, "keep_lines": True,
     "outline": 1, "next": "公告正文"},
    {"id": "AnnSign", "name": "公告落款", "type": "paragraph", "custom": True,
     "font": {"east_asia": "仿宋_GB2312", "ascii": "仿宋_GB2312", "hAnsi": "仿宋GB-2312"},
     "sz": 30, "jc": "right", "first_line_chars": 200, "first_line": 600,
     "line": 560, "exact": True, "next": "公告正文"},
]

# ---------------------------------------------------------------- 公文体裁样式
# 交易所自律监管指南 / 业务办理指南 / 格式指引正文（GB/T 9704 系公文体例）
# 数值取自实证文件《第一号 信息披露业务办理（2025年4月修订）》。
# 与"公告"体裁的三点差异（混用会串体）：
#   · 正文 14pt（公告 15pt）、行距固定 30 磅（公告 28 磅）
#   · 一级标题是**加粗仿宋同字号**（公告是黑体）——公文体例的典型写法
#   · 表内文字 10pt（公告 12pt）；全文只有文件标题用黑体
GUIDE_FONT = {"east_asia": "仿宋_GB2312", "ascii": "仿宋_GB2312", "hAnsi": "仿宋GB-2312"}
GUIDE_STYLES = [
    # 文件标题（"第一号 信息披露业务办理"）：黑体 20pt 居中
    {"id": "GuideTitle", "name": "文件标题", "type": "paragraph", "custom": True,
     "font": {"east_asia": "黑体", "ascii": "黑体", "hAnsi": "黑体"},
     "sz": 40, "jc": "center", "line": 600, "exact": True,
     "keep_next": True, "keep_lines": True, "outline": 0, "next": "公文正文"},
    # 一级标题（一、二、三、）：加粗仿宋 14pt
    {"id": "GuideH1", "name": "公文一级", "type": "paragraph", "custom": True,
     "font": GUIDE_FONT, "sz": 28, "bold": True, "jc": "left",
     "first_line_chars": 200, "first_line": 560,
     "line": 600, "exact": True, "keep_next": True, "keep_lines": True,
     "outline": 1, "next": "公文正文"},
    # 二级标题（1、公告标题 / 2、公告正文）：同字号不加粗，仅给大纲级别与 keepNext
    {"id": "GuideH2", "name": "公文二级", "type": "paragraph", "custom": True,
     "font": GUIDE_FONT, "sz": 28, "jc": "left",
     "first_line_chars": 200, "first_line": 560,
     "line": 600, "exact": True, "keep_next": True, "keep_lines": True,
     "outline": 2, "next": "公文正文"},
    # 正文（含 1. / （1） 等条目）：仿宋 14pt 首行缩进 2 字符 固定行距 30 磅
    {"id": "GuideBody", "name": "公文正文", "type": "paragraph", "custom": True,
     "font": GUIDE_FONT, "sz": 28, "jc": "left",
     "first_line_chars": 200, "first_line": 560,
     "line": 600, "exact": True, "next": "公文正文"},
    # 附件标签（附件1：/附件2：）：加粗仿宋
    {"id": "GuideAnnex", "name": "附件标签", "type": "paragraph", "custom": True,
     "font": GUIDE_FONT, "sz": 28, "bold": True, "jc": "left",
     "first_line_chars": 200, "first_line": 560,
     "line": 600, "exact": True, "keep_next": True, "next": "公文正文"},
    # 附件名（"上海证券交易所信息披露公告类别索引"）：居中仿宋
    {"id": "GuideCaption", "name": "附件标题", "type": "paragraph", "custom": True,
     "font": GUIDE_FONT, "sz": 28, "jc": "center", "line": 600, "exact": True,
     "keep_next": True, "keep_lines": True, "outline": 1, "next": "公文正文"},
    # 表内文字：仿宋 10pt 单倍行距
    {"id": "GuideCell", "name": "公文表内", "type": "paragraph", "custom": True,
     "font": GUIDE_FONT, "sz": 20, "jc": "left", "line": 240, "exact": False},
    # 表下注释：仿宋 12pt
    {"id": "GuideNote", "name": "公文注释", "type": "paragraph", "custom": True,
     "font": GUIDE_FONT, "sz": 24, "jc": "left",
     "first_line_chars": 200, "first_line": 480,
     "line": 360, "exact": True, "next": "公文正文"},
]

# ---------------------------------------------------------------- 编号合同（插件 numbering.xml）
# 【硬性规则】默认**绝不使用自动编号**，全部用"手写在标题文字里"的编号。
# 只有在使用者明确要求"用自动编号/多级列表"时才启用自动编号（AUTO_NUMBERING_DEFAULT=True）。
# 理由：自动编号在复制、拆分、合并文档时会错位（序号跟着段落走，跨文档不认原编号），
# 投行文件要的是"看得见、改得动、搬得走"的编号。
AUTO_NUMBERING_DEFAULT = False

# 各体裁的"手写编号"方案：level -> 格式模板（%s = 该级序号，中文由 cn_num 生成）
# 注意模板要带分隔符：一级用空格（"第一节 标题"），二至四级用顿号/括号（"一、标题"）。
# 少了空格会写成"第一节可复制简报"，与公文惯例不符。
MANUAL_NUMBER_SCHEME = {
    "doc_prospectus": {          # 招股说明书：第X节 / X、 / （X） / X、
        1: "第%s节 ",
        2: "%s、",
        3: "（%s）",
        4: "%s、",
    },
    "doc_reply": {               # 回复报告：问题X / X、
        1: "问题%s ",
        2: "%s、",
        3: "（%s）",
        4: "%s、",
    },
    # 申报文件（保荐书系）：一级标题独起一节，二级用手写中文序号，
    # 三/四级沿用「（X）」「X、」（实证 3-1/3-2 的做法）
    "doc_filing": {
        1: "",
        2: "%s、",
        3: "（%s）",
        4: "%s、",
    },
    # 公告/公文：一、/（一）/1. 手写
    "doc_notice": {
        1: "%s、",
        2: "（%s）",
        3: "%s.",
        4: "（%s）",
    },
}

# 自动编号合同（仅在使用者明确要求时使用）
NUMBERING = [
    # ilvl, numFmt, lvlText, styleId, suff
    {"ilvl": 0, "fmt": "chineseCounting", "text": "第%1节", "style": "Heading1", "suff": "space"},
    {"ilvl": 1, "fmt": "chineseCounting", "text": "%2、", "style": "Heading2", "suff": "nothing"},
    {"ilvl": 2, "fmt": "chineseCounting", "text": "（%3）", "style": "Heading3", "suff": "nothing"},
    {"ilvl": 3, "fmt": "decimal", "text": "%4、", "style": "Heading4", "suff": "nothing"},
    {"ilvl": 4, "fmt": "decimal", "text": "（%5）", "style": "Heading5", "suff": "nothing"},
    {"ilvl": 5, "fmt": "decimal", "text": "%6）", "style": "Heading6", "suff": "nothing"},
    {"ilvl": 6, "fmt": "decimalEnclosedCircle", "text": "%7", "style": "Heading7", "suff": "nothing"},
]
NUMBERING_REPLY = NUMBERING[:2]   # 回复报告：问题%1 / %2

CN_NUM = "一二三四五六七八九十"


def cn_num(n):
    """1..99 -> 一 / 二 / … / 十一 / 二十 / 二十一（手写编号用）。"""
    if n <= 0:
        return ""
    if n <= 10:
        return CN_NUM[n - 1]
    if n < 20:
        return "十" + CN_NUM[n - 11]
    tens, ones = divmod(n, 10)
    return CN_NUM[tens - 1] + "十" + (CN_NUM[ones - 1] if ones else "")


def manual_number(scheme, level, ordinal):
    """按手写编号方案生成某级序号文本（如 第X节 / X、 / （X））。

    scheme: MANUAL_NUMBER_SCHEME 的键，或直接给 dict
    level:  1..7
    ordinal: 该级的出现次序（1 起）
    """
    m = MANUAL_NUMBER_SCHEME.get(scheme) if isinstance(scheme, str) else scheme
    if not m:
        return ""
    tpl = m.get(level)
    if not tpl:
        return ""
    return tpl % cn_num(ordinal)

# 公告体裁编号：公告正文里的序号是手写在文字里的（一、/（一）/1.），
# 只有"重要内容提示"的要点行用真正的列表编号。
NUMBERING_ANNOUNCE = [
    {"ilvl": 0, "fmt": "decimal", "text": "%1.", "style": "AnnList", "suff": "space"},
]

# ---------------------------------------------------------------- 页眉页脚合同
HEADER_FOOTER = {
    "header_sz": 18,          # 9pt
    "header_align": "both",
    "header_bottom_border": {"val": "single", "sz": 4, "space": 1, "color": "auto"},
    "footer_sz": 21,          # 10.5pt
    "footer_align": "center",
    "footer_prefix_template": "{doc_code}-",   # 如 "3-2-"
    "footer_page_field": "PAGE   \\* MERGEFORMAT",
    "toc_field": 'TOC \\o "1-2" \\h \\z \\u',
}

# ---------------------------------------------------------------- 文字与数字规则
TEXT_RULES = {
    "chinese_punctuation": ["，", "。", "；", "：", "？", "！", "—", "～", "（", "）", "《", "》", "“", "”"],
    "english_to_chinese": {",": "，", ".": "。", ";": "；", ":": "：", "?": "？", "!": "！",
                          "-": "—", "~": "～", "(": "（", ")": "）", "<": "《", ">": "》"},
    "numeric_punctuation_keep_halfwidth": True,   # 数字之间保留半角 , 与 .
    "thousands_separator_format": "#,##0.00;-#,##0.00; ",   # 千分位、两位小数、0 显示空白
    "percent_format": "0.00%",
    "wipe_to_wan_format": "#,##0.00;-#,##0.00; ",
    "wipe_to_wan_divisor": 10000,
    "wipe_to_wan_min": 0.01,     # 不足 0.01 显示 "-"
    "unit_line_template": "单位：{unit}",       # 单位：万元 / 单位：元 / 单位：股
    "note_line_prefix": "注：",
    "circled_numbers": ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"],
    "forbidden_fonts_outside_quote": ["楷体"],   # 楷体仅用于需要提示/待核对的局部
}

# 一键清理等价清单（插件 一键清理 = 定稿前必做）
FINAL_CLEAN_CHECKLIST = [
    "接受文档中所有修订并关闭修订模式",
    "删除文档中所有批注",
    "刷新目录页码（TablesOfContents(1).UpdatePageNumbers）",
    "清除全文文字突出显示（HighlightColorIndex = wdNoHighlight）",
    "全文文字颜色设为黑色（Font.ColorIndex = wdBlack）",
    "标题、表格外无残留楷体（楷体转宋体+TNR）",
    "页眉页脚与正文页码连续、无残留域",
]

# ---------------------------------------------------------------- 宏名映射
# 插件 Ribbon 按钮 → VBA 宏名（供 COM 路线直接调用；Application.Run("宏名")）
MACRO_MAP = {
    "正文": "正文", "一级标题": "标题1", "二级标题": "标题2", "三级标题": "标题3",
    "四级标题": "标题4", "五级标题": "标题5", "六级标题": "标题6",
    "表格前单位": "表格前单位", "表格后说明": "表格后说明",
    "表格美颜①": "newtablestyleribbon", "表格美颜②": "newtablestyle2ribbon",
    "粘贴并套用表格样式": "pastenewtablestyle",
    "全文表格美颜①": "allnewtablestyle", "全文表格美颜②": "allnewtablestyle2",
    "首行重复": "firstlinerepeat", "与下同页": "keepwithnext", "表格重整": "表格重整",
    "空单元格处理": "beautifulbiao",
    "单位为万元": "Unit10000", "表下加注": "explain",
    "千分位": "qianfen", "除以一万": "DivIvan", "加百分号": "insertpercent",
    "小数位": ["round1", "round2", "round3", "round4", "round5", "round6", "round7", "round8", "round9", "round10"],
    "自动计算占比": "getpercent", "选行校验": "RowCheck", "选列校验": "ColumnCheck",
    "一键取和": "getsum",
    "设置为招股说明书常用样式": "formatstyle",
    "设置为回复报告常用样式": "formatfeedbackstyle",
    "本章样式": "importstyles", "关闭样式自动更新": "disablestyleupdate",
    "插入目录": "insertcatalog", "插入横向页": "Insertlandscape", "添加页码": "addpagenumber",
    "一键清理": "onekeyclean", "楷体转宋体": "kaititosong",
    "列表转文本": "convertlisttotext", "邮件合并": "mymailmerge",
    "合并批注": "合并批注", "清洗PDF": "清洗pdf",
    "智能粘贴": "pasteastext", "删除空格": "DelblankRibbon", "空白为无": "Delblank2ribbon",
}

STYLE_ALIASES = {
    "一级标题": "001一级标题", "二级标题": "002二级标题", "三级标题": "003三级标题",
    "四级标题": "004四级标题", "五级标题": "011五级标题", "正文": "005正文",
    "备注": "006备注", "单位": "009单位",
    "标题1": "标题 1", "标题2": "标题 2", "标题3": "标题 3",
    "标题4": "标题 4", "标题5": "标题 5", "标题6": "标题 6", "标题7": "标题 7",
    "目录标题": "TOC 标题1", "目录1": "toc 1", "目录2": "toc 2",
    # 公告体裁
    # 标签行（"适用情形：""重要内容提示：""上网公告文件""报备文件"）——
    # 实证公告中这些行是加粗仿宋，既不是黑体标题也不是普通正文
    "公告标签": "公告标签",
    "公告标题": "公告标题", "文种标题": "公告文种标题", "公告一级": "公告一级",
    "公告二级": "公告二级", "公告三级": "公告三级", "公告正文": "公告正文",
    "公告要点": "公告要点", "公告注释": "公告注释", "公告落款": "公告落款",
    # 公文体裁
    "文件标题": "文件标题", "公文一级": "公文一级", "公文二级": "公文二级",
    "公文正文": "公文正文", "公文表内": "公文表内", "公文注释": "公文注释",
    "附件标签": "附件标签", "附件标题": "附件标题",
}


def guide_styles():
    """公文体裁的全部段落样式。"""
    return GUIDE_STYLES


def announce_styles():
    """公告体裁的全部段落样式（与 PARAGRAPH_STYLES 合并使用）。"""
    return ANNOUNCE_STYLES


def all_styles():
    """全部体裁的样式（申报文件 + 公告 + 公文）。"""
    return PARA_STYLES + TABLE_STYLES + ANNOUNCE_STYLES + GUIDE_STYLES


def style_by_name(name):
    target = STYLE_ALIASES.get(name, name)
    for s in all_styles():
        if s["name"] == target or s["id"] == target:
            return s
    raise KeyError("unknown style: %s" % name)


def is_numeric_cell(text):
    """插件判定：rPr 的 IsNumeric(Acell) Or Acell="-" Or IsNumeric(Replace(Acell,"%",""))"""
    t = (text or "").strip()
    if t == "":
        return False
    if t == "-":
        return True
    t2 = t.replace(",", "").replace(" ", "").replace("％", "%")
    if t2.endswith("%"):
        t2 = t2[:-1]
    try:
        float(t2)
        return True
    except ValueError:
        return False
