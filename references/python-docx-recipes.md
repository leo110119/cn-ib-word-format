# python-docx 落地投行格式的可用代码

以下片段均已在本机实测（python-docx 1.2.0 / lxml 5.3.0 / Python 3.11）。先 `sys.path.insert(0, "<skill>/scripts")`，再 `import filing_docx as F, filing_contract as C`。

## 0. 取样式：必须用 get_style，不要用 doc.styles["文名"]

python-docx 的内置样式按**英文名**索引，中文样式名直接索引会 KeyError 或映射到错误对象。

```python
import filing_docx as F
s = F.get_style(doc, "001")        # styleId
s = F.get_style(doc, "001一级标题")  # 中文名，等价
s = F.get_style(doc, "005正文")
```

`get_style` 遍历 `doc.styles` 同时比对 `style_id` 与 `name`。旧名与新名（如 `003 三级标题` 与 `003三级标题`）不通用——**以本技能生成的样式名为准**。

## 1. 从零建一份申报文件

```python
from docx import Document
from docx.shared import Pt
import filing_docx as F

doc = F.new_document(profile="申报文件", auto_numbering=False)  # 默认手写编号
F.setup_section(doc.sections[0], cover=False, body=True)
F.set_header(doc, doc.sections[0], "发行人全称                      2025年度向特定对象发行股票之上市保荐书")
F.set_footer(doc, doc.sections[0], "3-2-")
doc.save("out.docx")
```

封面节另开分节：`F.setup_section(sec, cover=True)`（`pgNumType start=0` + `titlePg`）。

## 2. 写正文（只挂样式，不逐段设字体）

```python
doc.add_paragraph("释  义", style=F.get_style(doc, "001"))
p = doc.add_paragraph("正文内容……", style=F.get_style(doc, "005"))
p.paragraph_format.first_line_indent = Pt(24)      # 仅在需要覆盖样式时才写
```

**不要**对每段调用 `run.font.name = "宋体"`——样式已经承载字体，逐 run 设置会污染 XML 并让 check 报大量偏差。

## 3. 表格：加样式 + 美颜 + 数字对齐

```python
t = doc.add_table(rows=n, cols=m)
t.style = F.get_style(doc, "ReportTable")
for r, row in enumerate(data):
    for c, v in enumerate(row):
        t.cell(r, c).text = str(v)
F.format_table(t, "boxed")     # boxed / threeline / blueprint
```

`format_table` 会写：`tblW 100% pct`、居中、边框、单元格边距、行高 `atLeast 397`、首行 `tblHeader`、表内字体宋体+TNR + 字号、`vAlign center`、数字右对齐/文本左对齐、表内段落缩进清零、`合计`行加粗。幂等，可重复调用。

**合并单元格后必再跑一次** `format_table`——合并会生成新的 `tcPr`，之前设的 `vAlign` 可能丢失。

新增表格前务必重置 `t.style`：`add_table` 默认给 `Table Grid`（四边 0.5pt 全框），不替换会和标准冲突。

## 4. 页眉页脚与目录

```python
F.set_header(doc, doc.sections[0], header_text)
F.set_footer(doc, doc.sections[0], "3-2-")
F.add_toc(doc, "目  录")      # 插 TOC 标题段 + TOC 域，需用 Word 按 F9 或 COM 更新
```

域内容在 python-docx 里靠手工拼 `fldChar` / `instrText`（`F._add_field` 已实现）。**域不会自动求值**，缓存文本随便填；用 Word/COM `Fields.Update()` 或让用户按 F9 刷新。

## 5. 修补既有文件（手写编号原稿）

原稿如使用自动编号，改为局部调整样式属性；下面的整体替换会覆盖样式级编号绑定，不适用于该情况。

```python
from docx import Document
import filing_docx as F
doc = Document("原件.docx")
F.write_styles(doc, auto_numbering=False)  # 同名样式整体替换；不注入新编号
# 既有编号属于可见内容，先保留；不要为套格式清除 numPr/numbering。
for i, sec in enumerate(doc.sections):
    F.setup_section(sec, cover=(i == 0), body=(i > 0))
for t in doc.tables:
    F.format_table(t, "boxed")
doc.save("输出.docx")
```

`write_styles` 的行为：删掉与合同同 `styleId` 的旧样式元素，再按合同顺序追加新的。**不会删除文档里其他样式**（如 `标题 8`、`列表段落`），也**不会改动段落已有的 `pStyle` 引用**——段落仍指向同一个 `styleId`，因此外观自动生效。

## 6. 校验

```python
import filing_docx as F
ok, findings = F.check("输出.docx")
for x in findings:
    print(x["level"], x["area"], x["rule"], x["expected"], "->", x["actual"])
```

命令行等价：`python scripts/filing_docx.py check 输出.docx [--json]`，退出码 0 表示无 FAIL。

## 7. 读取既有文件的真实格式（用于比对）

```python
import zipfile
from lxml import etree
from docx.oxml.ns import qn

z = zipfile.ZipFile("参考文件.docx")
styles = etree.fromstring(z.read("word/styles.xml"))
for st in styles.findall(qn("w:style")):
    sid = st.get(qn("w:styleId"))
    nm = st.find(qn("w:name"))
    rpr = st.find(qn("w:rPr"))
    f = rpr.find(qn("w:rFonts")) if rpr is not None else None
    sz = rpr.find(qn("w:sz")) if rpr is not None else None
    print(sid, nm.get(qn("w:val")),
          f.get(qn("w:eastAsia")) if f is not None else None,
          sz.get(qn("w:val")) if sz is not None else None)
```

`docx.oxml.ns.qn("w:xxx")` 把前缀展开成完整命名空间，是读写 OOXML 的通用做法。

## 8. 用 Word COM 做最终确认（可选，Windows 有 Word 时）

python-docx 只能保证**XML 层面的值正确**；真实 Word 的样式解析、中文样式名识别、页脚页码连续性需要 COM 复核。本技能已验证过以下脚本可跑通（脚本模式：把中文样式名写进 UTF-8 JSON 传入，避免 PowerShell 编码问题）：

```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $false; $word.DisplayAlerts = 0
$doc = $word.Documents.Open($path, $false, $true)
$s = $doc.Styles.Item($name)                     # $name 用中文名，如 "001一级标题"
"$name sz=$($s.Font.Size) font=$($s.Font.NameFarEast) jc=$($s.ParagraphFormat.Alignment)"
$doc.Close($false); $word.Quit()
```

COM 返回值语义（实测）：
- `Font.Size` 单位是**磅**（12 对应 `w:sz=24`）。
- `ParagraphFormat.Alignment`：0=左 1=居中 2=右 3=两端。
- `ParagraphFormat.LineSpacingRule`：0=单倍 1=1.5倍 5=多倍；`LineSpacing` 在 1.5 倍时返回 18（磅）。
- `Shading.BackgroundPatternColor` 是 `WdColor` 整数，**BRG 序**：`13431551` → `#FFF2CC`。
- `KeepWithNext`/`KeepTogether`/`PageBreakBefore` 返回 -1（真）/ 0（假）。
- `Rows.Alignment`：1 = 居中。
- `Border.LineWidth`：12 = 1.5pt，4 = 0.5pt。

## 9. 常见报错与处理

| 现象 | 原因 | 处理 |
|---|---|---|
| `ValueError: not enough values to unpack` | 传给 `etree.Element` 的标签没带前缀 | 用 `qn("w:xxx")`，或本技能的 `F.make()`（自动补 `w:`） |
| `AttributeError: 'HeaderPart' object has no attribute 'document'` | 从页眉 part 反查 doc 失败 | 显式把 `doc` 传进 `set_header/set_footer` |
| 保存后 ipynb/zip 里出现 `Duplicate name: word/numbering.xml` | 同时用 `doc.part.package` 新建 part 又保留原 part | 直接改 `doc.part.numbering_part._element`，不要再建 part |
| 打印中文报 `UnicodeEncodeError: 'gbk'` | Windows 控制台编码 | 脚本开头 `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")` |
| Word 打开提示"文件已损坏" | 缺 `tblGrid`、缺 `[Content_Types].xml` Override、或 `body` 末元素不是 `sectPr` | 按 references/table-and-numbering.md 的完整性清单逐项检查 |
| 目录 F9 后为空 | 标题段缺 `_Toc…` 书签 | 插 `bookmarkStart/bookmarkEnd`；或直接让 Word 重新生成目录 |
| 表格宽度没满页 | 写了 `w:type="dxa"` 或用 `Table Grid` 样式覆盖 | 改回 `w:w="5000" w:type="pct"` |

## 10. 一个完整的可运行示例

见 [scripts/selftest.py](../scripts/selftest.py)：建文件 → 写 5 级标题 + 表格 → `check` → 断言必需 part 齐全。运行：

```bash
python scripts/selftest.py
# 期望输出：结果: 通过 (FAIL 0 / WARN 0) / check: PASS / missing parts: []
```
