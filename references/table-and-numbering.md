# 表格、编号、页眉页脚、目录的 XML 级细节

所有片段可直接嵌入 `word/document.xml` / `word/numbering.xml`。命名空间前缀 `w` = `http://schemas.openxmlformats.org/wordprocessingml/2006/main`。

## 一、表格 XML 骨架

### 1. `tblPr`（表格属性，按此顺序）

```xml
<w:tblPr>
  <w:tblStyle w:val="报告表格"/>
  <w:tblW w:w="5000" w:type="pct"/>
  <w:jc w:val="center"/>
  <w:tblBorders>
    <!-- boxed 口径 -->
    <w:top     w:val="single" w:sz="12" w:space="0" w:color="auto"/>
    <w:left    w:val="single" w:sz="12" w:space="0" w:color="auto"/>
    <w:bottom  w:val="single" w:sz="12" w:space="0" w:color="auto"/>
    <w:right   w:val="single" w:sz="12" w:space="0" w:color="auto"/>
    <w:insideH w:val="single" w:sz="4"  w:space="0" w:color="auto"/>
    <w:insideV w:val="single" w:sz="4"  w:space="0" w:color="auto"/>
  </w:tblBorders>
  <w:tblCellMar>
    <w:top    w:w="0"   w:type="dxa"/>
    <w:left   w:w="108" w:type="dxa"/>
    <w:bottom w:w="0"   w:type="dxa"/>
    <w:right  w:w="108" w:type="dxa"/>
  </w:tblCellMar>
  <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1"
             w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
</w:tblPr>
```

三线表口径替换 `tblBorders` 为：

```xml
<w:top    w:val="thinThickSmallGap" w:sz="24" w:space="0" w:color="auto"/>
<w:bottom w:val="thickThinSmallGap" w:sz="24" w:space="0" w:color="auto"/>
<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:left   w:val="none"/>  <w:right w:val="none"/>  <w:insideV w:val="none"/>
```

蓝线密集表口径：六个边全为 `w:val="single" w:sz="4" w:color="4881BD"`。

### 2. 行属性 `trPr`

```xml
<w:trPr>
  <w:cantSplit/>                      <!-- 禁止跨页断行 -->
  <w:trHeight w:val="397" w:hRule="atLeast"/>  <!-- 0.7cm 最小值 -->
  <w:tblHeader/>                      <!-- 仅首行：跨页重复 -->
  <w:jc w:val="center"/>
</w:trPr>
```

### 3. 单元格属性 `tcPr`

```xml
<w:tcPr>
  <w:tcW w:w="1779" w:type="pct"/>
  <w:shd w:val="clear" w:color="auto" w:fill="FFD900"/>   <!-- 仅美颜③表头 -->
  <w:vAlign w:val="center"/>
</w:tcPr>
```

`vAlign` 取值：`top` / `center` / `bottom`。表内对齐靠 `pPr/jc`，不是 `tcPr`。

### 4. 数值判定（决定左右对齐）

插件逻辑：`IsNumeric(单元格文本)` **或** 文本等于 `-` **或** 去掉 `%` 后 `IsNumeric` → 右对齐。实现见 `filing_contract.is_numeric_cell()`：

```python
t = text.strip().replace(",", "").replace(" ", "")
if t.endswith("%"): t = t[:-1]
float(t)  # 成功即为数字
```

注意 `1,234.56` 要先剥掉千分位逗号再判断；`28.31%` 剥掉 `%` 后判断。

## 二、多级编号 XML

```xml
<w:numbering>
  <w:abstractNum w:abstractNumId="0" w15:restartNumberingAfterBreak="0">
    <w:nsid w:val="FFFFFF7C"/>
    <w:multiLevelType w:val="multilevel"/>
    <w:tmpl w:val="FFFFFF7C"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="chineseCounting"/>
      <w:pStyle w:val="Heading1"/>
      <w:lvlText w:val="第%1节"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman"/>
        <w:sz w:val="32"/><w:szCs w:val="44"/>
      </w:rPr>
    </w:lvl>
    <!-- ilvl 1..6 依次：%2、 / （%3） / %4、 / （%5） / %6） / %7 -->
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
```

要点：
- `chineseCounting` = 一、二、三；`decimalEnclosedCircle` = ①；`decimal` = 1。
- `pStyle` 把该级编号与段落样式绑定；样式里再写 `numPr/ilvl` 指向同一个 `numId`（插件用 `numId=15`，只要 `num` 映射存在即可）。
- 想临时取消某段编号：段落 `pPr/numPr` 写 `<w:numId w:val="0"/>`（TOC 标题即如此处理）。
- 想从某级重新起编：`<w:lvlOverride w:ilvl="1"><w:startOverride w:val="1"/></w:lvlOverride>` 挂在 `w:num` 下。

## 三、节与页码

```xml
<!-- 封面节（文档首个 sectPr 在 body 末尾，或上一节末段的 pPr 内） -->
<w:sectPr>
  <w:headerReference w:type="default" r:id="rId9"/>
  <w:footerReference w:type="default" r:id="rId10"/>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"
           w:header="851" w:footer="992" w:gutter="0"/>
  <w:pgNumType w:start="0"/>
  <w:cols w:space="425"/>
  <w:titlePg/>                      <!-- 首页页眉页脚独立 -->
  <w:docGrid w:type="lines" w:linePitch="312"/>
</w:sectPr>
```

正文节同上但去掉 `titlePg`，`pgNumType w:start="1"`。分节符段落写作 `<w:p><w:pPr><w:sectPr>…</w:sectPr></w:pPr></w:p>`，插在上一节最后一页之后。

`headerReference w:type` 取值：`default`（奇数页+全篇）、`first`（首页）、`even`（偶数页）。只用 `default` + `titlePg` 就够。

## 四、页眉页脚内容

页眉（`word/header1.xml`）：

```xml
<w:p>
  <w:pPr>
    <w:pStyle w:val="Header"/>
    <w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>
    <w:jc w:val="left"/>
  </w:pPr>
  <w:r><w:t>乐山巨星农牧股份有限公司                              2023年度向特定对象发行股票之上市保荐书</w:t></w:r>
</w:p>
```

页脚（`word/footer1.xml`）：

```xml
<w:p>
  <w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>
  <w:r><w:t>3-2-</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText xml:space="preserve">PAGE   \* MERGEFORMAT</w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:t>3</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>
```

页码前缀（`3-2-`）随申报文件编号变化；页码字体 10.5pt，域内缓存值随便写，Word 打开会重算。

## 五、目录域

```xml
<w:p>
  <w:pPr><w:pStyle w:val="TOC1"/>
    <w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8296"/></w:tabs>
  </w:pPr>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:instrText xml:space="preserve"> TOC \o "1-2" \h \z \u </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:hyperlink w:anchor="_Toc107574856" w:history="1">
    <w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr><w:t>第一节</w:t></w:r>
    <w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r>
    <w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF _Toc107574856 \h </w:instrText></w:r>
    <w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:t>2</w:t></w:r>
    <w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
  </w:hyperlink>
</w:p>
```

- `\o "1-2"` 只收 1–2 级；`\h` 生成超链接；`\z` 网页视图隐藏页码；`\u` 用大纲级别而非样式名。
- **每个标题段落必须有 `w:bookmarkStart` / `w:bookmarkEnd` 的 `_Toc…` 书签**，否则 F9 刷新后目录为空。书签放在标题段内、文字之前。
- 制表位 `w:pos="8296"` = 页面正文宽度右边缘（A4 纵 + 左右 1800 边距），点线前导符。
- 程序生成时不必手写已填充的目录条目——只插标题段 + 域，然后 `doc.Fields.Update()` 或让用户按 F9。

## 六、样式 XML 骨架

```xml
<w:style w:type="paragraph" w:customStyle="1" w:styleId="001">
  <w:name w:val="001一级标题"/>
  <w:basedOn w:val="Normal"/>
  <w:next w:val="005"/>
  <w:qFormat/>
  <w:pPr>
    <w:keepNext/><w:keepLines/><w:pageBreakBefore/>
    <w:widowControl w:val="0"/>
    <w:spacing w:beforeLines="50" w:before="50" w:afterLines="50" w:after="50"
               w:line="360" w:lineRule="auto"/>
    <w:jc w:val="center"/>
    <w:outlineLvl w:val="0"/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman"
              w:cs="Times New Roman"/>
    <w:b/><w:bCs/><w:kern w:val="2"/>
    <w:sz w:val="32"/><w:szCs w:val="28"/>
  </w:rPr>
</w:style>
```

必须项与顺序：`name` → `basedOn` → `next` → `link` → `uiPriority` → `qFormat` → `pPr` → `rPr`。`w:customStyle="1"` 让样式出现在"自定义"组。

**中文字体名的坑**：`w:rFonts` 的 `eastAsia` 属性必须写中文名（`宋体`/`黑体`），写在 `ascii`/`hAnsi` 上无效。同时设 `w:cs` 保证复杂文种一致。`w:hint="eastAsia"` 用于告诉 Word 该段以中文字形渲染（构建基块里常见）。

## 七、常用枚举值对照

| 枚举 | 值 |
|---|---|
| `w:jc`（段落对齐） | `left` / `center` / `right` / `both`（两端）/ `distribute` |
| `w:vAlign`（单元格垂直） | `top` / `center` / `bottom` |
| `w:lineRule` | `auto`（倍数，360 = 1.5 倍）/ `exact`（固定值）/ `atLeast` |
| `w:hRule`（行高） | `auto` / `exact` / `atLeast` |
| `w:outlineLvl` | 0–9，**9 = 正文级**（TOC 标题用 9 表示不参与大纲） |
| `w:br` | `w:type="page"` 分页符 / 默认换行 |
| `w:sz` 边框 | 4=0.5pt / 8=1pt / 12=1.5pt / 18=2.25pt / 24=3pt |
| `w:sz` 字号 | 半磅：24=12pt / 21=10.5pt / 18=9pt / 28=14pt / 32=16pt / 44=22pt |

## 八、OOXML 完整性检查清单

生成文件后，用 `python scripts/filing_docx.py check` 之外，还应确认：

1. `[Content_Types].xml` 里有 `word/document.xml`、`word/styles.xml`、`word/numbering.xml` 三个 Override，且类型正确（前两者 `wordprocessingml.document.main+xml` / `wordprocessingml.styles+xml`）。
2. `word/_rels/document.xml.rels` 里有指向 `styles.xml`、`numbering.xml`、`header1.xml`、`footer1.xml` 的关系，且 `headerReference` 的 `r:id` 与之匹配。
3. `word/document.xml` 根元素带全部 `w`、`r` 命名空间声明；`w:document/w:body` 的最后一个子元素必须是 `sectPr`。
4. `styles.xml` 中的 `docDefaults` 在 `latentStyles` 之前。
5. 表格 `tblGrid/gridCol` 的 `w:w` 之和与 `tblW` 相符（`pct` 类型下按比例写即可，但缺失 `tblGrid` 会导致 Word 报文件损坏）。
