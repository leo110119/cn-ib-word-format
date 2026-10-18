# IBD Tools 插件能力清单与实证文件对照

本技能的全部格式参数来自两条独立来源的交叉核验。**冲突时以插件为准**（用户口径：实例偶有几处不一致，插件总体一定是对的）。

## 一、插件本体

`IBD Tools by SyH@20240129.zip` 含两个文件：

| 文件                        | 说明                                                             |
| ------------------------- | -------------------------------------------------------------- |
| `normal.dotm`（380,761 字节） | Word 全局模板。替换 `%APPDATA%\Microsoft\Templates\normal.dotm` 即完成安装 |
| `自动安装插件.bat`              | 安装/卸载脚本，分别支持微软 Office Word 与金山 WPS                             |

安装路径（插件脚本内定死）：

- Word：`C:\Users\<用户>\AppData\Roaming\Microsoft\Templates\normal.dotm`
- WPS：`C:\Users\<用户>\AppData\Roaming\kingsoft\office6\templates\wps\zh_CN\normal.dotm`

模板内部结构（`normal.dotm` 的 OOXML part）：

| Part                          | 用途                                                                  |
| ----------------------------- | ------------------------------------------------------------------- |
| `word/vbaProject.bin`（362 KB） | 全部宏代码（`NewMacros` 模块约 3,300 行）                                      |
| `word/numbering.xml`          | 多级编号定义（`第%1节 / %2、 / （%3） / %4、 …`）                                 |
| `word/styles.xml`             | 内置样式 + `表格前单位` / `表格后说明` / `报告表格` / `正文缩进`                          |
| `word/glossary/document.xml`  | 构建基块：`表格前单位`（"单位：万元"）、`表格后说明`（"注："）、`报告目录`（含 TOC 域的完整目录样例）、`申报文件页码` |
| `customUI/customUI14.xml`     | 功能区 XML（IBD Tools 选项卡、8 个按钮组）                                       |
| `customUI/images/*.png`       | 按钮图标                                                                |
| `word/customizations.xml`     | 快捷键绑定                                                               |

## 二、功能区按钮 → VBA 宏名

任务窗格按钮分 8 组。**通过 COM `Application.Run("宏名")` 可直接调用**（宏随 normal.dotm 全局可用）。

### 文字组

| 按钮                                          | 宏名                                                   | 作用                                          |
| ------------------------------------------- | ---------------------------------------------------- | ------------------------------------------- |
| 楷体加粗                                        | `BoldkaitiRibbon`                                    | 选中范围设为楷体加粗（提示待核对内容）                         |
| 宋体+Times                                    | `SongtiTNR`                                          | 设为宋体 + Times New Roman、不加粗                  |
| 智能粘贴                                        | `pasteastext`                                        | 粘贴为文本；粘到表格多个单元格时自动适配字体/字号/对齐                |
| 删除空格                                        | `DelblankRibbon`                                     | 删除选中范围内的空格与换行（`^l`），不删回车                    |
| 空白为无                                        | `Delblank2ribbon`                                    | 删除空格、回车（`^p`）、换行                            |
| 中英标点                                        | `ToggleInterpunction`                                | 英文标点 → 中文标点，再调 `formatnum` 保留数字间的半角 `,` `.` |
| 下个楷体 / 下个四位数字 / 下个表格                        | `findnextkaiti` / `findnextnumber` / `findnexttable` | 快速定位待核对处                                    |
| 一键标黄                                        | `HighlightRibbon`                                    | 选中范围黄色突出显示（切换）                              |
| 一键排版 → 正文 / 1–6 级标题 / 表格前单位 / 表格后说明 / 设置快捷键 | `正文1..6Ribbon`、`NEEQ正文Ribbon` 等                      | 段落级套用                                       |
| 模板文字排版 / 表格衔接排版 / 表格文字排版                    | `NEEQ…Ribbon`、`newtablestyle3`                       | 新三板专用                                       |

### 表格组

| 按钮        | 宏名                                       | 关键参数                                                                                       |
| --------- | ---------------------------------------- | ------------------------------------------------------------------------------------------ |
| 表格美颜①     | `newtablestyleribbon`                    | 三线表；字体 9pt（小五）；行高 0.7cm；表头底纹 `#FFD900`；顶 `thinThickSmallGap 3pt`、底 `thickThinSmallGap 3pt` |
| 表格美颜②     | `newtablestyle2ribbon`                   | 四边线表；字体 10.5pt（五号）；行高 0.7cm；无底纹；四边 `single 1.5pt`、内部 `single 0.5pt`                        |
| 粘贴并套用表格样式 | `pastenewtablestyle`                     | 粘贴 Excel 表格后自动美颜①                                                                          |
| 全文表格美颜①/② | `allnewtablestyle` / `allnewtablestyle2` | 遍历全文表格                                                                                     |
| 首行重复      | `firstlinerepeat`                        | `Rows.HeadingFormat = wdToggle`                                                            |
| 与下同页      | `keepwithnext`                           | `ParagraphFormat.KeepWithNext = wdToggle`                                                  |
| 单位为万元     | `Unit10000`                              | 插入构建基块"表格前单位"（"单位：万元"）                                                                     |
| 表下加注      | `explain`                                | 插入构建基块"表格后说明"（"注："）                                                                        |
| 粘贴链接      | `pastelink`                              | Excel 单元格区域链接，逐格改写为 `R{n}C{m}` 引用                                                          |
| 更新链接      | `updatelink`                             | 解锁 → 更新 → 重新锁定域                                                                            |
| 上方合计数     | `sumabove`                               | `=Sum(above)` 域，格式 `#,##0.00`                                                              |

### 数字组

| 按钮     | 宏名                 | 作用                                                   |
| ------ | ------------------ | ---------------------------------------------------- |
| 千分位    | `qianfen`          | `Format(x, "#,##0.00;-#,##0.00; ")`，仅处理 ≥4 位整数       |
| 除以一万   | `DivIvan`          | 值 / 10000；绝对值 < 0.01 → `-`                           |
| 自动加%   | `insertpercent`    | 文本以数字结尾则追加 `%`                                       |
| 小数位    | `round1`…`round10` | 插入 `①②③④⑤⑥⑦⑧⑨⑩`                                      |
| 表格序号   | `表格加序号`            | 选中区域填 1..n 并居中                                       |
| 数字降序   | `number321`        | 按第 2 列降序排序                                           |
| 自动计算占比 | `getpercent`       | 单元格 / 合计（末行同列），`Format(…, "0.00%")`；有合并单元格时弹框要求手输合计数 |

### 排版组

| 按钮           | 宏名                    | 作用                                                            |
| ------------ | --------------------- | ------------------------------------------------------------- |
| 设置为招股说明书常用样式 | `formatstyle`         | 建/改 `标题 1–7`、`正文缩进`、`表格前单位`、`表格后说明`、`TOC 标题`、`TOC 1/2`，并写多级编号 |
| 设置为回复报告常用样式  | `formatfeedbackstyle` | 同上，编号改 `问题%1 / %2`                                            |
| 本章样式         | `importstyles`        | 从其他文档导入样式                                                     |
| 关闭样式自动更新     | `disablestyleupdate`  | `Styles(x).AutomaticallyUpdate = False`                       |
| 插入目录         | `insertcatalog`       | 插入构建基块"报告目录"（含 `TOC \o "1-2" \h \z \u`）                       |
| 插入横向页        | `Insertlandscape`     | 前后各插分节符，再切纸向                                                  |
| 添加页码         | `addpagenumber`       | 第 1 节页脚居中加页码，字号 10.5pt，清空页眉内容并去页眉下框线                          |
| 插入图片（零边距）    | `insertpics`          | 新建 0 边距文档并批量插图                                                |
| 一键清理         | `onekeyclean`         | 接受全部修订 + 关闭修订 + 删全部批注 + 刷目录页码 + 清底纹 + 全文设黑                    |
| 楷体转宋体        | `kaititosong`         | 全文查找楷体 → 替换为宋体 + TNR + 不加粗                                    |
| 列表转文本        | `convertlisttotext`   | 自动编号转手写文字                                                     |
| 邮件合并         | `mymailmerge`         | 按 Excel 第 1/2 行字段合并生成多个文件                                     |
| 生成无 XBRL 文件  | `noxbrlfile`          | 另存为 `.doc`                                                    |
| 合并批注         | `合并批注`                | 选原文档 + 多份修订文档，`MergeDocuments` 只比较批注                          |
| 清洗 PDF       | `清洗pdf`               | 清理 PDF 复制产生的多余段落标记与断行正则                                       |

### 批注 / 行列校验组

| 按钮          | 宏名                                           |
| ----------- | -------------------------------------------- |
| 加批注 / 删批注   | `addcomments` / `delcomments`                |
| 计算字数（单处/多处） | `countword` / `countwords`                   |
| 选行校验 / 选列校验 | `RowCheck` / `ColumnCheck`——校验所选行/列标签总和是否为 0 |
| 一键取和        | `getsum`——求和并复制数值                            |

### 快捷键（插件 `设置排版快捷键` 宏写入当前文档）

`Ctrl+1..6` → 标题 1..6｜`Ctrl+\`` → 正文｜`Ctrl+7` → 表格前单位｜`Ctrl+8`→ 表格后说明。`word/customizations.xml` 另记录模板级绑定：`Ctrl+0`/`Ctrl+1..9`（acd1..acd9）、`Ctrl+Shift+S`（acd12）、`Ctrl+Shift+Z\`（acd11）。

## 三、实证文件对照

工作区 `D:\AI\deepseek\skill 技能炼化\` 下 4 份真实申报文件（均为世纪证券出具、标的为乐山巨星农牧股份有限公司的向特定对象发行项目）：

| 文件                 | 大小     | 关键格式信息                                                                                                                              |
| ------------------ | ------ | ----------------------------------------------------------------------------------------------------------------------------------- |
| `3-1 证券发行保荐书.docx` | 2.6 MB | 样式族：`标题 1–7`（带 autoRedefine）+ `001–011`；**表格实为四边线 `boxed`**；15 张表；页脚 `3-1-{PAGE}`                                                   |
| `3-2 上市保荐书.docx`   | 93 KB  | 纯 `001/002/003/004/005/006/009/011 + TOC1/TOC2`；10 张表（`boxed`）；页眉 9pt 宋体 + 下框线；2 节（封面 `pgNumType start=0` + `titlePg`，正文 `start=1`） |
| `3-3 发行保荐工作报告.doc` | 6.5 MB | 旧二进制 `.doc` 格式，无法用 python-docx 读取；如需处理先另存为 `.docx`                                                                                  |
| `3-4 尽职调查报告.docx`  | 4.1 MB | 20+ 节（大量横向页 `w:orient="landscape"`）；132+ 张表；`001/002/003/004/005/006/009` 族；目录域另含 `= 2 \* ALPHABETIC`、`= 1 \* GB3` 等嵌套编号域           |

### 插件与实例的已知差异（**以插件为准**）

| 项目           | 插件定值                               | 实例偏差                            | 采用                                          |
| ------------ | ---------------------------------- | ------------------------------- | ------------------------------------------- |
| 表格线          | 美颜① 三线表                            | 3-1/3-2 均用美颜② 四边线               | 两套口径都在合同里，按文件类型选一个并全篇统一                     |
| 002/003 首行缩进 | `firstLineChars=200 firstLine=480` | 3-1 写 `firstLine=200`（半字符）      | 插件值 480                                     |
| 001 大纲级别     | `outlineLvl=0`                     | 3-2 的 001 为 0，002 也是 0（应为 1）    | 合同按 0/0/1/2 层次的合理值（002=0 与插件一致，003=1、004=2） |
| 表格字号         | 美颜① 9pt / 美颜② 10.5pt               | 3-2 表格 run 里大量 `sz` 缺失（继承）      | 出错时按插件显式写入                                  |
| 页眉对齐         | 插件 `addpagenumber` 清空页眉            | 实例保留页眉（"发行人 + 空格 + 报告期 + 文件全名"） | 实例实践更符合申报要求，合同采纳                            |
| 自动编号         | 插件写入 `numId=15`                    | 实例 `numPr` 计数为 0（中文序号手写）        | 两种方式都在合同里，不混用                               |

## 四、何时改用插件 COM 路线

python-docx 能覆盖 95% 的格式需求。以下场景建议直接调插件宏（需本机装 Word + 已安装 IBD Tools）：

| 场景              | 建议命令                                                            |
| --------------- | --------------------------------------------------------------- |
| 一次整治整篇已有文档的标题层级 | `Application.Run("标题1")` … `("标题6")`、`("正文")`                   |
| 一次性美颜全文表格       | `Application.Run("allnewtablestyle")` 或 `("allnewtablestyle2")` |
| 定稿前清场           | `Application.Run("onekeyclean")`                                |
| 楷体残留清理          | `Application.Run("kaititosong")`                                |
| 批量换算单位          | `Application.Run("Unit10000")`、`("qianfen")`、`("DivIvan")`      |
| 校验表格勾稽          | `Application.Run("RowCheck")`、`("ColumnCheck")`                 |
| 合并多份修订批注        | `Application.Run("合并批注")`                                       |

COM 调用骨架（PowerShell）：

```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $false; $word.DisplayAlerts = 0
$doc = $word.Documents.Open($path)      # 第三个参数 $true = 只读
$word.Run("allnewtablestyle")
$doc.Save(); $doc.Close($false); $word.Quit()
```

注意：`allnewtablestyle` 会弹 `MsgBox("是否进行表格批量美颜？")` 阻塞自动化。无人值守时先在本机模板里把这两个宏的 `MsgBox` 去掉，或改用 python-docx 的 `format_table`（本技能已实现等价逻辑）。

## 五、安装 / 卸载

Word：

```powershell
taskkill /im winword.exe
Copy-Item "$env:APPDATA\Microsoft\Templates\normal.dotm" "$env:APPDATA\Microsoft\Templates\normal-backup.dotm"
Copy-Item "normal.dotm" "$env:APPDATA\Microsoft\Templates\" -Force
```

WPS：目标目录 `%APPDATA%\kingsoft\office6\templates\wps\zh_CN\`。

**关键警告**：装插件会替换全局 `normal.dotm`。本机实测 `Normal.dotm` 为 436,397 字节（已装插件版），备份 `normal-backup.dotm` 为 19,684 字节（原版）。卸载时 `del normal.dotm` 会删掉插件模板但**不会自动恢复备份**——需手工把 `normal-backup.dotm` 改名为 `normal.dotm`。

这也是为什么本技能选择**文档级注入**（把样式写进 .docx 自身）而不是要求装插件：文档自带样式后，在不装插件的机器上打开也完全一致；插件只是创作时的加速器。

## 六、字体依赖

标准要求 宋体 / 黑体 / Times New Roman 三个字体族。Windows 中文版自带；macOS / Linux 需自备（宋体可用 `Noto Serif CJK SC` 近似，但**不要**把近似字体名写进文档 XML——`w:eastAsia="宋体"` 必须保留，由打开方决定替换）。
