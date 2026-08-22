# leafdoc CLI

多格式 Office / PDF / 电子书 / 元文件命令行工具。按 `--input` 扩展名（或 `--type`）路由；xlsx 提供完整读写，其余格式支持提取、渲染、导出与格式转换。

> **自主研发 · 纯 Go**：解析、排版、渲染与导出均为 **纯 Golang 自研内核**，**不依赖** WPS、Microsoft Office，也**不依赖** LibreOffice / OpenOffice 等其它第三方办公套件或开源引擎。部署与运行只需本机 Go 构建产物（及系统字体等常规环境），即可完成文本提取、版式渲染、PDF/HTML/SVG 导出。

在线帮助：`leafdoc --help` · `leafdoc --help <命令>`

---

## 功能概览

| 能力 | 说明 |
|------|------|
| 多格式解析 | Word / Excel / PowerPoint（OOXML + 二进制）、PDF、DjVu、RTF、HTML、Markdown、纯文本、CSV、EPUB、MOBI、AZW3、EMF、WMF |
| 文本与媒体 | 提取正文/页眉页脚、段落、字体、内嵌图片 |
| 渲染导出 | 分页/分表/分幻灯片 → SVG / PNG / JPG；导出 PDF、HTML 预览（纯 Go，无需 Office） |
| 格式转换 | 见下文 **格式转换表**（`export.*` / `render`；底层 `document.Export` + `OutputFormatter`） |
| xlsx 读写 | 单元格、公式、样式、合并、批注、筛选、保护、打印、迷你图、图表、VBA 等 |
| 文档合并 | `doc.merge` 多文件合并为单一 DOCX |
| 浏览器预览 | `--open` / 交互模式：边渲边写 HTML，本地 HTTP 翻页 |

---

## 用法

```text
leafdoc                              # 交互模式（HTTP + REPL）
leafdoc --input <文件> --cmd <命令> [--output <路径>] [选项...]
leafdoc --open [--input <文件>]
leafdoc --cmd=create --output=new.xlsx
leafdoc --cmd=doc.merge --input=a.docx --value=b.xlsx --output=out.docx
leafdoc --help
leafdoc --help <命令>
```

### 交互模式

无参数启动后：

- REPL：`/open`（选文件）、`/open?path=...`、`/list`、`/help`、`/quit`
- 浏览器：`http://127.0.0.1:<port>/open?path=C:/file.pdf`
- 预览写入 `.cache/<md5>/`，多文档挂载在 `/v/<md5>/`

### 全局参数

| 参数 | 说明 |
|------|------|
| `--input` | 输入文件（`create` / `doc.merge` / `help` / `--open` 可缺） |
| `--output` | 输出文件或目录 |
| `--cmd` | 命令（可重复或逗号分隔；xlsx 可链式） |
| `--password` | 解密 / 加密密码 |
| `--type` | 强制类型：`docx` `doc` `pptx` `ppt` `xlsx` `xls` `pdf` `djvu` `rtf` `html` `csv` `markdown` `text` `emf` `wmf` `epub` `mobi` `azw3` |
| `--page` | 页/幻灯片索引（0-based） |
| `--filter` | 文本过滤：`all` / `document` / `header` / `footer`（可用 `+`） |
| `--pages` | PDF/DjVu 取文页码（0-based）：`0,2` 或 `0-2` |
| `--format` / `--dpi` | 渲染格式与 DPI |
| `--all-in-one` | `render`：流式文档输出单张连续长图（docx/doc/html/md/rtf） |
| `--option` | HTML：`canvas`（默认）\| `code`（Markdown 语义 HTML；`export.html` / `--open`） |
| `--max-width` / `--max-height` | EMF/WMF 尺寸上限 |
| `--disable-formula` | xlsx 跳过公式计算 |
| `--open` | HTML 预览（本地 HTTP；底栏页码与进度） |

### 类型识别

| 扩展名 | 组件 |
|--------|------|
| `.xlsx` `.xlsm` `.xltx` `.xltm` `.xlam` | xlsx |
| `.xls` | xls |
| `.docx` `.dotx` `.docm` `.dotm` | docx |
| `.doc` | doc |
| `.pptx` `.pptm` `.ppsx` `.ppsm` `.potx` `.potm` | pptx |
| `.ppt` | ppt |
| `.pdf` | pdf |
| `.djvu` `.djv` | djvu |
| `.rtf` | rtf |
| `.html` `.htm` | html |
| `.md` `.markdown` `.mdown` `.mkd` | markdown |
| `.txt` `.text` | text |
| `.csv` | csv |
| `.epub` | epub |
| `.mobi` `.prc` `.azw` | mobi |
| `.azw3` | azw3 |
| `.emf` / `.wmf` | emf / wmf |

---

## 格式转换表（From → To）

**图例**：✓ = CLI 有对应导出命令且实现支持 · — = 不支持或未提供命令。具体命令见下文「通用命令」与「格式特有命令」。

### 文档 / 幻灯片 / 表格

| From | → PDF | → DOCX | → HTML | → SVG/PNG/JPG | → PPTX | → EPUB | → MOBI | → XLSX | → CSV |
|------|:-----:|:------:|:------:|:-------------:|:------:|:------:|:------:|:------:|:-----:|
| **docx** | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | — | — |
| **doc** | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | — | — |
| **pptx** | ✓ | ✓ | ✓ | ✓ | ✓ | — | — | — | — |
| **ppt** | ✓ | ✓ | ✓ | ✓ | ✓ | — | — | — | — |
| **xlsx** | ✓ | ✓ | ✓ | ✓ | ✓ | — | — | ✓ | ✓ |
| **xls** | ✓ | ✓ | ✓ | ✓ | ✓ | — | — | ✓ | ✓ |
| **pdf** | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | — | — |
| **djvu** | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | — | — |
| **rtf** | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | — | — |
| **html** | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | — | — |
| **markdown** | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | — | — |
| **text** | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | — | — |
| **csv** | ✓ | ✓ | ✓ | ✓ | ✓ | — | — | ✓ | — |
| **epub** | ✓ | ✓ | ✓ | ✓ | — | — | — | — | — |
| **mobi** | ✓ | ✓ | ✓ | ✓ | — | — | — | — | — |
| **azw3** | ✓ | ✓ | ✓ | ✓ | — | — | — | — | — |

**固定页格式**（pdf / djvu / pptx / ppt）：跨格式导出走 **display 元素**（文本框、图片、矩形），不经整页 PNG 光栅。ppt 无绘制管线时先升为 pptx。

### 元文件（EMF / WMF）

| From | → SVG | → PNG/JPG |
|------|:-----:|:---------:|
| **emf** | ✓ | ✓ |
| **wmf** | ✓ | ✓ |

### 多文件合并

| From | → DOCX |
|------|:------:|
| docx + xlsx / csv / … | ✓ |

### 常用别名

| 目标 | 命令别名 |
|------|----------|
| PDF | `export-pdf` · `topdf` |
| DOCX | `todocx` |
| HTML | `export-html` · `tohtml` |
| PPTX | `topptx` |
| EPUB | `toepub` · `export-epub` |
| MOBI | `tomobi` · `export-mobi` |
| XLSX | `toxlsx` |
| CSV | `tocsv` |
| SVG | `to.svg` · `export.svg` |
| 位图 | `to.image` · `to.png` · `export.image` |

### Go API（与 CLI 对应）

```go
document.Export(doc, pdf.NewPDFFormatter(w))
document.Export(doc, epub.NewEPUBFormatter(w, "书名"))
document.Export(doc, mobi.NewMOBIFormatter(w, "书名"))
document.Export(doc, docx.NewDOCXFormatter(w))   // 需 DocxExporter
document.Export(doc, pptx.NewPPTXFormatter(w))   // 需 PptxExporter
document.Export(doc, document.NewHTMLFormatter(dir, opts...))
document.Export(doc, document.NewRenderFormatter(dir, "svg", document.WithAllInOne(true)))
```

---

## 通用命令

适用于多数 `document.Office` 格式（docx / doc / pptx / ppt / xlsx / xls / pdf / djvu / rtf / html / markdown / text / csv / epub / mobi / azw3）：

| 命令 | 说明 |
|------|------|
| `info` / `mime` / `encrypted` | 元信息 |
| `text` / `paragraphs` / `fonts` / `images` | 提取 |
| `render` | 全部页/表/幻灯片渲染到目录 |
| `export.pdf` / `export.html` / `export.docx` / `export.epub` / `export.mobi` | 导出 PDF / HTML / DOCX / EPUB / MOBI（见 **格式转换表**） |
| `page.count` / `render.page` | 页数 / 单页（docx、doc、pdf、djvu、pptx） |
| `doc.merge` | 多文件合并为 DOCX |

```bash
leafdoc --input=a.docx --cmd=text --filter=document --output=a.txt
leafdoc --input=a.pdf --cmd=text --pages=0-1 --output=p01.txt
leafdoc --input=a.pdf --cmd=render --format=svg --output=./pages
leafdoc --input=note.md --cmd=render --format=png --all-in-one --output=./out
leafdoc --input=a.docx --cmd=export.pdf --output=a.pdf
leafdoc --input=a.pptx --cmd=render.page --page=0 --format=png --output=s0.png
leafdoc --cmd=doc.merge --input=a.docx --value=b.xlsx,c.csv --output=merged.docx
```

Markdown 的 `export.html` / `--open` 额外支持 `--option`：

| `--option` | 说明 |
|------------|------|
| `canvas`（默认） | ContentDoc 布局 → HTML Canvas 预览目录 |
| `code` | CommonMark/GFM → 语义 HTML（文件 `.html` 或目录下 `index.html`） |

```bash
leafdoc --input=note.md --cmd=export.html --output=./html
leafdoc --input=note.md --cmd=export.html --option=code --output=note.html
leafdoc --open --input=note.md --option=code --no-open
```

### 与 `document.Office` 选项对应

| CLI | Go API |
|-----|--------|
| `--format=png` | `document.WithFormat("png")` |
| `--all-in-one` | `document.WithAllInOne(true)` |
| `--option=code` | `document.WithHTMLCode()` / `WithHTMLMode("code")` |
| `--filter=document+header` | `WithTextDocument()` + `WithTextHeader()` |
| `--pages=0-1` | `document.WithTextPages(0, 1)` |
| `--dpi=96` | `render.WithHTMLDPI(96)` |

### `--open` 预览

缺省弹选文件框；也可 `--input`。导出到 `.cache/<md5>/`，首页就绪后打开浏览器，边渲边加载。`--no-open` 只导出不弹窗。

```bash
leafdoc --open
leafdoc --open --input=doc.docx --dpi=96
```

---

## 格式特有命令

| 格式 | 命令 | 说明 |
|------|------|------|
| doc | `export.docx` | → DOCX |
| rtf / html / markdown / text | `export.docx` / `export.epub` / `export.mobi` | → DOCX / EPUB / MOBI |
| pdf / djvu | `export.epub` / `export.mobi` | 按页分章（文本层/抽文本） |
| epub / mobi / azw3 | `export.docx` | 电子书 → DOCX |
| djvu | `page.count` / `render.page` / `meta` | 页数、单页栅格、尺寸/DPI |
| ppt | `export.pptx` / `slide.list` | → PPTX；幻灯片文本 |
| xls | `export.xlsx` / `export.csv` / `sheet.list` | → XLSX / CSV |
| xlsx | `export.csv` | → CSV（默认首表，`--sheet` 指定表） |
| csv | `export.xlsx` / `rows` | → XLSX；导出行 |
| docx | `save` / `accept-revisions` / `reject-revisions` | 回写；修订 |
| pptx | `save` / `slide.count` / `slide.export` / `slide.size` / `tags` | 回写；单页拆出 |
| pdf | `save` / `outlines` / `hyperlinks` / `meta` | 回写；书签/链接/Info |
| emf | `to.svg` / `to.image` / `to.bitmap` / `decode` / `info` | SVG / PNG / 位图 |
| wmf | `to.svg` / `to.image` / `to.bitmap` / `text` / `info` | 同上 |

```bash
leafdoc --input=old.doc --cmd=export.docx --output=new.docx
leafdoc --input=old.xls --cmd=export.xlsx --output=new.xlsx
leafdoc --input=book.xlsx --cmd=export.csv --output=sheet.csv
leafdoc --input=book.xlsx --cmd=export.csv --sheet=Data --output=data.csv
leafdoc --input=old.ppt --cmd=export.pptx --output=new.pptx
leafdoc --input=note.md --cmd=export.docx --output=note.docx
leafdoc --input=pic.emf --cmd=to.svg --output=pic.svg --max-width=1024
```

标 `[write]` 的写回类命令必须带 `--output`（详见 `leafdoc --help`）。

---

## xlsx 命令

### 工作簿

`create` · `info` · `mime` · `encrypted` · `save` · `calc-chain`

### 工作表 / 单元格

`sheet.list` `sheet.add` `sheet.delete` `sheet.rename`  
`cell.get` `cell.set` `cell.style` `range.get` `rows.get`  
`cell.merge` `cell.unmerge` `col.width` `col.visible` `row.height` `row.visible`  
`calc`

### 文本 / 图片 / 名称 / 批注

`text` `paragraphs` `fonts` `images` `picture.add`  
`name.list` `name.set` `name.delete`  
`comment.list` `comment.add` `threaded.list` `threaded.add`

### 表 / 筛选 / 链接 / 验证 / 保护 / 打印

`table.add` `filter.set` `filter.apply` `filter.clear`  
`link.add` `validate.set` `validate.delete`  
`protect.sheet` `unprotect.sheet` `protect.workbook` `unprotect.workbook`  
`print.setup` `print.header` `print.margins` `print.options`

### 迷你图 / 条件格式 / 图表 / 样式 / 渲染 / VBA

`sparkline.add` `cf.clear` `chart.add` `chart.templates`  
`style.new` `style.apply`  
`render` `export.html` `export.pdf` `export.pptx` `export.epub` `export.mobi` `export.csv`  
`vba.has` `vba.get` `vba.set` `vba.remove` `vba.sig.get` `vba.sig.set`  
`query.ole` `query.rich` `query.slicer` `query.timeline`

```bash
leafdoc --input=book.xlsx --cmd=cell.get --sheet=Sheet1 --cell=A1
leafdoc --input=book.xlsx --cmd=export.pdf --output=out.pdf
leafdoc --cmd=create --cmd=cell.set --sheet=Sheet1 --cell=A1 --value=42 --output=out.xlsx
```

单条命令完整用法：`leafdoc --help cell.get`。

---

## 作者

微信：`yuyi297341015`
