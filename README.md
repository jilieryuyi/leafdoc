# leafdoc CLI

**AI 大时代，专为 AI 设计的文档 CLI 工具。**

多格式 Office / PDF / 电子书 / 元文件命令行工具。按 `--input` 扩展名（或 `--type`）路由；xlsx 提供完整读写，其余格式支持提取、渲染、导出与格式转换。

简体中文 | [English](README.en.md)

> **自主研发 · 纯 Go**：解析、排版、渲染与导出均为 **纯 Golang 自研内核**，**不依赖** WPS、Microsoft Office，也**不依赖** LibreOffice / OpenOffice 等其它第三方办公套件或开源引擎。部署与运行只需本机 Go 构建产物（及系统字体等常规环境），即可完成文本提取、版式渲染、PDF/HTML/SVG 导出。

在线帮助：`leafdoc --help` · `leafdoc --help <命令>`

---

## 功能概览

| 能力 | 说明 |
|------|------|
| 多格式解析 | Word / Excel / PowerPoint（OOXML + 二进制）、PDF、DjVu、CBZ / CBR（漫画压缩包）、PDG（超星电子书）、RTF、HTML、Markdown、纯文本、CSV、EPUB、FB2（FictionBook）、MOBI、AZW3、Microsoft Reader（LIT）、已编译 HTML 帮助（CHM）、EMF、WMF |
| 文本与媒体 | 提取正文/页眉页脚、段落、字体、内嵌图片 |
| 渲染导出 | 分页/分表/分幻灯片 → SVG / PNG / JPG；导出 PDF、HTML 预览（纯 Go，无需 Office） |
| 格式转换 | 见下文 **格式转换表**（`export.*` / `render`；底层 `document.Export` + `OutputFormatter`） |
| xlsx 读写 | 单元格、公式、样式、合并、批注、筛选、保护、打印、迷你图、图表、VBA 等 |
| 文档合并 | `doc.merge` 多文件合并为单一 DOCX |
| 嵌入编辑器 | `--open` / `--editor`：已接入排版引擎的格式（`.docx` / `.doc` / `.rtf` / `.azw3` / `.epub` / `.fb2` / `.mobi` / `.lit` / `.chm` / `.xlsx` / `.xls` / `.csv` / `.pdf` / `.djvu` / `.cbz` / `.cbr` / `.pdg` / `.pdz` / `.pptx` / `.ppt`）直接开嵌入式编辑器（`#hash=` 多文档会话，支持 `--readonly`）；其余已声明格式说明「缺哪个世界的引擎」后回落 HTML 预览 |
| 浏览器预览 | `--cmd=open`（别名 `preview`）/ 无参数启动的欢迎页：边渲边写 HTML，本地 HTTP 翻页 |
| 欢迎页 | 无参数启动即自动打开：页面上一个「打开文件…」按钮 → 本机对话框 → 以同一份可执行文件起 `leafdoc --open --readonly --input=<文件>`（即只读编辑器） |

---

## 用法

```text
leafdoc                              # 欢迎页（自动开浏览器）+ HTTP + REPL
leafdoc --input <文件> --cmd <命令> [--output <路径>] [选项...]
leafdoc --open [--readonly] [--input <文件>]
leafdoc --cmd=open [--input <文件>] [--option=canvas|code]
leafdoc --cmd=create --output=new.xlsx
leafdoc --cmd=doc.merge --input=a.docx --value=b.xlsx --output=out.docx
leafdoc --help
leafdoc --help <命令>
```

### 欢迎页（无参数）

无参数启动后自动打开浏览器到欢迎页，页面上只有「打开文件…」一个动作：

- 点按钮 → `POST /welcome/open` → **服务端**弹本机文件对话框（Windows：PowerShell OpenFileDialog）
- 选中后以同一份可执行文件另起进程：`leafdoc --open --readonly --input=<选中的文件>`
  - 子进程脱离欢迎进程（Windows `CREATE_NO_WINDOW`，不闪黑框），新标签页打开只读编辑器
  - 即「无需记命令行也能进嵌入式编辑器只读模式」；选项/行为与手敲那条命令完全一致
- 页面还会列出：本次已打开（服务器记录，最多 12 条）、本次进程已开过的 HTML 预览文档

### 交互模式（同一进程的 REPL）

欢迎页之外，同一个 HTTP 服务仍保留原有命令：

- REPL：`/open`（选文件）、`/open?path=...`、`/list`、`/help`、`/quit`
- 浏览器：`http://127.0.0.1:<port>/open?path=C:/file.pdf`
- 欢迎页 JSON 状态：`http://127.0.0.1:<port>/?format=json`
- 预览写入 `.cache/<md5>/`，多文档挂载在 `/v/<md5>/`

### 全局参数

| 参数 | 说明 |
|------|------|
| `--input` | 输入文件（`create` / `doc.merge` / `help` / `--open` 可缺） |
| `--output` | 输出文件或目录 |
| `--cmd` | 命令（可重复或逗号分隔；xlsx 可链式） |
| `--password` | 解密 / 加密密码 |
| `--type` | 强制类型：`docx` `doc` `pptx` `ppt` `xlsx` `xls` `pdf` `djvu` `cbz` `cbr` `pdg` `rtf` `html` `csv` `markdown` `text` `emf` `wmf` `epub` `fb2` `mobi` `azw3` `lit` `chm` |
| `--page` | 页/幻灯片索引（0-based） |
| `--filter` | 文本过滤：`all` / `document` / `header` / `footer`（可用 `+`） |
| `--pages` | PDF/DjVu 取文页码（0-based）：`0,2` 或 `0-2`（CBZ / CBR / PDG 无文本层，不消费该开关；单页出图用 `--page`） |
| `--format` / `--dpi` | 渲染格式与 DPI |
| `--all-in-one` | `render`：流式文档输出单张连续长图（docx/doc/html/md/rtf） |
| `--option` | HTML：`canvas`（默认）\| `code`（Markdown 语义 HTML）；docx：`ast`（经 Document AST 排版；`--debug` / `render` / `export.html`） |
| `--max-width` / `--max-height` | EMF/WMF 尺寸上限 |
| `--disable-formula` | xlsx 跳过公式计算 |
| `--open` | 在嵌入编辑器中打开文档（缺 `--input` 弹文件选择框；未接入引擎的格式回落 HTML 预览） |
| `--editor` | 嵌入式编辑器（`--readonly` 为纯画布只读查看）；`.docx` / `.doc` / `.rtf` / `.azw3` / `.epub` / `.fb2` / `.mobi` / `.lit` / `.chm` / `.xlsx` / `.xls` / `.csv` / `.pdf` / `.djvu` / `.cbz` / `.cbr` / `.pdg` / `.pdz` / `.pptx` / `.ppt` 已接入引擎 |

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
| `.cbz` | cbz |
| `.cbr` | cbr |
| `.pdg` `.pdz` | pdg |
| `.rtf` | rtf |
| `.html` `.htm` | html |
| `.md` `.markdown` `.mdown` `.mkd` | markdown |
| `.txt` `.text` | text |
| `.csv` | csv |
| `.epub` | epub |
| `.fb2` | fb2 |
| `.mobi` `.prc` `.azw` | mobi |
| `.azw3` | azw3 |
| `.lit` | lit |
| `.chm` `.chw` | chm |
| `.emf` / `.wmf` | emf / wmf |

> `.mobi` `.prc` `.azw` 挂同一个 mobi 引擎；但 **`.azw` / `.prc` 是「一个后缀装两种容器」**
> （Amazon 早期既发过老式 MOBI7、也发过 KF8 换壳）：这两个后缀的判族**不看名字、落到内容上** ——
> 内容是 KF8 就改派 azw3 引擎（`cmd/detect.go::ambiguousKindleDocType`，CLI 与编辑器共用），
> 否则仍是 mobi。`.mobi` / `.azw3` 语义单一，不参与该覆盖。

> `.cbz`（Comic Book Zip）只按**扩展名**判定，不看字节：ZIP 容器本身不自我介绍，
> `sniffType` 无法把「一包漫画图」与「一包 docx 部件」分开（见 `cmd/detect.go::sniffType` 注释）。
> 对应地，把 `.docx` 改名成 `.cbz` 会被 cbz 解析器拒掉并提示真实格式。
>
> `.cbr`（Comic Book Rar）则**两种信号都认**：RAR 有独特魔数（`Rar!\x1a\x07`），
> 故 `sniffType` 能在 ZIP/文本探针之前认出它（`.cbr` 在市面上常被改名成 `.cbz`/`.zip`，
> 甚至干脆没有扩展名）；但扩展名仍优先，所以改名成 `.cbz` 的 RAR 由 cbz 引擎报错时会直接告诉你
> 「这其实是 RAR，请改回 `.cbr`」，而不是让你看到一个看不出所以然的「不是 ZIP」。
>
> `.pdg` / `.pdz`（超星 PDG）的分布比上面两个多一条路：**一本 PDG 书最常见的分发形态是
> 普普通通的 `book.zip`**（超星阅读器的「打包下载」给的就是它），扩展名里没有任何线索。
> 所以 pdg 除了按扩展名以外，**还多一条内容裁定**：`sniffType` 看 ZIP/RAR 的成员表，
> **里面都是 `.pdg`（`pdg` 成员数 > 0 且不少于图片成员数）就当 PDG 开**（`pdg.IsPDGArchive`）；
> 真的不是 PDG 时不会抢 —— 不满足该比例就回到原来的 ZIP 探针（docx/epub/…）或报错。
> 单页文件只有 `.pdg` / `.pdz` 两条扩展名。两种失败都有可照做的说明：里面是 `word/document.xml`
> 就提示改回 `.docx`，里面是图片就提示改回 `.cbz` / `.cbr`，只有一个 `bookinfo.dat` 则是下载不完整。

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
| **cbz** | — | — | — | ✓ | — | — | — | — | — |
| **cbr** | — | — | — | ✓ | — | — | — | — | — |
| **pdg** | — | — | — | ✓ | — | — | — | — | — |
| **rtf** | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | — | — |
| **html** | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | — | — |
| **markdown** | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | — | — |
| **text** | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | — | — |
| **csv** | ✓ | ✓ | ✓ | ✓ | ✓ | — | — | ✓ | — |
| **epub** | ✓ | ✓ | ✓ | ✓ | — | — | — | — | — |
| **fb2** | ✓ | ✓ | ✓ | ✓ | — | — | — | — | — |
| **mobi** | ✓ | ✓ | ✓ | ✓ | — | — | — | — | — |
| **azw3** | ✓ | ✓ | ✓ | ✓ | — | — | — | — | — |
| **lit** | ✓ | ✓ | ✓ | ✓ | — | — | — | — | — |
| **chm** | ✓ | ✓ | ✓ | ✓ | — | — | — | — | — |

**固定页格式**（pdf / djvu / cbz / cbr / pdg / pptx / ppt）：跨格式导出走 **display 元素**（文本框、图片、矩形），不经整页 PNG 光栅。ppt 无绘制管线时先升为 pptx。

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

适用于多数 `document.Office` 格式（docx / doc / pptx / ppt / xlsx / xls / pdf / djvu / cbz / cbr / pdg / rtf / html / markdown / text / csv / epub / fb2 / mobi / azw3 / lit / chm）：

| 命令 | 说明 |
|------|------|
| `info` / `mime` / `encrypted` | 元信息 |
| `text` / `paragraphs` / `fonts` / `images` | 提取 |
| `render` | 全部页/表/幻灯片渲染到目录 |
| `export.pdf` / `export.html` / `export.docx` / `export.epub` / `export.mobi` | 导出 PDF / HTML / DOCX / EPUB / MOBI（见 **格式转换表**） |
| `page.count` / `render.page` | 页数 / 单页（docx、doc、pdf、djvu、cbz、cbr、pdg、pptx） |
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

Markdown 的 `export.html` / `--cmd=open` 额外支持 `--option`：

| `--option` | 说明 |
|------------|------|
| `canvas`（默认） | ContentDoc 布局 → HTML Canvas 预览目录 |
| `code` | CommonMark/GFM → 语义 HTML（文件 `.html` 或目录下 `index.html`） |
| `ast` | docx：Layer-3 Document AST（ECMA-376 `w:body`）排版后再渲染，对标默认 Element 路径 |

```bash
leafdoc --input=note.md --cmd=export.html --output=./html
leafdoc --input=note.md --cmd=export.html --option=code --output=note.html
leafdoc --cmd=open --input=note.md --option=code --no-open
leafdoc --debug --input=doc.docx --jpg --option=ast --auto-exit
leafdoc --input=doc.docx --cmd=render --format=png --option=ast --output=./out
```

### 与 `document.Office` 选项对应

| CLI | Go API |
|-----|--------|
| `--format=png` | `document.WithFormat("png")` |
| `--all-in-one` | `document.WithAllInOne(true)` |
| `--option=code` | `document.WithHTMLCode()` / `WithHTMLMode("code")` |
| `--option=ast` | `document.WithAST(true)` / `docx.File.SetLayoutUseAST(true)` |
| `--filter=document+header` | `WithTextDocument()` + `WithTextHeader()` |
| `--pages=0-1` | `document.WithTextPages(0, 1)` |
| `--dpi=96` | `render.WithHTMLDPI(96)` |

### `--open`（嵌入编辑器）

docx 缺省弹选文件框（也可 `--input`），直接进入嵌入式编辑器（与 `--editor` 同一运行时）：URL 带 `#hash=<sid>`，刷新可恢复，改 hash 即可切文件且无需刷新页面。`--readonly` 为纯画布只读（无工具栏/底栏，文本可选可复制）。

准入按能力矩阵判定，与 CLI / HTTP 同一套答案：

| 格式 | 世界 | 编辑器行为 |
|------|------|-----------|
| `.docx` `.docm` `.dotx` `.dotm` | flow | 完整编辑（CE 可编辑模型 + OOXML 写回） |
| `.doc` | flow | 打开时先转 OOXML 再走 docx 引擎；**写回产出 `.docx`**（不做二进制 `.doc` 回写，且转换有损） |
| `.rtf` | flow | 打开时先转 OOXML 包再走 docx 引擎（与 `.doc` 同一条链）；**写回产出 `.docx`**（仓库没有 RTF 序列化器，且转换有损：只有字符/段落属性、表格文本、页眉页脚、图片与页面尺寸会被搬过去） |
| `.azw3` | flow | 打开时先转 OOXML 包再走 docx 引擎（Kindle KF8 正文的 `<h1>..<h6>` 转成 `w:outlineLvl`，内嵌目录落到**既有**大纲面板，不新增端点）；**写回产出 `.docx`**（仓库没有 KF8 序列化器，且转换有损：只有字符/段落属性、标题层级、列表、表格与图片会被搬过去）；老式 MOBI7 与受 DRM 保护的电子书直接拒开并给出可行动的中文说明 |
| `.epub` | flow | **原生打开**——与 `.doc` / `.rtf` / `.azw3` / `.mobi` 不同，**不走 OOXML 中转**（打开期零 OOXML 字节）：`epub.ContentDoc` 按 spine 顺序把章节 XHTML 合并成 `ContentDoc`（图片先内联、再按作者 CSS 内联样式、样式表按 `@font-face`/字体族原样带过），再复用 docx 的流式引擎。**目录优先用原书 nav/NCX**（按条目标题在编辑后的正文里顺序定位，匹配不到退回该章开头；没有作者目录才退回正文标题层级）；**写回产出 `.epub`**（`celement.BuildXHTMLDocument` + `epub.WriteDocument` 重建 OCF：mimetype 首项 Store / 逐章 XHTML / nav.xhtml / 原作者有 NCX 时补 toc.ncx / OPF 带 `properties="nav"`；原包的 CSS、字体、图片**按原路径原样搬运**，样式与包内 `url()` 引用关系都不破）；时间戳与条目顺序固定 ⇒ 二次保存逐字节一致；`META-INF/encryption.xml` 标记的 DRM 电子书直接拒开并给出可行动的中文说明 |
| `.fb2` | flow | **原生打开**——与 `.doc` / `.rtf` / `.azw3` / `.mobi` 不同，**不走 OOXML 中转**（打开期零 OOXML 字节）：FB2 是**单文件 XML**（无 zip 容器层），`fb2.ContentDoc` 按 `<body>` 里 section 的顺序把正文合并成 `ContentDoc`（`<binary>` 里的 base64 图片解成内联图片、`<image l:href="#id"/>` 按 id 取回，`<style>` 里的样式原样带过），再复用 docx 的流式引擎。**目录取自正文 section/`<title>` 层级**（标题写成 `<p><strong>卷一…</strong></p>` 这种也认；无 id 的标题由包内合成 `fb2hN` 锚点，否则条目点不动）；**写回产出 `.fb2`**（`fb2.WriteDocument` 重建 XML：`<description>` 元数据、`<binary>` 图片、正文 section 层级**原样带走**；固定元素顺序 + 稳定 id ⇒ 二次保存逐字节一致）；XML 声明里的编码按 `charsetReader` 支持 UTF-8 / Windows-1251 / KOI8-R / Windows-1252 / GBK / Big5 |
| `.mobi` `.prc` `.azw` | flow | 打开时先转 OOXML 包再走 docx 引擎（与 `.azw3` 同一条链、共用 docx 引擎）。目录**不跑整套** `mobi.prepareReaderContent`（它会把正文重写成阅读器形态），改成只补一小步 `mobi/toc_links.go` 的 `materializeTOCHyperlinks`：用 NCX 标签认出正文标题块 → 就地包成 `<h1>..<h6 id="fileposN">`，再把内嵌目录页的 `<a filepos=N>`（**没有 href**，跳转目标只是 PalmDOC 字节偏移）补成 `<a href="#fileposN">`。两端同源后目录条目在编辑器里可点：`#fileposN` 交给既有大纲面板的 `jumpToAnchor`（按 `locator.nodeId` 匹配同一个书签名）。可见文字一字不增删；**写回产出 `.docx`**（仓库没有 MOBI 序列化器，且转换有损：只有字符/段落属性、标题层级、列表、表格与图片会被搬过去）；PalmDOC 头 `Encryption != 0` 的 DRM 电子书直接拒开并给出可行动的中文说明 |
| `.lit` | flow | **原生打开**——与 `.doc` / `.rtf` / `.azw3` / `.mobi` 不同，**不走 OOXML 中转**：LIT 是**二进制容器**（`ITOLITLS` 二级头 + IFCM/AOLL 目录 + 分节 LZX 压缩块），`lit.ContentDoc` 把章节的**二进制标记流**按锚点与 atom 表反编译成 XHTML（`<img src>` 已接 `SetImageRawResolver`），再复用 docx 的流式引擎。**目录取自 OPF spine**（容器里的 `manifest`/`spine` 与 `.opf` 同构），目录条目的**跨章链接**改写成全书唯一的文档内锚点（`#lit-cN`，见 `lit/anchors.go`）⇒ 点目录在文内跳转而不是开新页 404；**在线书**（正文图片全是 CDN 绝对 URL、容器里只有一张封面）默认**联网取图**（`cmd/editor_lit_remote.go`：8 路并发 + `<UserCacheDir>/leafdoc/lit-images/` 内容寻址磁盘缓存，超时 30s / 单图上限 12MB / `image.DecodeConfig` 拦「200 但返回 HTML」），取不到的图才退回原书自带的 `alt` 占位名（`img_003` 这类），**打开从不因取图失败而失败**；`LEAFDOC_LIT_NO_REMOTE_IMAGES=1` 关掉联网（纯离线），`LEAFDOC_LIT_DEBUG=1` 打印取图耗时与成败统计；**写回产出 `.docx`**（仓库没有 LIT 序列化器，且转换有损：只有字符/段落属性、标题层级、列表、表格与图片会被搬过去）；受 DRM 保护（DES/LZX transform）的电子书直接拒开并给出可行动的中文说明 |
| `.xlsx` 等表格族、`.xls` | grid | 只读打开（无 CE 模型；保存按 `no_edit_model` 拒绘）。`.xls` 同为服务端先转 OOXML 工作簿再走网格引擎：转换有损（字体/颜色/边框/数字格式/图表/图片不保留），加密工作簿因无口令输入口而直接拒开 |
| `.csv` | grid | 只读打开（无 CE 模型；保存按 `no_edit_model` 拒绘）。服务端先按 RFC 4180 解析（合法 UTF-8 直读、否则按 GBK；首行制表符多于逗号时按 TSV），再转 OOXML 工作簿走同一个网格引擎：单元格一律是文本值（无公式/样式/图表），一张表 = 一页 |
| `.pdf` | fixed | 只读 + 批注：改版式类操作被拒（`tier_forbidden`），批注 / 表单 / 导出放行 |
| `.djvu` `.djv` | fixed | 只读：一页 = 整页位图（`display.OpImage`，IW44/JB2 解码） + OCR 文本层几何（可选中/复制，**不**另落 `OpText`）。页数与页盒在打开即精确（`INFO` 尺寸 / DPI）。DjVu 没有可增量改写的对象层 ⇒ 不提供 `annotatorSource` / `formSource` / `saverSource`：批注与保存按未接入明确拒绝（保存回 409 `no_edit_model`），不会静默丢改动 |
| `.cbz` | fixed | 只读：一页 = 整页位图（`display.OpImage`，JPEG 成员走 `display.EncodedAsset` 零转码直通，其余走 `image.Decode`）。页数与页盒在打开即精确（PNG/JPEG/GIF/BMP/TIFF/WebP **文件头**，`px*72/150`，只缩不放地收进 A4）。页序是 **自然序**（`page2` 在 `page10` 之前），不信 ZIP 中央目录顺序；非图片成员 / `__MACOSX/` / 隐藏文件一律排除，**没有** `OpText`（漫画页上的字是位图的一部分）。同样不提供 `annotatorSource` / `formSource` / `saverSource`：ZIP 里根本没有可增量改写的对象层（保存回 409 `no_edit_model`）。各页规则与 `.cbr` 共用 `comic` 包，两个引擎保持同一把尺 |
| `.cbr` | fixed | 只读：一页 = 整页位图，与 `.cbz` 同形（共用 `comic` 页规则与 `display.OpImage` 路径），差别全在 **RAR 容器**本身：①**页序**同样自然序（不信 RAR 头里的成员顺序）；②**非固实包**（WinRAR `-s-`）可逐成员随机访问，页盒在打开即精确；③**固实包**（WinRAR 默认 `-s`）只能顺序解压（跳到第 N 页要把前面全解一遍），页盒只能拿第 1 页的尺寸当全书页盒——渲染走 contain 缩放，所以估小了也不会拉伸画面，只是四周会留白；`meta` 会明确报告 `solid`；④**-hp 加密头**（连文件表都加密）在无口令输入口的前提下直接拒开（`ErrEncrypted`）；⑤**分卷**（.part1.rar）需要相邻卷同在目录下，缺卷时按 `ErrOpenFailed` 报出并提示。同样无 `annotatorSource` / `formSource` / `saverSource`：RAR 只能读不能写（保存回 409 `no_edit_model`） |
| `.pptx` `.ppt` | canvas | 只读打开（画布世界暂无编辑模型；保存按 `no_edit_model` 拒绘）。`.ppt` 同为服务端先升 OOXML 演示文稿再走画布引擎，升级有损（形状/图片/母版版式/动画/备注不保留） |
| `.pdg` `.pdz` / 内容全是 `.pdg` 的 `.zip` / `.rar` | fixed | 只读：超星 PDG，一页 = 整页位图（`display.OpImage`，清晰版的 JPEG 页走 `display.EncodedAsset` 零转码直通）。**页序是阅读顺序**（`cov*` 封面 → `bok*` 书名 → `leg*` 版权 → `fow*` 前言 → `dir*` 目录 → 纯数字正文 → `att*` 附录 → `bac*` / `cov002` 封底），与印刷页码一致（`--page=N` 才对得上）；**与 cbz/cbr 不同，页规则独立实现**（不复用 `comic`：页序语义不同）。页数与页盒在打开即精确（ZIP 逐页读**文件头**；RAR 非固实包同样精确、固实包只能拿第 1 页当全书页盒）；落位用 **contain**（等比 + 居中）而非铺满，所以固实包估小了只留白、不拉伸。**坏页保留**在页表里（扫描书的页码是页的身份，静默少一页会让「第 34 页」指向第 35 页），渲染到它时如实报错。页字节只支持**清晰版**；超星**旧版 00H/02H** 编码（专有加密 + 专有 CCITT 行程编码，解它要 WASM 解码器）**明确拒绝**并给两条可行的路（`meta` 里 `decodable=false` 即该信号）。同样无 `annotatorSource` / `formSource` / `saverSource`（保存回 409 `no_edit_model`） |
| `.chm` | flow | **原生打开**——与 `.doc` / `.rtf` / `.azw3` / `.mobi` 不同，**不走 OOXML 中转**：CHM 是**已编译 HTML 帮助**（`ITSF` 容器：`ITSP` 目录页 + PMGL/PMGI 索引树 + LZX 压缩的内容区），`chm.ContentDoc` 按 `.hhc` 目录顺序逐页读主题页（外部 CSS 内联成行内样式、`<img>` 图片从容器里就地取、`<script>` 丢掉），再拼成一本书交给 docx 流式引擎。**目录优先取 `.hhc`**（`#SYSTEM` 里声明的那份，`Local` 指向不存在的页时按警告丢弃），没有 `.hhc` 才退回 `#TOPICS` / `#STRINGS` / `#URLTBL` 主题表；目录条目的**跨章链接**改写成全书唯一的文档内锚点（`#chm-cN`，见 `chm/links.go`）⇒ 点目录在文内跳转而不是开新页 404；中文正文/标题按容器声明的代码页（GBK）解码。`LEAFDOC_CHM_DEBUG=1` 打印解析过程的取词/丢页统计；**写回产出 `.docx`**（仓库没有 CHM 序列化器，CHM 的写侧除了已停产的 `hh.exe` 没有任何读者，且转换有损：只有字符/段落属性、标题层级、列表、表格与图片会被搬过去） |
| `.chw` | flow | 与 `.chm` **同容器同魔数**，是 CHM 工程编译时留下的**索引同伴文件**：只有目录/索引树、**没有内容区**（正文在同名的 `.chm` 里）。故如实报「这是索引同伴文件，正文请打开同名的 `.chm`」，而不是假装打不开 |
| `.md` `.html` | — | 已声明但引擎未接入：给出「缺哪个世界的引擎」并回落 HTML 预览 |

档位（`tier`）语义 —— 等级是**操作集的包含关系**，不是强弱排序（`cellEdit` 与 `objectEdit`
互不包含，但都严格包含于 `fullEdit`）：

| `tier` | 中文标签 | 含义 | 当前落在这一档的格式 |
|--------|---------|------|--------------------|
| `fullEdit` | 完全编辑 | 正文、段落、布局、页、媒体、批注都可改 | `.docx` `.docm` `.dotx` `.dotm`、`.doc`、`.rtf`、`.azw3`、`.epub`、`.fb2`、`.mobi` `.prc` `.azw`、`.lit`、`.chm` |
| `cellEdit` | 单元格编辑 | 单元格值 / 公式可改，工作表结构不可改 | `.xlsx` `.xls` `.csv`（引擎已接入，但只读打开：无 CE 模型，保存按 `no_edit_model` 拒绝） |
| `objectEdit` | 对象编辑 | 对象几何 / 属性 / 对象内文字可改，文档流不可改 | `.pptx` `.ppt`（引擎已接入，但只读打开：无 CE 模型，保存按 `no_edit_model` 拒绝） |
| `inkAnnotation` | 只读批注 | 只读页 + 批注 / 墨迹 / 表单填写，改不了正文与布局 | `.pdf` `.djvu` `.cbz` `.cbr` `.pdg`（引擎已接入，但 djvu / cbz / cbr / pdg 只实现了渲染：无 CE 模型、无批注/写回器 ⇒ 保存按 `no_edit_model` 拒绝） |

服务端在 `/api/editor/info` 里下发 `tier` / `tierLabel` / `world` / `engineKind` / `allowedOps`，
前端据此渲染工具条（体验优化）；**放行与拒绝一律在 HTTP 层强制**（拒绝给结构化原因，
如 `tier_forbidden`），不依赖前端隐藏按钮。

编辑器资源缺失时一律回落 HTML 预览。

```bash
leafdoc --open
leafdoc --open --input=doc.docx
leafdoc --open --readonly --input=doc.docx
```

### `--cmd=open`（HTML 预览，别名 `preview`）

缺省弹选文件框；也可 `--input`。导出到 `.cache/<md5>/`，首页就绪后打开浏览器，边渲边加载。`--no-open` 只导出不弹窗。

```bash
leafdoc --cmd=open
leafdoc --cmd=open --input=doc.docx --dpi=96
```

---

## 格式特有命令

| 格式 | 命令 | 说明 |
|------|------|------|
| doc | `export.docx` | → DOCX |
| rtf / html / markdown / text | `export.docx` / `export.epub` / `export.mobi` | → DOCX / EPUB / MOBI |
| pdf / djvu | `export.epub` / `export.mobi` | 按页分章（文本层/抽文本） |
| epub / fb2 / mobi / azw3 / lit / chm | `export.docx` | 电子书 → DOCX |
| djvu | `page.count` / `render.page` / `meta` | 页数、单页栅格、尺寸/DPI |
| cbz | `page.count` / `render.page` / `meta` | 页数、单页栅格（png/jpeg）、逐页尺寸（px + pt）；出图像素 = 页盒 × `--dpi`/72，与编辑器同一把尺（`--dpi` 默认同样 96；未被 A4 收缩的页在 `--dpi=150` 恰好回到打包进去的原始像素） |
| cbr | `page.count` / `render.page` / `meta` | 同 cbz（共用 `comic` 页规则）；`meta` 额外带 `solid`（固实压缩标记）。非固实包逐页随机访问；**固实包**（WinRAR 默认 `-s`）只能顺序解压，且整本共用第 1 页的页盒（原图尺寸无从得知），故 `Solid() == true` 时页盒是估的——渲染用 contain 缩放，估小了也不会拉伸画面 |
| pdg | `page.count` / `render.page` / `meta` | 页数、单页栅格（png/jpeg）、逐页尺寸（px + pt）与 `decodable`；`meta` 还报告容器（`zip`/`rar`/`single`）。出图像素 = 页盒 × `--dpi`/72，与编辑器同一把尺（`--dpi` 默认同样 96；未被 A4 收缩的页在 `--dpi=150` 恰好回到打包进去的原始像素）。**页序是阅读顺序**（与印刷页码一致）；`decodable=false` 就是超星旧版 00H/02H 编码或坏页 |
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
leafdoc --input=comic.cbz --cmd=render.page --page=0 --format=png --output=p0.png
leafdoc --input=comic.cbr --cmd=render.page --page=0 --format=jpeg --output=p0.jpg
leafdoc --input=book.zip --cmd=render.page --page=33 --format=png --output=p34.png
leafdoc --input=book.pdz --cmd=meta --output=meta.json
```

标 `[write]` 的写回类命令必须带 `--output`（详见 `leafdoc --help`）。

> **没有文档模型的格式**（如 `.cbz` / `.cbr` / `.pdg`）：`text` / `export.docx` / `export.html` / `export.pdf` 这些
> 需要段落或对象层的命令**不会**落到「unknown command」或 `no handler for document type`，
> 而是给一段能照做的中文说明：看整本用 `leafdoc --open --input=book.zip`，
> 单页出图用 `--cmd=render.page`。理由：页就是一张位图，
> 段落抽取与对象层导出在语义上无处落脚，「拒绝」本身就必须携带出路。
> （CBR 的两条额外出路：改名成 `.cbz`/`.zip` 的 RAR 会被识别并提示改回；
> `-hp` 加密头的包不支持——本工具没有密码输入界面。）
>
> **PDG 的额外一点**（超星特有的报错价值）：拿 `book.zip` 去开时，工具会看成员表
> 当场告诉你手里拿的到底是什么 —— 里面有 `word/document.xml` 就是改错后缀的 `.docx`，
> 只有 `bookinfo.dat` 就是下载不完整；而「页能认出来但解不开」会被单独指认为
> **超星旧版 00H/02H 编码**（不是「文件损坏」），并给出两条可行的路：
> 用超星阅读器导出 PDF，或换成 2005 年后的「清晰版」。

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
