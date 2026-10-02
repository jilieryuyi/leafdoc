# leafdoc CLI

**The AI era: a document CLI built for AI.**

A multi-format Office / PDF / ebook / metafile command-line tool. It routes by the `--input` extension (or `--type`); xlsx offers full read/write, while the other formats support extraction, rendering, export and format conversion.

English | [简体中文](README.md)

> **Self-developed · Pure Go**: parsing, layout, rendering and export are all built on a **pure Golang in-house core** — **no** WPS, **no** Microsoft Office, and **no** LibreOffice / OpenOffice or any other third-party office suite or open-source engine. Deploying and running only needs the local Go build artifact (plus system fonts and other ordinary environment setup) to do text extraction, layout rendering and PDF/HTML/SVG export.

Online help: `leafdoc --help` · `leafdoc --help <command>`

---

## Feature overview

| Capability | Description |
|------|------|
| Multi-format parsing | Word / Excel / PowerPoint (OOXML + binary), PDF, DjVu, CBZ / CBR (comic archives), PDG (Chaoxing ebooks), RTF, HTML, Markdown, plain text, CSV, EPUB, FB2 (FictionBook), MOBI, AZW3, Microsoft Reader (LIT), compiled HTML help (CHM), EMF, WMF |
| Text & media | Extract body/header/footer, paragraphs, fonts, embedded images |
| Render & export | Paginate / per-sheet / per-slide → SVG / PNG / JPG; export PDF and HTML preview (pure Go, no Office required) |
| Format conversion | See the **Format conversion table** below (`export.*` / `render`; underpinned by `document.Export` + `OutputFormatter`) |
| xlsx read/write | Cells, formulas, styles, merges, comments, filters, protection, printing, sparklines, charts, VBA, etc. |
| Document merge | `doc.merge` merges multiple files into a single DOCX |
| Embedded editor | `--open` / `--editor`: formats already wired to the layout engine (`.docx` / `.doc` / `.rtf` / `.azw3` / `.epub` / `.fb2` / `.mobi` / `.lit` / `.chm` / `.xlsx` / `.xls` / `.csv` / `.pdf` / `.djvu` / `.cbz` / `.cbr` / `.pdg` / `.pdz` / `.pptx` / `.ppt`) open the embedded editor directly (`#hash=` multi-document session, supports `--readonly`); other declared formats explain "which world's engine is missing" and fall back to the HTML preview |
| Browser preview | `--cmd=open` (alias `preview`) / the welcome page started with no arguments: render-as-you-write HTML, page through over local HTTP |
| Welcome page | Starting with no arguments opens it automatically: a single "Open file…" button on the page → native dialog → launches `leafdoc --open --readonly --input=<file>` with the same executable (i.e. the read-only editor) |

---

## Usage

```text
leafdoc                              # welcome page (auto-opens browser) + HTTP + REPL
leafdoc --input <file> --cmd <command> [--output <path>] [options...]
leafdoc --open [--readonly] [--input <file>]
leafdoc --cmd=open [--input <file>] [--option=canvas|code]
leafdoc --cmd=create --output=new.xlsx
leafdoc --cmd=doc.merge --input=a.docx --value=b.xlsx --output=out.docx
leafdoc --help
leafdoc --help <command>
```

### Welcome page (no arguments)

Started with no arguments, it automatically opens the browser to the welcome page, which has exactly one action: "Open file…":

- Click the button → `POST /welcome/open` → the **server** opens the native file dialog (Windows: PowerShell OpenFileDialog)
- Once selected, it launches a separate process with the same executable: `leafdoc --open --readonly --input=<selected file>`
  - The child process detaches from the welcome process (Windows `CREATE_NO_WINDOW`, no black console flash) and opens the read-only editor in a new tab
  - In other words, "enter the embedded editor's read-only mode without remembering any command line"; options/behavior are identical to typing that command by hand
- The page also lists: what was opened this session (server-side record, up to 12 entries) and the HTML-preview documents already opened by this process

### Interactive mode (the same process's REPL)

Besides the welcome page, the same HTTP service keeps the original commands:

- REPL: `/open` (pick a file), `/open?path=...`, `/list`, `/help`, `/quit`
- Browser: `http://127.0.0.1:<port>/open?path=C:/file.pdf`
- Welcome-page JSON status: `http://127.0.0.1:<port>/?format=json`
- Previews are written to `.cache/<md5>/`, multi-document is mounted at `/v/<md5>/`

### Global options

| Option | Description |
|------|------|
| `--input` | Input file (`create` / `doc.merge` / `help` / `--open` may omit it) |
| `--output` | Output file or directory |
| `--cmd` | Command (repeatable or comma-separated; xlsx can be chained) |
| `--password` | Decrypt / encrypt password |
| `--type` | Force a type: `docx` `doc` `pptx` `ppt` `xlsx` `xls` `pdf` `djvu` `cbz` `cbr` `pdg` `rtf` `html` `csv` `markdown` `text` `emf` `wmf` `epub` `fb2` `mobi` `azw3` `lit` `chm` |
| `--page` | Page / slide index (0-based) |
| `--filter` | Text filter: `all` / `document` / `header` / `footer` (combinable with `+`) |
| `--pages` | PDF/DjVu text page range (0-based): `0,2` or `0-2` (CBZ / CBR / PDG have no text layer and don't consume this switch; use `--page` to render a single page) |
| `--format` / `--dpi` | Render format and DPI |
| `--all-in-one` | `render`: streaming formats output a single continuous long image (docx/doc/html/md/rtf) |
| `--option` | HTML: `canvas` (default) \| `code` (Markdown semantic HTML); docx: `ast` (layout through the Document AST; `--debug` / `render` / `export.html`) |
| `--max-width` / `--max-height` | EMF/WMF size limits |
| `--disable-formula` | xlsx: skip formula calculation |
| `--open` | Open the document in the embedded editor (pops a file picker when `--input` is missing; formats not wired to an engine fall back to the HTML preview) |
| `--editor` | Embedded editor (`--readonly` is a pure canvas, read-only viewer); `.docx` / `.doc` / `.rtf` / `.azw3` / `.epub` / `.fb2` / `.mobi` / `.lit` / `.chm` / `.xlsx` / `.xls` / `.csv` / `.pdf` / `.djvu` / `.cbz` / `.cbr` / `.pdg` / `.pdz` / `.pptx` / `.ppt` are already wired to an engine |

### Type detection

| Extension | Component |
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

> `.mobi` `.prc` `.azw` all hook into the same mobi engine; but **`.azw` / `.prc` are "one extension, two containers"**
> (early Amazon shipped both old-style MOBI7 and KF8 in disguise): these two extensions decide the family
> **by content, not by name** — if the content is KF8 the azw3 engine takes over (`cmd/detect.go::ambiguousKindleDocType`,
> shared by the CLI and the editor), otherwise it stays mobi. `.mobi` / `.azw3` are semantically unambiguous and don't take part in this override.

> `.cbz` (Comic Book Zip) is decided by **extension only**, never by bytes: a ZIP container doesn't introduce itself,
> and `sniffType` cannot tell "a pack of comic images" from "a pack of docx parts" (see the comment in `cmd/detect.go::sniffType`).
> Correspondingly, renaming a `.docx` to `.cbz` will be rejected by the cbz parser with a hint about the real format.
>
> `.cbr` (Comic Book Rar), on the other hand, **accepts both signals**: RAR has a unique magic number (`Rar!\x1a\x07`),
> so `sniffType` can recognize it before the ZIP/text probes (out in the wild `.cbr` is often renamed to `.cbz`/`.zip`,
> or even has no extension at all); but the extension still wins, so when a RAR renamed to `.cbz` fails in the cbz engine,
> it tells you outright "this is actually a RAR, please rename it back to `.cbr`" instead of leaving you with an
> opaque "not a ZIP".
>
> `.pdg` / `.pdz` (Chaoxing PDG) has one more path than the two above: **the most common distribution form of a PDG book
> is a plain `book.zip`** (that's what Chaoxing Reader's "packaged download" hands you), with no clue in the extension.
> So besides the extension, pdg has an extra **content adjudication**: `sniffType` looks at the ZIP/RAR member table, and
> **if it's all `.pdg` (`pdg` member count > 0 and not fewer than the image member count), it opens as PDG** (`pdg.IsPDGArchive`);
> when it really isn't PDG it doesn't steal it — if the ratio isn't met it falls back to the original ZIP probe (docx/epub/…) or an error.
> Single-page files only have the `.pdg` / `.pdz` extensions. Both failures come with actionable instructions: if it contains
> `word/document.xml`, it tells you to rename it back to `.docx`; if it contains images, it tells you to rename it back to
> `.cbz` / `.cbr`; a lone `bookinfo.dat` means the download was incomplete.

---

## Format conversion table (From → To)

**Legend**: ✓ = the CLI has a matching export command and the implementation supports it · — = unsupported or no command provided. For the exact commands see "Common commands" and "Format-specific commands" below.

### Documents / slides / spreadsheets

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

**Fixed-page formats** (pdf / djvu / cbz / cbr / pdg / pptx / ppt): cross-format export goes through **display elements** (text boxes, images, rectangles), not full-page PNG rasterization. When ppt has no draw pipeline it is first upgraded to pptx.

### Metafiles (EMF / WMF)

| From | → SVG | → PNG/JPG |
|------|:-----:|:---------:|
| **emf** | ✓ | ✓ |
| **wmf** | ✓ | ✓ |

### Multi-file merge

| From | → DOCX |
|------|:------:|
| docx + xlsx / csv / … | ✓ |

### Common aliases

| Target | Command aliases |
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
| Bitmap | `to.image` · `to.png` · `export.image` |

### Go API (mirrors the CLI)

```go
document.Export(doc, pdf.NewPDFFormatter(w))
document.Export(doc, epub.NewEPUBFormatter(w, "Title"))
document.Export(doc, mobi.NewMOBIFormatter(w, "Title"))
document.Export(doc, docx.NewDOCXFormatter(w))   // needs a DocxExporter
document.Export(doc, pptx.NewPPTXFormatter(w))   // needs a PptxExporter
document.Export(doc, document.NewHTMLFormatter(dir, opts...))
document.Export(doc, document.NewRenderFormatter(dir, "svg", document.WithAllInOne(true)))
```

---

## Common commands

Applicable to most `document.Office` formats (docx / doc / pptx / ppt / xlsx / xls / pdf / djvu / cbz / cbr / pdg / rtf / html / markdown / text / csv / epub / fb2 / mobi / azw3 / lit / chm):

| Command | Description |
|------|------|
| `info` / `mime` / `encrypted` | Metadata |
| `text` / `paragraphs` / `fonts` / `images` | Extraction |
| `render` | Render all pages/sheets/slides to a directory |
| `export.pdf` / `export.html` / `export.docx` / `export.epub` / `export.mobi` | Export PDF / HTML / DOCX / EPUB / MOBI (see the **format conversion table**) |
| `page.count` / `render.page` | Page count / single page (docx, doc, pdf, djvu, cbz, cbr, pdg, pptx) |
| `doc.merge` | Merge multiple files into DOCX |

```bash
leafdoc --input=a.docx --cmd=text --filter=document --output=a.txt
leafdoc --input=a.pdf --cmd=text --pages=0-1 --output=p01.txt
leafdoc --input=a.pdf --cmd=render --format=svg --output=./pages
leafdoc --input=note.md --cmd=render --format=png --all-in-one --output=./out
leafdoc --input=a.docx --cmd=export.pdf --output=a.pdf
leafdoc --input=a.pptx --cmd=render.page --page=0 --format=png --output=s0.png
leafdoc --cmd=doc.merge --input=a.docx --value=b.xlsx,c.csv --output=merged.docx
```

Markdown's `export.html` / `--cmd=open` additionally support `--option`:

| `--option` | Description |
|------------|------|
| `canvas` (default) | ContentDoc layout → HTML Canvas preview directory |
| `code` | CommonMark/GFM → semantic HTML (a file `.html` or `index.html` under a directory) |
| `ast` | docx: lay out the Layer-3 Document AST (ECMA-376 `w:body`) before rendering, benchmarked against the default Element path |

```bash
leafdoc --input=note.md --cmd=export.html --output=./html
leafdoc --input=note.md --cmd=export.html --option=code --output=note.html
leafdoc --cmd=open --input=note.md --option=code --no-open
leafdoc --debug --input=doc.docx --jpg --option=ast --auto-exit
leafdoc --input=doc.docx --cmd=render --format=png --option=ast --output=./out
```

### Mapping to `document.Office` options

| CLI | Go API |
|-----|--------|
| `--format=png` | `document.WithFormat("png")` |
| `--all-in-one` | `document.WithAllInOne(true)` |
| `--option=code` | `document.WithHTMLCode()` / `WithHTMLMode("code")` |
| `--option=ast` | `document.WithAST(true)` / `docx.File.SetLayoutUseAST(true)` |
| `--filter=document+header` | `WithTextDocument()` + `WithTextHeader()` |
| `--pages=0-1` | `document.WithTextPages(0, 1)` |
| `--dpi=96` | `render.WithHTMLDPI(96)` |

### `--open` (embedded editor)

For docx it pops a file picker by default (or use `--input`) and enters the embedded editor directly (same runtime as `--editor`): the URL carries `#hash=<sid>`, a refresh restores it, and changing the hash switches files without reloading the page. `--readonly` is a pure canvas, read-only view (no toolbar/footer bar, text selectable and copyable).

Admission is decided by the capability matrix, giving the same answer as the CLI / HTTP:

| Format | World | Editor behavior |
|------|------|-----------|
| `.docx` `.docm` `.dotx` `.dotm` | flow | Full editing (CE editable model + OOXML write-back) |
| `.doc` | flow | Converted to OOXML on open, then the docx engine; **write-back produces `.docx`** (no binary `.doc` write-back, and the conversion is lossy) |
| `.rtf` | flow | Converted to an OOXML package on open, then the docx engine (same chain as `.doc`); **write-back produces `.docx`** (the repo has no RTF serializer, and the conversion is lossy: only character/paragraph properties, table text, headers/footers, images and page size are carried over) |
| `.azw3` | flow | Converted to an OOXML package on open, then the docx engine (Kindle KF8 body `<h1>..<h6>` becomes `w:outlineLvl`, and the embedded TOC lands in the **existing** outline panel, no new endpoint); **write-back produces `.docx`** (the repo has no KF8 serializer, and the conversion is lossy: only character/paragraph properties, heading levels, lists, tables and images are carried over); old MOBI7 and DRM-protected ebooks are rejected outright with an actionable Chinese explanation |
| `.epub` | flow | **Opens natively** — unlike `.doc` / `.rtf` / `.azw3` / `.mobi`, it **does not go through OOXML conversion** (zero OOXML bytes at open time): `epub.ContentDoc` merges the chapter XHTML in spine order into a `ContentDoc` (images inlined first, then author CSS inlined, stylesheets carried over as-is by `@font-face`/font family), then reuses the docx streaming engine. **The TOC prefers the book's own nav/NCX** (locating each entry by title order in the edited body, falling back to the chapter start when there's no match; only when the author's TOC is absent does it fall back to body heading levels); **write-back produces `.epub`** (`celement.BuildXHTMLDocument` + `epub.WriteDocument` rebuilds the OCF: mimetype first item Store / per-chapter XHTML / nav.xhtml / toc.ncx added when the author had an NCX / OPF with `properties="nav"`; the original package's CSS, fonts and images are **carried over at their original paths**, so styles and in-package `url()` references stay intact); timestamps and entry order are fixed ⇒ a second save is byte-for-byte identical; DRM ebooks marked by `META-INF/encryption.xml` are rejected outright with an actionable Chinese explanation |
| `.fb2` | flow | **Opens natively** — unlike `.doc` / `.rtf` / `.azw3` / `.mobi`, **no OOXML conversion** (zero OOXML bytes at open time): FB2 is a **single XML file** (no zip container layer), and `fb2.ContentDoc` merges the body in section order into a `ContentDoc` (base64 images in `<binary>` are decoded into inline images, `<image l:href="#id"/>` is resolved back by id, and styles in `<style>` are carried over as-is), then reuses the docx streaming engine. **The TOC comes from the body section/`<title>` hierarchy** (titles written as `<p><strong>Volume One…</strong></p>` are recognized too; titles without an id get a package-synthesized `fb2hN` anchor, otherwise the entry isn't clickable); **write-back produces `.fb2`** (`fb2.WriteDocument` rebuilds the XML: `<description>` metadata, `<binary>` images and the body section hierarchy **carried over as-is**; fixed element order + stable ids ⇒ a second save is byte-for-byte identical); the encoding in the XML declaration is supported via `charsetReader` for UTF-8 / Windows-1251 / KOI8-R / Windows-1252 / GBK / Big5 |
| `.mobi` `.prc` `.azw` | flow | Converted to an OOXML package on open, then the docx engine (same chain as `.azw3`, sharing the docx engine). The TOC **does not run the whole** `mobi.prepareReaderContent` (which rewrites the body into reader form); instead only a small step from `mobi/toc_links.go`, `materializeTOCHyperlinks`, is added: NCX labels identify body heading blocks → wrapped in place as `<h1>..<h6 id="fileposN">`, then `<a filepos=N>` in the embedded TOC page (**no href**, the jump target is only a PalmDOC byte offset) is augmented to `<a href="#fileposN">`. With both ends matching, TOC entries become clickable in the editor: `#fileposN` goes to the existing outline panel's `jumpToAnchor` (matching the same bookmark by `locator.nodeId`). Not a single visible character is added or removed; **write-back produces `.docx`** (the repo has no MOBI serializer, and the conversion is lossy: only character/paragraph properties, heading levels, lists, tables and images are carried over); ebooks with PalmDOC header `Encryption != 0` DRM are rejected outright with an actionable Chinese explanation |
| `.lit` | flow | **Opens natively** — unlike `.doc` / `.rtf` / `.azw3` / `.mobi`, it **does not go through OOXML conversion**: LIT is a **binary container** (`ITOLITLS` secondary header + IFCM/AOLL directory + LZX-compressed section blocks), and `lit.ContentDoc` decompiles the chapter's **binary markup stream** into XHTML using anchors and the atom table (`<img src>` is already wired to `SetImageRawResolver`), then reuses the docx streaming engine. **The TOC comes from the OPF spine** (the container's `manifest`/`spine` is isomorphic to `.opf`), and TOC entries' **cross-chapter links** are rewritten to a document-unique in-book anchor (`#lit-cN`, see `lit/anchors.go`) ⇒ clicking the TOC jumps within the document instead of opening a new page that 404s; **online books** (body images are all CDN absolute URLs, with only a cover in the container) **fetch images over the network by default** (`cmd/editor_lit_remote.go`: 8-way concurrency + `<UserCacheDir>/leafdoc/lit-images/` content-addressed disk cache, 30s timeout / 12MB per image / `image.DecodeConfig` catches "200 but returns HTML"), and only images that can't be fetched fall back to the book's own `alt` placeholder names (like `img_003`); **opening never fails because image fetching failed**; `LEAFDOC_LIT_NO_REMOTE_IMAGES=1` disables networking (fully offline), `LEAFDOC_LIT_DEBUG=1` prints fetch timing and success/failure stats; **write-back produces `.docx`** (the repo has no LIT serializer, and the conversion is lossy: only character/paragraph properties, heading levels, lists, tables and images are carried over); DRM-protected ebooks (DES/LZX transform) are rejected outright with an actionable Chinese explanation |
| `.xlsx` family, `.xls` | grid | Read-only open (no CE model; saving is rejected with `no_edit_model`). `.xls` likewise converts server-side to an OOXML workbook, then the grid engine: the conversion is lossy (fonts/colors/borders/number formats/charts/images are not preserved), and encrypted workbooks are rejected outright because there is no password input |
| `.csv` | grid | Read-only open (no CE model; saving is rejected with `no_edit_model`). The server first parses per RFC 4180 (valid UTF-8 read directly, otherwise GBK; if the first row has more tabs than commas, treated as TSV), then converts to an OOXML workbook and runs the same grid engine: cells are plain text values (no formulas/styles/charts), one sheet = one page |
| `.pdf` | fixed | Read-only + annotation: layout-changing operations are rejected (`tier_forbidden`), while annotations / forms / export are allowed |
| `.djvu` `.djv` | fixed | Read-only: one page = a full-page bitmap (`display.OpImage`, IW44/JB2 decoding) + OCR text-layer geometry (selectable/copyable, but **not** separately emitted as `OpText`). Page count and page box are exact at open time (`INFO` size / DPI). DjVu has no incrementally rewritable object layer ⇒ no `annotatorSource` / `formSource` / `saverSource`: annotation and saving are explicitly rejected as not wired (saving returns 409 `no_edit_model`), so changes are never silently dropped |
| `.cbz` | fixed | Read-only: one page = a full-page bitmap (`display.OpImage`; JPEG members pass through zero-transcode via `display.EncodedAsset`, the rest go through `image.Decode`). Page count and page box are exact at open time (PNG/JPEG/GIF/BMP/TIFF/WebP **file headers**, `px*72/150`, only shrinking — never enlarging — to fit A4). Page order is **natural order** (`page2` before `page10`), not trusting the ZIP central directory order; non-image members / `__MACOSX/` / hidden files are all excluded, and there is **no** `OpText` (the words on a comic page are part of the bitmap). It likewise provides no `annotatorSource` / `formSource` / `saverSource`: a ZIP simply has no incrementally rewritable object layer (saving returns 409 `no_edit_model`). The per-page rules are shared with `.cbr` via the `comic` package, so both engines use the same yardstick |
| `.cbr` | fixed | Read-only: one page = a full-page bitmap, shaped like `.cbz` (sharing `comic` page rules and the `display.OpImage` path); all the differences lie in the **RAR container** itself: ① **page order** is likewise natural order (not trusting the order in the RAR header); ② a **non-solid archive** (WinRAR `-s-`) allows per-member random access, so the page box is exact at open time; ③ a **solid archive** (WinRAR's default `-s`) can only be decompressed sequentially (jumping to page N requires decompressing everything before it), so the page box can only use page 1's size as the whole book's page box — rendering uses contain scaling, so an underestimate never stretches the image, it just leaves whitespace around it; `meta` explicitly reports `solid`; ④ **-hp encrypted header** (encrypting even the file table) is rejected outright with no password input (`ErrEncrypted`); ⑤ **multi-volume** (.part1.rar) needs the adjacent volumes in the same directory, and a missing volume is reported as `ErrOpenFailed` with a hint. It likewise has no `annotatorSource` / `formSource` / `saverSource`: RAR is read-only (saving returns 409 `no_edit_model`) |
| `.pptx` `.ppt` | canvas | Read-only open (the canvas world has no editing model yet; saving is rejected with `no_edit_model`). `.ppt` likewise converts server-side to an OOXML presentation then the canvas engine, and the upgrade is lossy (shapes/images/master layouts/animations/notes are not preserved) |
| `.pdg` `.pdz` / `.zip` / `.rar` whose content is all `.pdg` | fixed | Read-only: Chaoxing PDG, one page = a full-page bitmap (`display.OpImage`; clear-version JPEG pages pass through zero-transcode via `display.EncodedAsset`). **Page order is reading order** (`cov*` cover → `bok*` book title → `leg*` copyright → `fow*` foreword → `dir*` TOC → pure-number body → `att*` appendix → `bac*` / `cov002` back cover), matching the printed page numbers (so `--page=N` lines up); **unlike cbz/cbr, the page rules are implemented independently** (the `comic` package isn't reused: the page-order semantics differ). Page count and page box are exact at open time (ZIP reads each page's **file header**; non-solid RAR is equally exact, while a solid archive can only use page 1 as the whole book's page box); placement uses **contain** (proportional + centered) rather than fill, so an underestimated solid archive only leaves whitespace and doesn't stretch. **Bad pages are kept** in the page table (for a scanned book the page number is the page's identity; silently dropping one would make "page 34" point to page 35), and rendering one reports an error as it is. Page bytes support only the **clear version**; Chaoxing's **old 00H/02H** encoding (proprietary encryption + proprietary CCITT run-length coding, needing a WASM decoder to handle) is **explicitly rejected** with two viable paths (the `decodable=false` signal in `meta` is that signal). It likewise has no `annotatorSource` / `formSource` / `saverSource` (saving returns 409 `no_edit_model`) |
| `.chm` | flow | **Opens natively** — unlike `.doc` / `.rtf` / `.azw3` / `.mobi`, **no OOXML conversion**: CHM is **compiled HTML help** (`ITSF` container: `ITSP` directory page + PMGL/PMGI index tree + LZX-compressed content area), and `chm.ContentDoc` reads topic pages in `.hhc` TOC order (external CSS inlined into inline styles, `<img>` images taken in place from the container, `<script>` dropped), then assembles them into a book and hands it to the docx streaming engine. **The TOC prefers `.hhc`** (the one declared in `#SYSTEM`; pages whose `Local` points nowhere are dropped with a warning), and only without `.hhc` does it fall back to the `#TOPICS` / `#STRINGS` / `#URLTBL` topic table; TOC entries' **cross-chapter links** are rewritten to a document-unique in-book anchor (`#chm-cN`, see `chm/links.go`) ⇒ clicking the TOC jumps within the document instead of opening a new page that 404s; Chinese body text/headings are decoded per the container's declared code page (GBK). `LEAFDOC_CHM_DEBUG=1` prints the parser's word-lookup/dropped-page stats; **write-back produces `.docx`** (the repo has no CHM serializer, CHM's write side has no readers besides the discontinued `hh.exe`, and the conversion is lossy: only character/paragraph properties, heading levels, lists, tables and images are carried over) |
| `.chw` | flow | The **same container and magic number** as `.chm`; it's the **index companion file** left when a CHM project is compiled: only the TOC/index tree, **no content area** (the body is in the same-named `.chm`). So it truthfully reports "this is an index companion file, please open the same-named `.chm` for the body" rather than pretending it can't be opened |
| `.md` `.html` | — | Declared but not wired to an engine: explains "which world's engine is missing" and falls back to the HTML preview |

Tier (`tier`) semantics — the levels are an **inclusion relation over operation sets**, not a strength ranking (`cellEdit` and `objectEdit` don't contain each other, but both are strictly contained in `fullEdit`):

| `tier` | Label | Meaning | Formats currently at this tier |
|--------|---------|------|--------------------|
| `fullEdit` | Full edit | Body, paragraphs, layout, pages, media and annotations are all editable | `.docx` `.docm` `.dotx` `.dotm`, `.doc`, `.rtf`, `.azw3`, `.epub`, `.fb2`, `.mobi` `.prc` `.azw`, `.lit`, `.chm` |
| `cellEdit` | Cell edit | Cell values / formulas are editable, worksheet structure is not | `.xlsx` `.xls` `.csv` (engine wired, but opened read-only: no CE model, saving rejected with `no_edit_model`) |
| `objectEdit` | Object edit | Object geometry / properties / text inside objects are editable, the document flow is not | `.pptx` `.ppt` (engine wired, but opened read-only: no CE model, saving rejected with `no_edit_model`) |
| `inkAnnotation` | Read-only + annotation | Read-only pages + annotations / ink / form filling; body and layout can't be changed | `.pdf` `.djvu` `.cbz` `.cbr` `.pdg` (engine wired, but djvu / cbz / cbr / pdg only implement rendering: no CE model, no annotator/writer ⇒ saving rejected with `no_edit_model`) |

On the server, `/api/editor/info` delivers `tier` / `tierLabel` / `world` / `engineKind` / `allowedOps`, and the frontend renders the toolbar accordingly (a UX nicety); **allow/deny is always enforced at the HTTP layer** (denials return a structured reason such as `tier_forbidden`), never relying on the frontend hiding buttons.

When editor resources are missing, it always falls back to the HTML preview.

```bash
leafdoc --open
leafdoc --open --input=doc.docx
leafdoc --open --readonly --input=doc.docx
```

### `--cmd=open` (HTML preview, alias `preview`)

Pops a file picker by default; or use `--input`. Exports to `.cache/<md5>/` and opens the browser once the first page is ready, loading while rendering. `--no-open` exports without opening a window.

```bash
leafdoc --cmd=open
leafdoc --cmd=open --input=doc.docx --dpi=96
```

---

## Format-specific commands

| Format | Command | Description |
|------|------|------|
| doc | `export.docx` | → DOCX |
| rtf / html / markdown / text | `export.docx` / `export.epub` / `export.mobi` | → DOCX / EPUB / MOBI |
| pdf / djvu | `export.epub` / `export.mobi` | Chapter per page (text layer / extracted text) |
| epub / fb2 / mobi / azw3 / lit / chm | `export.docx` | Ebook → DOCX |
| djvu | `page.count` / `render.page` / `meta` | Page count, single-page raster, size/DPI |
| cbz | `page.count` / `render.page` / `meta` | Page count, single-page raster (png/jpeg), per-page size (px + pt); output pixels = page box × `--dpi`/72, the same yardstick as the editor (`--dpi` also defaults to 96; pages not shrunk to A4 return to exactly the original packaged pixels at `--dpi=150`) |
| cbr | `page.count` / `render.page` / `meta` | Same as cbz (sharing `comic` page rules); `meta` additionally carries `solid` (solid-compression flag). Non-solid archives allow per-page random access; a **solid archive** (WinRAR's default `-s`) can only be decompressed sequentially, and the whole book shares page 1's page box (the original size is unknowable), so when `Solid() == true` the page box is estimated — rendering uses contain scaling, so an underestimate never stretches the image |
| pdg | `page.count` / `render.page` / `meta` | Page count, single-page raster (png/jpeg), per-page size (px + pt) and `decodable`; `meta` also reports the container (`zip`/`rar`/`single`). Output pixels = page box × `--dpi`/72, the same yardstick as the editor (`--dpi` also defaults to 96; pages not shrunk to A4 return to exactly the original packaged pixels at `--dpi=150`). **Page order is reading order** (matching the printed page numbers); `decodable=false` means Chaoxing's old 00H/02H encoding, or a bad page |
| ppt | `export.pptx` / `slide.list` | → PPTX; slide text |
| xls | `export.xlsx` / `export.csv` / `sheet.list` | → XLSX / CSV |
| xlsx | `export.csv` | → CSV (first sheet by default; specify with `--sheet`) |
| csv | `export.xlsx` / `rows` | → XLSX; export rows |
| docx | `save` / `accept-revisions` / `reject-revisions` | Write-back; revisions |
| pptx | `save` / `slide.count` / `slide.export` / `slide.size` / `tags` | Write-back; split out a single slide |
| pdf | `save` / `outlines` / `hyperlinks` / `meta` | Write-back; bookmarks/links/Info |
| emf | `to.svg` / `to.image` / `to.bitmap` / `decode` / `info` | SVG / PNG / bitmap |
| wmf | `to.svg` / `to.image` / `to.bitmap` / `text` / `info` | Same as above |

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

Write-back commands marked `[write]` must carry `--output` (see `leafdoc --help` for details).

> **Formats without a document model** (such as `.cbz` / `.cbr` / `.pdg`): commands like `text` / `export.docx` / `export.html` / `export.pdf`
> that need a paragraph or object layer **will not** land on "unknown command" or `no handler for document type`;
> instead they give an actionable explanation in Chinese: to view the whole book use `leafdoc --open --input=book.zip`,
> and to render a single page use `--cmd=render.page`. The reason: a page is just a bitmap,
> paragraph extraction and object-layer export have nowhere to land semantically, and "rejection" must itself carry a way out.
> (CBR's two extra ways out: a RAR renamed to `.cbz`/`.zip` is recognized and prompted to be renamed back;
> archives with an `-hp` encrypted header are unsupported — this tool has no password input UI.)
>
> **One extra point for PDG** (Chaoxing's particular value in error reporting): when you open `book.zip`, the tool looks at the member table
> and tells you right there what you're actually holding — `word/document.xml` inside means it's a mis-renamed `.docx`,
> a lone `bookinfo.dat` means an incomplete download; and "pages recognized but undecompressable" is separately identified as
> **Chaoxing's old 00H/02H encoding** (not "corrupt file"), with two viable paths given:
> export to PDF with Chaoxing Reader, or switch to the "clear version" from 2005 onward.

---

## xlsx commands

### Workbook

`create` · `info` · `mime` · `encrypted` · `save` · `calc-chain`

### Worksheet / cells

`sheet.list` `sheet.add` `sheet.delete` `sheet.rename`  
`cell.get` `cell.set` `cell.style` `range.get` `rows.get`  
`cell.merge` `cell.unmerge` `col.width` `col.visible` `row.height` `row.visible`  
`calc`

### Text / images / names / comments

`text` `paragraphs` `fonts` `images` `picture.add`  
`name.list` `name.set` `name.delete`  
`comment.list` `comment.add` `threaded.list` `threaded.add`

### Tables / filters / links / validation / protection / printing

`table.add` `filter.set` `filter.apply` `filter.clear`  
`link.add` `validate.set` `validate.delete`  
`protect.sheet` `unprotect.sheet` `protect.workbook` `unprotect.workbook`  
`print.setup` `print.header` `print.margins` `print.options`

### Sparklines / conditional formatting / charts / styles / rendering / VBA

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

Full usage of a single command: `leafdoc --help cell.get`.

---

## Author

WeChat: `yuyi297341015`
