# HtmlToOpenXml — Complete Reference Guide (v3.5.0)

HTML → DOCX converter for .NET. Open source, MIT, maintained since 2009 (active repo: last push 2026-07-20).

| | |
|---|---|
| NuGet package | `HtmlToOpenXml.dll` |
| Repository | https://github.com/onizet/html2openxml |
| Wiki | https://github.com/onizet/html2openxml/wiki |
| Author | Olivier Nizet |
| Targets | net10.0, net8.0, netstandard2.0, net462 |
| Dependencies | `DocumentFormat.OpenXml` (3.5.1), `AngleSharp` (1.5.0), `Microsoft.Extensions.Logging.Abstractions` (6.0.0) |
| Parsing engine | AngleSharp (W3C-compliant HTML5); since v3.0, previously a custom regex parser |
| History | started 2009 to convert user comments into Word |

Note: version 3.2.6 on NuGet is marked deprecated (bad packaging, same code as 3.2.5) → use ≥ 3.5.0.

---

## 1. Quickstart

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;

using var ms = new MemoryStream();
using (var package = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
{
    var mainPart = package.MainDocumentPart
        ?? throw new InvalidOperationException();
    if (mainPart.Document == null) new Document(new Body()).Save(mainPart);

    var converter = new HtmlConverter(mainPart);
    await converter.ParseBody(html);          // appends to the Body
    mainPart.Document.Save();
}
File.WriteAllBytes("out.docx", ms.ToArray());
```

Documented alternative: open an existing template (`WordprocessingDocument.Open`) and append the conversion to it — recommended, because a template already carries styles, themes and document properties.

---

## 2. Public API — `HtmlConverter`

Constructor: `HtmlConverter(MainDocumentPart mainPart, IWebRequest? webRequester = null)` — `webRequester` is the image-download factory (default: `DefaultWebRequest`).

| Method | Signature | Behavior |
|---|---|---|
| `Parse` | `IList<OpenXmlCompositeElement> Parse(string html)` | Synchronous; returns the generated paragraphs (does not insert them) |
| `ParseAsync` | `Task<IEnumerable<OpenXmlCompositeElement>> ParseAsync(string html, CancellationToken = default)` | Same, async |
| `ParseAsync` (overload) | `Task<...> ParseAsync(string html, ParallelOptions)` | Controls download parallelism for images |
| `ParseBody` | `Task ParseBody(string html, CancellationToken = default)` | Appends to the `Body`; auto-creates the `_top` bookmark in the 1st paragraph if missing; moves `SectionProperties` to the end (required by the OpenXml schema) |
| `ParseHeader` | `Task ParseHeader(string html, HeaderFooterValues? = Default, CancellationToken = default)` | Appends to the header (creates/resolves `HeaderPart`, handles image/link relationships) |
| `ParseFooter` | `Task ParseFooter(string html, HeaderFooterValues? = Default, CancellationToken = default)` | Same for the footer |
| `RefreshStyles` | `void RefreshStyles()` | Reloads the document's style cache |

Obsolete APIs (do not use): `ParseHtml` → `ParseBody`; `Parse(html, CancellationToken)` / `Parse(html, ParallelOptions)` → `ParseAsync`; `ExcludeLinkAnchor` → `SupportsAnchorLinks` (`ExcludeLinkAnchor=true` ⇒ `SupportsAnchorLinks=false`).

---

## 3. Configuration properties

| Property | Default | Description |
|---|---|---|
| `ImageProcessing` (`ImageProcessingMode`) | `Embed` | `Embed` = download and embed all images (self-contained document, large file). `LinkExternal` = external links via relationships (small file, needs internet; **data URIs stay embedded**). `EmbedDataUriOnly` = embed only data URIs, skip external images |
| `SupportsAnchorLinks` | `true` | Enables internal document anchor links |
| `ContinueNumbering` | `true` | Consecutive `ol` continue the numbering; `false` = always restart at 1 |
| `SupportsHeadingNumbering` | `true` | Multi-level heading numbering (1., 1.1, 2.1…): detects headings starting with a number; `false` disables it |
| `AcronymPosition` | `PageEnd` | Where to render `abbr`/`acronym`: `PageEnd` (footnote) or `DocumentEnd` (document note) |
| `RenderPreAsTable` | `false` | Renders `<pre>` as a single-cell bordered table (style `TableGrid`) instead of a monospaced paragraph |
| `TableCaptionPosition` (`CaptionPositionValues` SDK) | — | Table caption position: `Above` / `Below` |
| `HtmlStyles` | — | Style manager (`WordDocumentStyle`, see §4) |

---

## 4. Styles — `WordDocumentStyle` (property `converter.HtmlStyles`)

- **`DefaultStyles`** — names of the OpenXml styles applied to generated elements (change them to reuse styles already present in a template):

  | Property | Default | Used for |
  |---|---|---|
  | `Paragraph` | `Normal` | Body paragraphs |
  | `HeadingStyle` | `Heading` | Headings (the level is **appended**: Heading1…Heading6) |
  | `NumberedHeadingStyle` | `Heading` | Alternate style for numbered headings |
  | `HyperlinkStyle` | `Hyperlink` | Links |
  | `CaptionStyle` | `Caption` | `<caption>` |
  | `QuoteStyle` / `IntenseQuoteStyle` | `Quote` / `IntenseQuote` | `<blockquote>` / intense variant |
  | `ListParagraphStyle` | `ListParagraph` | List items |
  | `TableStyle` / `PreTableStyle` | `TableGrid` / `TableGrid` | Tables / `<pre>` table |
  | `FootnoteTextStyle`, `FootnoteReferenceStyle` | `FootnoteText` / `FootnoteReference` | Footnotes |
  | `EndnoteTextStyle`, `EndnoteReferenceStyle` | `EndnoteText` / `EndnoteReference` | Document notes |
  | `HeaderStyle` / `FooterStyle` | `Header` / `Footer` | Header/footer paragraphs |

- **`StyleMissing`** — event (`StyleEventArgs` with `Name` + `Type`) fired when a style name (from `class`, `DefaultStyles`, etc.) does not exist in the document; in the handler you can add it with `HtmlStyles.AddStyle(new Style {...})`.
- **`AddStyle(Style)`** — registers an OpenXml style in the document and refreshes the cache.
- **`QuoteCharacters`** — quote pairs for `<q>`: `QuoteChars.IE` (« »), `Gecko` (“ ”), `WebKit` (" "). Default `IE`.

**Predefined styles auto-inserted** when missing from the document: `Caption`, `Heading1`–`Heading6`, `Hyperlink`, `TableGrid`, `FootnoteReference/FootnoteText`, `EndnoteReference/EndnoteText`, `IntenseQuote`, `ListParagraph`, `Quote`, `QuoteChar`, `Normal`.

Style rules:
- **External CSS and `<style>` are ignored** → inline CSS is required (suggested: pre-inline with [PreMailer.Net](https://www.nuget.org/packages/PreMailer.Net) `PreMailer.MoveCssInline(html, css)`).
- `class` attribute on a tag → looks up a Word style with that name (case-insensitive); if absent, `StyleMissing` fires. Multiple names supported (`class="Standard_Table TableWhite"` tries them in order).
- Inline `style` attributes take precedence over the applied default style.

---

## 5. Remote-image I/O — `HtmlToOpenXml.IO`

| Type | Role |
|---|---|
| `IWebRequest` | `Task<Resource?> FetchAsync(Uri, CancellationToken)` + `bool SupportsProtocol(string)` — abstracts resource (image) download. Custom implementation for credentials/proxy/conversions |
| `DefaultWebRequest` | Default. Protocols: **http, https, file**. `DefaultWebRequest(ILogger?)` or `DefaultWebRequest(HttpClient, ILogger?)` (custom HttpClient for credentials). Automatic gzip/deflate decompression. `virtual` methods `DownloadLocalFile`/`DownloadHttpFile` to override |
| `BaseImageUrl` (on DefaultWebRequest) | Absolute base to resolve relative URLs (`new Uri("http://site/path/")` or a local path). Ensures trailing `/` for `file://` |
| `Resource` | `StatusCode`, `Headers` (incl. Content-Type), `Content` (Stream), `IDisposable` |

Example: WebP is unsupported → override `DownloadHttpFile`, convert to PNG with SixLabors.ImageSharp, replace `resource.Content`.

---

## 6. Supported HTML tags (complete list)

Inline text: `a`, `abbr`, `acronym`, `b`, `cite`, `del`, `dfn`, `em`, `i`, `ins`, `q`, `s`, `strike`, `strong`, `sub`, `sup`, `time`, `u`
Block/structure: `article`, `aside`, `blockquote`, `body`, `br`, `div`, `dl` (`dt`, `dd`), `figure` (`figcaption`), `font`, `h1`–`h6`, `hr`, `img`, `li`, `ol`, `p`, `pre`, `section`, `span`, `svg`, `ul`, `table` (`caption`, `col`, `colgroup`, `tbody`, `td`, `tfoot`, `th`, `thead`, `tr`)

### Attributes per tag

| Tag | Supported attributes |
|---|---|
| `a` | `href` (only valid absolute or relative URIs: `://www.site.com`, `www.site.com`, `http://…`; **ignored** `javascript:`, empty `#`, empty, `site.com` without domain → rendered as a plain run); `title` → tooltip |
| `img` | `src` (absolute, relative with `BaseImageUrl`, or base64 data URI), `alt`, `width`, `height`, `border`; style: `border-style`, `border-width`, `border-color` |
| `table` | `width` (px, pt, %), `align`, `border`, `cellspacing`, `dir`; style: `margin` (`auto` to center), `padding`. ⚠️ **`bgcolor`/`background-color` on `<table>` is NOT applied** (issue #12) — put the background on the cells (`td`/`tr`) instead |
| `td`/`th` | `colspan`, `rowspan`, `width` (px, pt, %), `bgcolor`, `valign`, `align`; style: `writing-mode: tb-lr\|tb-rl` (vertical text), `padding` |
| `tr` | `height` (px, pt, %), `bgcolor`, `valign`, `align` |
| `caption` | `align`; style: `text-align` |
| `col` | `span` (copies the styles onto subsequent columns); style: `color`, `background-color`, `text-align`, … |
| `ol` | `type` (`1\|a\|A\|i\|I`), `start`, `dir`; style: `list-style-type` |
| `ul` | `dir`; style: `list-style-type` |
| `font` | `size` (numeric 1-7 or named `x-large`…), `face`, `color`, `lang` |
| `p`/`div`/`span`/`pre`/`body` | `align`, `border`, `lang`, `dir`, `margin-top`/`margin-bottom` (on `p`); common styles |
| `h1`–`h6` | `id`/`name` (anchor target), `data-bookmark` (registers an explicit bookmark) |
| `pre` | `border`, `lang`, `class` |
| any | `class` → Word style; `id`/`name` → bookmark (anchor target); `data-bookmark` → custom bookmark; `dir` → direction |

### Ignored tags
`button`, `head`, `input`, `script`, `select`, `style`, `textarea`, `xml`, `meta`, HTML comments. Unknown tags are treated as `div`.

---

## 7. Supported CSS (`style` attribute)

| Property | Supported values |
|---|---|
| `text-align` | left/right/center/justify (not on `font`) |
| `color` | rgb(a), hsl(a), hex, named (`HtmlColorTranslator.FromHtml(...).ToHexString()`) |
| `background-color` (and `background` fallback) | as above |
| `text-decoration` | `underline`, `line-through`, `overline` |
| `font-style` | `italic`, `oblique` |
| `font-weight` | `bold`, `bolder`, `lighter`, `normal`, `100`–`900` |
| `font-size` | CSS units (px, em, pt, %, 16.9px…) |
| `font-family` | font name |
| `font-variant` | `small-caps` |
| `text-indent` | px, em (demo: `50px`, `4.5em`) |
| `line-height` | yes (since 3.1.0) |
| `margin`, `padding` | inline and block (since 3.5.0); `margin: auto` centers tables/images |
| `page-break-before` / `page-break-after` | only `always` (on `p`, `pre`, `div`, `span`, `body`; both on the same tag allowed) |
| `break-before` / `break-after` | `page` (since 3.3.0, same effect) |
| `page-orientation` | **non-standard attribute** (only on `body`): `landscape` |
| `writing-mode` | `tb-lr`, `tb-rl` (vertical text, on `td`) |
| `border`, `border-style`, `border-width`, `border-color` | `dotted`, `dashed`, `solid`, `double`, `inset`, `outset`, `none`; `border-width` only px or `medium`/`thick`/`thin`; `border` shorthand (e.g. `1px dashed rgba(0,0,0,.4)`) |

Table (`tr`, `td`, `thead`, `tbody`, `tfoot`): `bgcolor`, `valign`, `align` and styles `background-color`, `vertical-align`, `text-align`, `width`. ⚠️ The `<table>` element itself ignores `bgcolor`/`background-color` (issue #12) — apply backgrounds on `tr`/`td` only.

---

## 8. Lists (`ol`/`ul`)

- Nesting up to **8 levels** (MS Word limit).
- `list-style-type`: `decimal`, `disc`, `square`, `circle`, `upper-alpha`/`lower-alpha`, `upper-roman`/`lower-roman`, `upper-greek`/`lower-greek`, `dash` (non-standard, handy), **custom symbol** (`list-style-type: '+'` or `'👍'`, since 3.4.0).
- `type` attribute on `ol`: `1|a|A|i|I`.
- `start` on `ol`: sets/resets the first number (`<ol start="0">`, `<ol start="50">`); by default the numbering **continues** between consecutive lists (`ContinueNumbering=false` to always reset).
- **Automatic tiered numbering** on headings (1., 1.1, 2.1…) via the `decimal-tiered` style; disable with `SupportsHeadingNumbering=false`. Number-only headings without text do not count.
- Non-W3C nested lists (lists outside `li`) are handled; tables inside lists align with the list-item indentation.

---

## 9. Tables

- `rowspan` + `colspan` supported (also on the same cell; defensive code against overlapping spans).
- **Vertical text**: `style="writing-mode: tb-lr"` (or `tb-rl`) on `td`.
- Nested tables supported.
- `<col>`/`<colgroup>` with `span` for column styles.
- `thead`/`tbody`/`tfoot` sections are **always reordered** to header→body→footer.
- Width: `auto` (fit content), `100%` (fit page, default), `%` (percentage of page), `px`/`pt` (fixed). Percentage cell widths supported.
- Border: no `border` attribute or `border="0"` → no border in Word; `table` tag styles apply to the whole table.
- `margin: auto` for centered alignment (since 3.2.5).
- ⚠️ **macOS Pages** interprets tables poorly (column widths, row/col spans can look messy).
- Caption: `TableCaptionPosition` = above/below the table.

---

## 10. Images

- **Native formats** (supported by OpenXml): `bmp`, `emf`, `gif`, `ico`, `jp2`, `jpe`, `jpeg`, `pcx`, `png`, `svg`, `tif`, `tiff`, `wmf`. **`webp` NOT supported** → needs a custom `IWebRequest` (PNG conversion, e.g. SixLabors.ImageSharp).
- **Data URIs** (`data:image/png;base64,…`) supported (IETF spec).
- Dimensions: `width`/`height` attributes in **px** or **%** (`width:100%` can extend beyond the page margins); with only one dimension → scaled keeping the **aspect ratio**; without dimensions → auto-detection via `ImageHeader.GetDimensions()` (reads only the file header).
- Embedding (default) vs external links (`ImageProcessingMode.LinkExternal`) vs data-URI only (`EmbedDataUriOnly`).
- Download: **parallel prefetch** at parse start, with **per-URL cache** (duplicate images downloaded once).
- Failed downloads / unsupported protocols (`tcp://`, relative without `BaseImageUrl`) → image is skipped without failing the conversion.
- Image border: `border` attribute and `border-style`/`border-width`/`border-color` styles.
- **`<figure>`/`<figcaption>`**: caption above or below the image.
- **SVG**: both inline `<svg>` and `<img src="*.svg">` are embedded as images; the SVG `<title>`/`<desc>` become the Word image description. ⚠️ macOS Pages does not support svg. ⚠️ **Bare-number `width`/`height` attributes are rejected** by the internal `Unit.Parse` (e.g. `width="46"` → a 0x0 image): always emit an explicit unit (`width="46px"`). OfficeSupportTool normalizes this with `NormalizeSvgSizes` (appends `px` to bare numbers on inline `<svg>` and in svg data-URI `<img src>` values) — generators producing SVG must keep the unit.
- Clickable images inside `<a href>` (link or anchor) supported.
- `alt` → figure alt text.

---

## 11. Links, anchors and bookmarks

- Valid external links: `http(s)://`, `://host`, `www.host` (prefixed with http). `title` attribute → tooltip. `History=true` on the relationship.
- Mixed **image + text** links supported (multiple runs inside the hyperlink; `figcaption` inside `a` does not generate paragraphs).
- **Internal anchors**: `<a href="#id">` → bookmark; target reachable via `id` or `name` attributes on any element (h1, div, …) or via `data-bookmark="name"`.
- Reserved anchor `#_top` (and alias `#top`): always accepted even with `SupportsAnchorLinks=false`; the `_top` bookmark is auto-created in the 1st paragraph if missing (merged into the heading when possible, no empty paragraph).
- Links `javascript:`, `#`, empty or without a domain → **plain run** (no hyperlink).
- `data-bookmark` bookmarks: generates `BookmarkStart`/`BookmarkEnd` without a hyperlink.

---

## 12. Footnotes (`abbr`/`acronym`)

- `<abbr title="…">`/`<acronym>` → footnote with the `title` content (the two tags are equivalent).
- The `title` can be a **link** (http/https, file, ftp/ftps, mailto, file share `\\server\share`), not text+link together.
- `AcronymPosition.DocumentEnd` → document note instead of footnote.

---

## 13. Header / Footer

- `ParseHeader`/`ParseFooter` with `HeaderFooterValues` (`Default`, `First`, `Even` = which pages: first, even, odd).
- Images and hyperlinks in headers/footers need relationships with their own part: handled automatically.
- Default paragraph style: `HeaderStyle`/`FooterStyle`.

---

## 14. Preformatted text (`<pre>`)

- Spaces and line breaks are **preserved** (monospaced font).
- `RenderPreAsTable=true` → rendered as a single-cell bordered table (style `TableGrid`), useful for code blocks.
- Supports `border`, `lang`, `class`.

---

## 15. RTL / LTR

- **`dir`** attribute (`rtl`/`ltr`, case-insensitive) on `body` (document scope, also sets the document RTL settings), `p` (text), `ol`/`ul` (lists), `table` (tables), `li`.
- Alternatively `lang`: if it differs from the `body` `lang`, the parser infers the direction from the culture. `dir` is still recommended.
- Since 3.1.0 full RTL handling for text, lists, tables and document scope.

---

## 16. Page

- `page-break-before` / `page-break-after: always` on `p`, `pre`, `div`, `span`, `body` (both on the same tag allowed); modern equivalent `break-before`/`break-after: page` (3.3.0).
- `page-orientation: landscape` — **non-standard** library attribute, only on `<body>` (orients the whole document).

---

## 17. Parsing behavior and limitations

- Whitespace handled across spans/runs (historical fixes #179, #185, #195, #224); `Span<char>`-based parser (3.3.0, +25% performance) with timeouts on remaining regexes (anti-DoS).
- Output is a **self-contained offline document** when images are embedded (Embed).
- If an image download fails, the conversion continues (image skipped).
- `MainDocumentPart.Document`/`Body` auto-created if absent.
- Nested lists beyond 8 levels: not guaranteed (Word limit).
- ⚠️ macOS Pages: tables and SVGs render poorly.
- External CSS/`<style>` not supported → inline CSS only.

## 18. Relevant changelog (v3.x)

| Version | Highlights |
|---|---|
| 3.5.0 | `margin`/`padding` inline and block; AngleSharp 1.5 (CVE fix) |
| 3.4.0 | `list-style-type: dash`; custom symbols `'👍'` |
| 3.3.x | Greek numbering; `Span<char>` parser; .NET 10 target; `break-before/after`; `data-bookmark`; page-break and nested-span fixes |
| 3.2.x | external image links (`LinkExternal`); percentage `col`; EMF; `ol type`; `table width:auto`; `margin auto` tables; **Header/Footer API**; auto `_top`; SVG |
| 3.1.0 | full RTL; `line-height`; `background` fallback; cell-border fixes |
| 3.0.0 | rewrite on AngleSharp + Interpreter/Composite pattern; parallel image download; lists and tables rewrite; 190+ tests |
| 2.4.2+ | OpenXML SDK 3.1; SVG and JPEG2000 embedding; nullable |

---

## 19. Extra from the wiki (OpenXML SDK recipes, not HtmlToOpenXml)

- **dotx → docx**: `template.ChangeDocumentType(WordprocessingDocumentType.Document)` on a writable copy + `mainPart.DocumentSettingsPart.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate", new Uri(templatePath, UriKind.Absolute))`.
- **Prevent document edition**: `ExtendedFilePropertiesPart.Properties.DocumentSecurity.Text = "8"` + `DocumentProtection { Edit = ReadOnly, Enforcement = true }` in the `Settings` (simple lock, not real security; for passwords see the Track-Changes-with-password pattern).
- **Custom properties**: `CustomFilePropertiesPart`; `FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"`, unique `PropertyId` ≥ 2; `DOCPROPERTY` fields in header/body/footer must be updated via XML (first `w:t` of the field).
- **Server delivery**: Content-Type `application/vnd.openxmlformats-officedocument.wordprocessingml.document` + `Content-Disposition: attachment; filename=…` (sanitize `;` and accents; no no-cache headers on IE+SSL).

---

## 20. Pros and cons (summary)

**Pros**: MIT, active, 5.7M downloads, .NET 10, no native binaries/browser, shared `DocumentFormat.OpenXml` dependency, complete image/SVG/table/list/RTL/header-footer handling, parallel image fetching with cache, header/footer API.
**Cons**: inline CSS only (no `<style>`/stylesheet), no native `webp`, heading numbering requires numbered headings, partial macOS Pages support, `page-orientation` only at document level.
