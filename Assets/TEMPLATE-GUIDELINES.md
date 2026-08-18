# OfficeSupportTool — Template Design Guidelines (LLM)

Design directives for creating HTML templates for **HtmlToOpenXml** (HTML → DOCX).
All templates live in this `Assets` folder and must be visually uniform, professionally styled,
and "LLM-first": easy for an LLM to edit and repurpose.

The converter supports only the HTML/CSS subset listed in §6. Inline CSS only: no `<style>`,
no external CSS, no scripts, no flexbox/grid, no external image URLs.

---

## 1. Set-theory classification (7 document types)

Documents are partitioned into **7 disjoint sets** by their *primary purpose*.
Each set owns **one hue family** and all its documents use **pastel gradations of that hue**.
This gives a coherent, instantly recognizable visual system while keeping every document distinct.

| Set | Type | Hue family | Base pastel | Ink (dark text) | Accent rule | Count |
|---|---|---|---|---|---|---|
| A | Financial & Accounting | Blue | `#DBEAFE` | `#1E40AF` | `#60A5FA` | 10 |
| B | Commercial & Trade | Green (emerald) | `#D1FAE5` | `#047857` | `#34D399` | 9 |
| C | Human Resources | Violet (purple) | `#EDE9FE` | `#5B21B6` | `#A78BFA` | 6 |
| D | Legal & Corporate | Slate (grey-blue) | `#E2E8F0` | `#334155` | `#94A3B8` | 6 |
| E | Operations & Procedures | Teal (cyan) | `#CCFBF1` | `#115E59` | `#5EEAD4` | 7 |
| F | Communication & Marketing | Amber (orange) | `#FEF3C7` | `#92400E` | `#FBBF24` | 6 |
| G | Safety & Compliance | Red (rose) | `#FEE2E2` | `#991B1B` | `#FCA5A5` | 5 |

**Classification rule:** a document belongs to the set that matches its *primary purpose*, not
any secondary one. Borderline cases are decided as follows:
- **Sales Contract → B** (it is a transaction document), while generic **Contract → D** (it defines legal terms).
- **Business Plan → F** (its purpose is to communicate/persuade investors and banks).
- **Report & Memorandum → E** (internal operational records).
- **Meeting Minutes** appear twice by design: **board minutes → D** (governance), **team minutes → E** (operations).

### Master manifest (file names, kebab-case, lowercase)

| Set | Document (title) | File name | Pastel shade |
|---|---|---|---|
| A | Invoice | `invoice.html` | `#DBEAFE` (blue-100) |
| A | Receipt | `receipt.html` | `#E0F2FE` (sky-100) |
| A | Proforma Invoice | `proforma-invoice.html` | `#BFDBFE` (blue-200) |
| A | Expense Report | `expense-report.html` | `#EFF6FF` (blue-50) |
| A | Budget | `budget.html` | `#BAE6FD` (sky-200) |
| A | Income Statement | `income-statement.html` | `#E0E7FF` (indigo-100) |
| A | Balance Sheet | `balance-sheet.html` | `#C7D2FE` (indigo-200) |
| A | Cash Flow Statement | `cash-flow-statement.html` | `#E0F2FE` (sky-100) |
| A | Explanatory Notes | `explanatory-notes.html` | `#EFF6FF` (blue-50) |
| A | Financial Report | `financial-report.html` | `#BFDBFE` (blue-200) |
| B | Purchase Order | `purchase-order.html` | `#D1FAE5` (emerald-100) |
| B | Customer Order | `customer-order.html` | `#ECFDF5` (emerald-50) |
| B | Delivery Note | `delivery-note.html` | `#BBF7D0` (green-200) |
| B | Quotation | `quotation.html` | `#DCFCE7` (green-100) |
| B | Sales Proposal | `sales-proposal.html` | `#A7F3D0` (emerald-200) |
| B | Sales Contract | `sales-contract.html` | `#F0FDF4` (green-50) |
| B | Packing List / Invoice | `packing-list-invoice.html` | `#D1FAE5` (emerald-100) |
| B | Price List | `price-list.html` | `#DCFCE7` (green-100) |
| B | Product Catalog | `product-catalog.html` | `#A7F3D0` (emerald-200) |
| C | Payroll | `payroll.html` | `#F5F3FF` (violet-50) |
| C | Employment Contract | `employment-contract.html` | `#EDE9FE` (violet-100) |
| C | Offer Letter | `offer-letter.html` | `#F3E8FF` (purple-100) |
| C | Job Description | `job-description.html` | `#DDD6FE` (violet-200) |
| C | Performance Appraisal | `performance-appraisal.html` | `#E9D5FF` (purple-200) |
| C | HR Form | `hr-form.html` | `#EDE9FE` (violet-100) |
| D | Contract | `contract.html` | `#E2E8F0` (slate-200) |
| D | Non-Disclosure Agreement (NDA) | `non-disclosure-agreement-nda.html` | `#F1F5F9` (slate-100) |
| D | Articles of Incorporation | `articles-of-incorporation.html` | `#CBD5E1` (slate-300) |
| D | Bylaws | `bylaws.html` | `#E2E8F0` (slate-200) |
| D | Meeting Minutes (board) | `meeting-minutes-board.html` | `#F1F5F9` (slate-100) |
| D | Business License / Permit | `business-license-permit.html` | `#E2E8F0` (slate-200) |
| E | Work Order | `work-order.html` | `#CCFBF1` (teal-100) |
| E | Meeting Minutes (team) | `meeting-minutes-team.html` | `#ECFEFF` (cyan-50) |
| E | Standard Operating Procedure (SOP) | `standard-operating-procedure-sop.html` | `#99F6E4` (teal-200) |
| E | Manual | `manual.html` | `#F0FDFA` (teal-50) |
| E | Operational Plan | `operational-plan.html` | `#CFFAFE` (cyan-100) |
| E | Report | `report.html` | `#A5F3FC` (cyan-200) |
| E | Memorandum | `memorandum.html` | `#CCFBF1` (teal-100) |
| F | Business Letter | `business-letter.html` | `#FEF3C7` (amber-100) |
| F | Presentation | `presentation.html` | `#FFF7ED` (orange-50) |
| F | Brochure | `brochure.html` | `#FDE68A` (amber-200) |
| F | Newsletter | `newsletter.html` | `#FFEDD5` (orange-100) |
| F | Press Release | `press-release.html` | `#FEF3C7` (amber-100) |
| F | Business Plan | `business-plan.html` | `#FED7AA` (orange-200) |
| G | Risk Assessment Document (DVR) | `risk-assessment-document-dvr.html` | `#FEE2E2` (red-100) |
| G | GDPR / Privacy Document | `gdpr-privacy-document.html` | `#FFE4E6` (rose-100) |
| G | HACCP Manual | `haccp-manual.html` | `#FECDD3` (rose-200) |
| G | Fire Safety Register | `fire-safety-register.html` | `#FFF1F2` (rose-50) |
| G | PPE Delivery Record | `ppe-delivery-record.html` | `#FECACA` (red-200) |

> **Gradation rule:** within a set, cycle the pastel shades in the table above (50 → 100 → 200 and
> the sibling-hue variants). Never use the *ink* or *accent rule* colors as fills — only pastels.

---

## 2. Global design tokens (identical in every document)

| Token | Value | Usage |
|---|---|---|
| Font | `'Calibri','Carlito',sans-serif` | Body, 10pt |
| Body text | `#1F2937` | Default paragraph text |
| Primary / headings | `#111827` | Company name, key values |
| Secondary text | `#4B5563` | Addresses, descriptions |
| Labels / small-caps | `#6B7280` | Meta labels, field captions |
| Muted | `#9CA3AF` | Footer tagline, tiny notes |
| Neutral fill | `#F8FAFC` | Meta strip, party boxes, zebra rows |
| Footer band | `#F3F4F6` | Footer background (light grey) |
| Line-height | `1.4` | Global |
| Accent rule | set's `rule` color | 2px bar under header; 1.5px above footer |

**Per-set tokens** (replace `pastel` / `ink` / `rule` with the set values from §1):
- Badge circle fill = `pastel`; badge strokes = `ink`
- Document title (top-right, small-caps, bold) = `ink`
- Section bars = `pastel` fill, `ink` bold small-caps text
- Table header rows = `pastel` fill, `ink` bold text
- Totals / highlight rows = `pastel` fill, `ink` bold text
- Accent rules = `rule`

---

## 3. Layout blueprint (uniform skeleton)

Every template follows the same block order:

1. **Header** — full-width single-row table, white background, left/right cells:
   - *Left:* icon badge (46×46 SVG) + `{{ company_name }}` (15pt bold `#111827`) + `{{ company_tagline }}` (8pt `#6B7280`)
   - *Right:* document title (18–22pt bold small-caps, `ink`), optional secondary line (8pt `#6B7280`)
2. **Accent rule** — 1-row table, 2px, `rule` color
3. **Meta strip** — 3–4 cells, `#F8FAFC`, each with label (7.5pt small-caps `#6B7280`) + value (10pt bold `#111827`)
4. **Parties / intro** — two half-width boxes (`#F8FAFC`) with small-caps `ink` headings; optional intro paragraph
5. **Sections** — for each section: a **section bar** (`pastel` + `ink` small-caps title) followed by content (paragraphs, tables, lists)
6. **Totals / summary block** (when relevant) — right-aligned table, highlight row in `pastel` with `ink` bold
7. **Notes** — optional `#F8FAFC` box
8. **Signatures** — 2–3 columns, small-caps `ink` titles, `hr` line, "Signature &amp; date" caption (8pt `#6B7280`)
9. **Footer** — accent rule (1.5px) + `#F3F4F6` band, 3 centered lines (company info, muted note)

**Page breaks:** before signature blocks and before confidential clauses use an empty paragraph
`<p style="page-break-before:always; font-size:2pt; color:#FFFFFF; margin:0;">&nbsp;</p>`.

**Tables for layout** must use `width="100%" cellspacing="0" cellpadding="0"` and `vertical-align` on
`<td>`; inner spacing via `style="padding:…"` on the cells.

---

## 4. Icons (SVG)

- **Badge recipe (inline SVG):** the header badge is an inline SVG circle with the set pastel
  fill and the set ink strokes:
  ```html
  <svg width="46" height="46" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
    <title>{{ Document Title }}</title>
    <circle cx="12" cy="12" r="11.5" fill="{{ pastel }}"/>
    <g stroke="{{ ink }}" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round">
      …simple geometric paths that evoke the document…
    </g>
  </svg>
  ```
  Draw simple, minimalist geometric paths yourself (stroke-based, Lucide-style); the `<title>`
  becomes the image description in Word. Inline SVG is embedded by HtmlToOpenXml as an image part.
- **Additional icons (auto-resolved placeholders):** anywhere else in the document you may add
  extra SVG icons with a self-descriptive placeholder file name that encodes size and color:
  `<icon-name>.<size>.<rrggbb>.svg`. These are auto-generated from the host icon set and embedded
  as data URIs. Usage example: `disc.32.aa0000.svg` (a disc icon, 32×32 px, hex color #aa0000) →
  `<img src="disc.32.aa0000.svg" alt="disc">`.
- An LLM may replace the badge with `{{ company_logo }}` via the HTML comment that accompanies it.

---

## 5. Placeholders and LLM directives

- **Placeholder format:** `{{ placeholder_name }}` — lowercase snake_case, used for every variable
  field (dates, names, amounts, clauses, table cells…).
- **LLM comments:** HTML comments (`<!-- … -->`) at the top of the file and next to repetitive
  structures. They must state:
  - what the template is and the converter constraints (inline CSS only, no scripts),
  - which row to duplicate (e.g. line items, risk rows) with a clear marker comment,
  - any color rule the LLM must apply (e.g. risk-level cell colors),
  - where a `page-break` is placed and why,
  - how to move the footer into `converter.ParseFooter()` if repeating on every page is desired.
- **Money values:** always append the `{{ currency }}` placeholder to amounts.
- **Keep inline styles:** the LLM must preserve all `style="…"` attributes when filling data.

---

## 6. Supported HTML/CSS subset (do not exceed)

**Allowed tags:** `a abbr acronym b blockquote body br cite del dfn div dl dt dd em figure figcaption
font h1–h6 hr i img li ol p pre q s section span strike strong sub sup table caption col colgroup
thead tbody tfoot tr th td time u ul svg`

**Allowed CSS (inline `style` only):** `text-align`, `color`, `background-color`, `text-decoration`,
`font-style`, `font-weight`, `font-size`, `font-family`, `font-variant`, `text-indent`, `line-height`,
`margin`, `padding`, `border(-style/-width/-color)`, `page-break-before/after` (`always`),
`break-before/after` (`page`), `writing-mode` (`tb-lr`, `tb-rl`), `page-orientation` (only `body`,
`landscape`).

**Also allowed attributes:** `width`/`height` (px, pt, %), `align`, `valign`, `border`, `bgcolor`,
`cellspacing`, `colspan`, `rowspan`, `dir`, `lang`, `class` (only for a Word style that exists or will
be added via `StyleMissing`), `id`/`name` (bookmarks/anchors), `start`/`type` on `ol`.

**Forbidden:** `<style>`, external CSS, `<script>`, flexbox/grid, `display`, `position`, `float`,
images from external URLs (all assets must be inline/data-URI), `webp`, everything not listed above.

---

## 7. Do's and Don'ts

**Do**
- Use **pastel fills** (`pastel`) only for graphic elements: badge, section bars, table headers,
  totals rows, legend chips. Keep adequate contrast: pastel fill + `ink`/dark text.
- Keep body text in the neutral office scale (§2). Never set body paragraphs in the category color.
- Use zebra striping `#F8FAFC` on data tables (odd rows).
- Use small-caps for titles/labels; use `&mdash;`, `&middot;`, `&nbsp;`, `&ndash;` entities.
- Keep every file fully self-contained (all SVGs inline/data-URI, no external references).
- Start each file with the same 3-comment header (title, LLM instructions, design line).

**Don't**
- No saturated full-width bands (no dark blue/purple/red headers), no dark footer bands.
- No `style` on `font` for `text-align`; no `page-orientation` outside `body`.
- No placeholder format other than `{{ … }}`.
- No unnecessary repetition: reuse tokens from §2, never invent new hex values for neutral text.

---

## 8. Checklist for creating a new template

1. Identify the set (A–G) → pick `pastel` from the manifest, `ink`, `rule`.
2. Build the badge per §4 (simple geometric paths; the icon is auto-resolved at runtime).
3. Assemble the skeleton (§3) with the correct tokens; keep the meta strip 3–4 fields.
4. Write sections with real structure (tables for tabular data, lists where natural); mark duplicable
   rows with LLM comments.
5. Add placeholders everywhere; keep inline styles intact; put money values with `{{ currency }}`.
6. Add page breaks before signatures/confidential parts.
7. The template must be professionally valid: balanced structure, correct table nesting,
   complete `{{ … }}` placeholders for every variable field, closing tags for every element.
