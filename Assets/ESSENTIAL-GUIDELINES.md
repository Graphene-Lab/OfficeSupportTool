# Essential Rules — Creating/Editing Document HTML

You are writing HTML code that will be converted into a Word (DOCX) document. Follow these rules strictly.

## Allowed tags

Use ONLY these tags: `a abbr acronym b blockquote body br cite del dfn div dl dt dd em figure figcaption font h1 h2 h3 h4 h5 h6 hr i img li ol p pre q s section span strike strong sub sup table caption col colgroup thead tbody tfoot tr th td time u ul svg`

## Inline CSS only

- Style ONLY with inline `style="…"` attributes.
- Forbidden: `<style>` tags, external CSS, `<script>`, flexbox/grid, `display`, `position`, `float`, images from external URLs, and anything not listed in the allowed tags/properties above.
- Allowed CSS properties: `text-align`, `color`, `background-color`, `text-decoration`, `font-style`, `font-weight`, `font-size`, `font-family`, `font-variant`, `text-indent`, `line-height`, `margin`, `padding`, `border`, `border-style`, `border-width`, `border-color`, `page-break-before`/`page-break-after` (`always`), `break-before`/`break-after` (`page`).

## Backgrounds

- Put `background-color` on `td` or `tr` ONLY. Never on `<table>` — the converter ignores it and the background silently disappears (text ends up on white).

## SVG

- Give every `svg` a `width` and `height` with an explicit unit, e.g. `width="46px" height="46px"` — bare numbers (`width="46"`) are rejected and the image becomes invisible (0×0).

## Comments

- Do NOT nest HTML comments: a comment must never contain another `<!--` or `-->` inside it (the inner `-->` closes the outer comment and the remaining text becomes visible in the document). Keep marker comments such as `<!-- SLA-ROW -->` outside any banner comment, directly above the element they mark.

## Placeholders and content

- Replace every `{{ placeholder_name }}` with real data from the material; keep all inline styles intact.
- Money values always keep the `{{ currency }}` placeholder.
- Images must be inline/data-URI — no external URLs.

## Editing behaviour (when modifying a document)

- Be faithful to the existing document: keep the template structure, CSS classes and inline styles unchanged; change ONLY what is requested.
- The exact strings in the change request (titles, labels, text, values) MUST appear verbatim in the output — do not reword or replace them.

## Output

- Output ONLY full HTML code. No opening or closing comments, no fences, no explanations.
