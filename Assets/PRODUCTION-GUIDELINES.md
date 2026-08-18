# OfficeSupportTool — Production Guidelines (environment-specific)

Companion to [`DESIGN-GUIDELINES.md`](DESIGN-GUIDELINES.md): everything in this file refers to
local paths, files and tooling that exist **only in the development/production environment**.
**Never pass this file to the LLM** (it must not see these paths); the LLM-facing part of the
design rules is entirely inside DESIGN-GUIDELINES.md.

---

## Icon source (development machine)

The Lucide-style icon set used for document badges lives on the dev machine at:

```
C:\Users\andre\OneDrive\Sorgenti\AIOrchestrator\assets\icons
```

1743 stroke-based SVG icons (MIT). When *you* build a template by hand, pick the icon that best
matches the document's purpose and paste its paths into the badge recipe (DESIGN-GUIDELINES §4).
The LLM does not need this path: it emits `<icon-name>.<size>.<rrggbb>.svg` placeholders that the
host resolves automatically.

## Converter reference

Full HtmlToOpenXml capabilities and limits: `../HtmlToOpenXml-Guide.md` (in this repository).
`ESSENTIAL-GUIDELINES.md` is the strictly-required HTML/CSS subset (attached to the LLM prompts as
"Follow these rules:"); `DESIGN-GUIDELINES.md` is the complementary design system.

## Verification harness (before delivery)

1. `dotnet build OfficeSupportTool.sln`-equivalent builds; deterministic self-tests:
   `OfficeSupportTool.Harness --selftest` (no LLM, no network).
2. Behavioral LLM tests: `OfficeSupportTool.Harness` (provider via `--provider`, default
   DeepSeekBridge) — creates/updates real documents in `%TEMP%\OfficeSupportTool.Tests-workspace`.
3. Validate the DOCX with `OpenXmlValidator`.
4. Render with LibreOffice headless (`soffice --headless --convert-to png`) and inspect the
   first page visually before delivery.

## Runtime layout

- The templates ship inside the NuGet package at `lib/<tfm>/assets/templates/*.html` and are
  loaded at runtime from `AppContext.BaseDirectory/assets/templates/` (plugin output folder).
- LLM-generated templates are saved into the same folder (fallback: `<workspace>/_templates`
  when the plugin folder is not writable) so they are reused on later calls.
- Icons for the placeholder mechanism are resolved from `AppContext.BaseDirectory/assets/icons`
  (provided by the host, same convention as MD2PDF/PresentationPlugin).
