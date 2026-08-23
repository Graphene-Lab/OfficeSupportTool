using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Xml.Linq;
using System.IO.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;
using HtmlToOpenXml;
using SixLabors.ImageSharp;
using UISupportGeneric;

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("OfficeSupportTool.Harness")]

namespace AIOrchestrator.API
{
    /// <summary>Office document (DOCX) operations for agent use: create documents, update existing ones, and create or modify the per-type templates behind them. Types without a template get a new one generated automatically. Every save creates a new version in the workspace git repo (rollback via GitTool.restore). File paths are Unix-style, relative to the workspace root — never escape it.</summary>
    public class OfficeSupportTool : BaseAgentTool, IFileTool
    {
        private static readonly JsonSerializerOptions JsonOpts = new() { PropertyNameCaseInsensitive = true };
        private const int MaxHtmlAttempts = 4;
        private const string OnlyOutputAnswer = "- Output only full HTML code. No opening or closing comments, no fences. [Output only]";
        private const string DraftHint = "If the document is draft, use the draft parameter to generate an incomplete document (it will not be rejected if there is no data to complete it).";
        private const string PlaceholderFormat = "{{ placeholder_name }}";
        private const string HtmlDataRoot = "htmlData";

        /// <summary>Create an office document (DOCX): the template matching the requested type is filled with the provided material and the document is saved. Documents created this way can be updated later with update_document. The new content becomes a new version (rollback via GitTool.restore).</summary>
        /// <param name="type">Type of the document to create, e.g. "balance sheet". Matching is case-insensitive and ignores "-"/space differences. Available template types: [[available_templates]]. Types not in this list are accepted too: a new template is generated and saved for reuse.</param>
        /// <param name="note">Descriptive notes about the document you need to create, which may contain practical suggestions and tips for creating the document.</param>
        /// <param name="draft">Optional: when true, the material check is skipped and the document is generated even if the context lacks the data needed to fill the template (e.g. a draft). Default false: the tool rejects the request when the material cannot fill the template.</param>
        /// <param name="contextText">Optional essential material for generating the document (header, company data, data of the parties involved, tables, lists, values to be entered in the fields, and everything that is required in the general cases for creating this type of document). (Provide contextText or contextFile unless draft is true)</param>
        /// <param name="contextFile">Optional workspace file read as context material (Unix-style path, e.g. "/docs/company-data.md"), same semantics as contextText. (Provide contextText or contextFile unless draft is true)</param>
        /// <param name="imageFiles">Optional workspace image files to place in the document (Unix-style paths, e.g. "/images/logo.png"). Each image is used at most once.</param>
        /// <param name="saveFullNameFile">Optional output file path and name (Unix-style, must end with ".docx", e.g. "/out/invoice.docx"). Default: "/documents/document_yyyyMMdd_HHmmss.docx" in the workspace.</param>
        /// <param name="outputTwoLetterLanguage">Optional two-letter language code for the document content (e.g. "en", "fr"); if omitted, the context data language is used.</param>
        /// <returns>The generated .docx path in workspace form, or an "Error: ..." message (missing input, unsupported image type, insufficient material, unclear description, LLM failure).</returns>
        public string CreateDocument(string type, string note, bool draft = false, string? contextText = null, string? contextFile = null, string[]? imageFiles = null, string? saveFullNameFile = null, string? outputTwoLetterLanguage = null)
        {
            if (string.IsNullOrWhiteSpace(type)) return "Error: type is required.";
            if (string.IsNullOrWhiteSpace(note)) return "Error: note is required.";
            if (saveFullNameFile != null && !saveFullNameFile.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                return "Error: saveFullNameFile must end with '.docx' (the document is saved as a DOCX file).";

            string hostPath;
            try
            {
                hostPath = SandboxPath.Resolve(saveFullNameFile
                    ?? $"/documents/document_{DateTime.Now:yyyyMMdd_HHmmss}.docx");
            }
            catch (UnauthorizedAccessException ex) { return $"Error: {ex.Message}"; }
            Directory.CreateDirectory(Path.GetDirectoryName(hostPath)!);

            var context = new StringBuilder();
            var contextFiles = new List<string>();
            if (!string.IsNullOrWhiteSpace(contextText)) context.AppendLine(contextText);
            if (!string.IsNullOrWhiteSpace(contextFile))
            {
                string ctxHost;
                try { ctxHost = SandboxPath.Resolve(contextFile); }
                catch (UnauthorizedAccessException ex) { return $"Error: {ex.Message}"; }
                if (!File.Exists(ctxHost)) return $"Error: context file '{contextFile}' not found in the workspace.";
                contextFiles.Add(SandboxPath.ToAgent(ctxHost));
                context.AppendLine(ReadTextCapped(ctxHost, 60_000));
            }

            var (images, imagesError) = ResolveImages(imageFiles);
            if (imagesError != null) return imagesError;

            Log.LogStep($"OfficeSupportTool.CreateDocument: type='{type}' draft={draft} images={images!.Count} contextLen={context.Length}");

            var key = NormalizeType(type);
            var template = ResolveTemplate(key);
            if (template == null)
            {
                Log.LogStep($"OfficeSupportTool.CreateDocument: no template for '{key}' — generating one via LLM");
                var generated = GenerateTemplate(key, type);
                if (generated == null) return $"Error: the LLM returned no usable template after {MaxHtmlAttempts} attempts. Retry later.";
                template = generated;
                SaveGeneratedTemplate(key, generated);
            }

            if (!draft)
            {
                var opinion = AskMaterialSufficient(template, context.ToString());
                if (opinion == null) return "Error: the LLM returned no usable evaluation of the material. Retry later.";
                if (!opinion.Sufficient)
                {
                    var fields = TemplateFields(template);
                    return $"Error: the context does not provide all the information needed to fill the document template. {Reasons(opinion.Missing, "the material does not cover the template fields")}." +
                        (fields.Count > 0 ? $" Document fields: {string.Join(", ", fields)}." : "") +
                        $" {DraftHint}";
                }
            }

            var html = GenerateHtml(BuildCreatePrompt(type, note, contextFiles, context.ToString(), images, template, outputTwoLetterLanguage));
            if (html == null) return $"Error: the LLM returned no usable HTML after {MaxHtmlAttempts} attempts. Retry later.";
            html = EnsureImagesUsed(html, images);
            html = EmbedImages(html, images);
            html = EmbedSvgIcons(html);

            string? versionId = null;
            try
            {
                var bytes = ConvertToDocx(html);
                File.WriteAllBytes(hostPath, bytes);
                versionId = GitSupport.Snapshot(hostPath, "OfficeSupportTool create");
            }
            catch (Exception ex)
            {
                Log.LogStep($"OfficeSupportTool.CreateDocument: failed '{hostPath}': {ex}");
                return "Error: cannot create the document (conversion or write failed). Retry later.";
            }
            Log.LogStep($"OfficeSupportTool.CreateDocument: wrote '{hostPath}' ({html.Length} chars html) version='{versionId}'");
            return versionId != null
                ? $"Document created at {SandboxPath.ToAgent(hostPath)}. New version: {versionId}. (Rollback via GitTool.restore.)"
                : $"Document created at {SandboxPath.ToAgent(hostPath)}.";
        }

        /// <summary>Updates an existing DOCX document on request (e.g. change a value, add a clause, fix a table): the requested changes are applied and the file is overwritten in place. The new content becomes a new version (rollback via GitTool.restore).</summary>
        /// <param name="filePath">Path of the document to update (Unix-style, e.g. "/documents/invoice.docx").</param>
        /// <param name="changes">The changes to apply (e.g. "change the invoice number to INV-2026-042, add a 2% late-payment clause, update the total").</param>
        /// <param name="contextText">Optional extra context the update must respect.</param>
        /// <param name="imageFiles">Optional workspace image files the update must place in the document (Unix-style paths, e.g. "/images/logo.png"), same semantics as in CreateDocument: each image is used at most once.</param>
        /// <returns>The updated .docx path in workspace form (with the new version id), or an "Error: ..." message (missing input, document that cannot be updated, unclear changes, LLM failure).</returns>
        public string UpdateDocument(string filePath, string changes, string? contextText = null, string[]? imageFiles = null)
        {
            if (string.IsNullOrWhiteSpace(changes)) return "Error: changes is required.";
            string hostPath;
            try { hostPath = SandboxPath.Resolve(filePath); }
            catch (UnauthorizedAccessException ex) { return $"Error: {ex.Message}"; }
            if (!File.Exists(hostPath)) return $"Error: file '{filePath}' not found in the workspace.";
            if (!hostPath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                return "Error: only .docx documents can be updated (they are generated by CreateDocument).";

            string? currentHtml;
            try { currentHtml = ReadStoredHtml(hostPath); }
            catch (Exception ex)
            {
                Log.LogStep($"OfficeSupportTool.UpdateDocument: cannot read metadata '{hostPath}': {ex}");
                return "Error: the document is corrupted or unreadable and cannot be updated.";
            }
            if (currentHtml == null) return "Error: this document cannot be updated (only documents created by CreateDocument support updates).";

            var (images, imagesError) = ResolveImages(imageFiles);
            if (imagesError != null) return imagesError;

            Log.LogStep($"OfficeSupportTool.UpdateDocument: '{hostPath}' changes='{Truncate(changes, 120)}' contextText='{Truncate(contextText ?? "", 120)}' images={images!.Count}");
            var verdict = AskChangesClear(changes, contextText, images);
            if (verdict == null) return "Error: the LLM returned no usable evaluation of the changes. Retry later.";
            if (!verdict.Clear)
                return $"Error: the requested changes are not clear enough to apply. {Reasons(verdict.Explanation, "the changes do not say what to modify")}";

            var html = GenerateHtml(BuildUpdatePrompt(currentHtml, changes, contextText, images));
            if (html == null) return $"Error: the LLM returned no usable HTML after {MaxHtmlAttempts} attempts. Retry later.";
            html = EnsureImagesUsed(html, images);
            html = EmbedImages(html, images);
            html = EmbedSvgIcons(html);

            string? versionId = null;
            try
            {
                var bytes = ConvertToDocx(html);
                File.WriteAllBytes(hostPath, bytes);
                versionId = GitSupport.Snapshot(hostPath, "OfficeSupportTool update");
            }
            catch (Exception ex)
            {
                Log.LogStep($"OfficeSupportTool.UpdateDocument: failed '{hostPath}': {ex}");
                return "Error: cannot apply the changes (conversion or write failed). Retry later.";
            }

            Log.LogStep($"OfficeSupportTool.UpdateDocument: '{hostPath}' updated ({html.Length} chars html) version='{versionId}'");
            return versionId != null
                ? $"Document updated at {SandboxPath.ToAgent(hostPath)}. New version: {versionId}. (Rollback via GitTool.restore.)"
                : $"Document updated at {SandboxPath.ToAgent(hostPath)}.";
        }

        /// <summary>Modifies an existing document template: the requested changes are checked for feasibility, applied and validated before the updated template is saved for reuse by create_document. The updated template keeps the same standard as created templates.</summary>
        /// <param name="type">Type of the document template to modify, e.g. "balance sheet". Matching is case-insensitive and ignores "-"/space differences. Available template types: [[available_templates]].</param>
        /// <param name="changes">Descriptive description of the changes to apply to the template (e.g. "add a 'Payment Terms' field in the meta strip, right after 'Customer Ref.'").</param>
        /// <returns>A confirmation message with the updated template key, or an "Error: ..." message (unknown template type, infeasible changes, LLM failure).</returns>
        public string UpdateTemplate(string type, string changes)
        {
            if (string.IsNullOrWhiteSpace(type)) return "Error: type is required.";
            if (string.IsNullOrWhiteSpace(changes)) return "Error: changes is required.";
            var key = NormalizeType(type);
            var template = ResolveTemplate(key);
            if (template == null) return $"Error: no template exists for '{type}'. Create one first by calling CreateDocument with this type (the template is generated and saved for reuse).";

            Log.LogStep($"OfficeSupportTool.UpdateTemplate: type='{type}' key='{key}' changes='{Truncate(changes, 120)}'");
            var verdict = AskTemplateChangesFeasible(template, changes);
            if (verdict == null) return "Error: the LLM returned no usable evaluation of the changes. Retry later.";
            if (!verdict.Feasible)
                return $"Error: the requested changes are not feasible for the '{key}' template. {Reasons(verdict.Explanation, "the changes do not fit this template")}";

            var updated = GenerateModifiedTemplate(key, type, template, changes);
            if (updated == null) return $"Error: the LLM returned no usable template after {MaxHtmlAttempts} attempts. Retry later.";
            SaveGeneratedTemplate(key, updated);
            Log.LogStep($"OfficeSupportTool.UpdateTemplate: saved '{key}' ({updated.Length} chars)");
            return $"Template '{key}' updated. Future CreateDocument calls for this type use the updated template.";
        }

        // ---------- Templates ----------

        /// <summary>Canonical template key for a document type: lowercase, "-"/space unified
        /// (e.g. "Balance Sheet", "balance-sheet" and "BALANCE SHEET" → "balance-sheet").</summary>
        internal static string NormalizeType(string type) =>
            Regex.Replace(type.Trim().ToLowerInvariant(), @"[\s-]+", "-");

        /// <summary>Deterministic list of the template's variable fields: every distinct
        /// {{ placeholder_name }} in the template, humanized to Title Case (snake_case → words).
        /// The "{{ placeholder }}" metavariable used by the template header comments is skipped.
        /// This list is appended to the material-gate rejection as "Document fields: ..." so the
        /// agent knows exactly which fields to gather.</summary>
        internal static List<string> TemplateFields(string template) =>
            Regex.Matches(template, @"\{\{\s*([a-z0-9_]+)\s*\}\}", RegexOptions.IgnoreCase)
                .Select(m => m.Groups[1].Value)
                .Where(n => n != "placeholder")
                .Select(n => CultureInfo.InvariantCulture.TextInfo.ToTitleCase(n.Replace('_', ' ')))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

        /// <summary>Template lookup directories in priority order: the plugin's shipped/auto-generated
        /// templates, then a workspace fallback folder (used when the plugin folder is not writable).</summary>
        private static IEnumerable<string> TemplateDirs()
        {
            yield return Path.Combine(AppContext.BaseDirectory, "assets", "templates");
            yield return Path.Combine(Setup.DocumentsPath, "_templates");
        }

        /// <summary>Registers the resolver for the [[available_templates]] dynamic placeholder: the
        /// tool description (param "type") shows the current template set — the shipped ones plus any
        /// template generated at runtime by a user prompt and saved for reuse.</summary>
        static OfficeSupportTool()
        {
            Analyzer.DynamicDescriptionRequested += (_, e) =>
            {
                if (e.ToolType == typeof(OfficeSupportTool) && e.Placeholder == "available_templates")
                    e.Value = string.Join(", ", GetAvailableTemplates());
            };
        }

        /// <summary>Every template key currently resolvable: the shipped ones plus any generated at
        /// runtime and saved next to them (or in the workspace _templates fallback).</summary>
        internal static string[] GetAvailableTemplates()
        {
            var keys = new SortedSet<string>(StringComparer.Ordinal);
            foreach (var dir in TemplateDirs())
                if (Directory.Exists(dir))
                    foreach (var file in Directory.GetFiles(dir, "*.html"))
                        keys.Add(Path.GetFileNameWithoutExtension(file));
            return keys.ToArray();
        }

        /// <summary>Returns the template HTML for a canonical key, or null when no template exists.</summary>
        internal static string? ResolveTemplate(string key)
        {
            foreach (var dir in TemplateDirs())
            {
                var file = Path.Combine(dir, key + ".html");
                if (File.Exists(file)) return File.ReadAllText(file, Encoding.UTF8);
            }
            return null;
        }

        /// <summary>Saves a generated template next to the shipped ones (so future calls reuse it);
        /// falls back to the workspace _templates folder when the plugin folder is not writable.</summary>
        internal static string SaveGeneratedTemplate(string key, string html)
        {
            var primary = Path.Combine(AppContext.BaseDirectory, "assets", "templates");
            try
            {
                Directory.CreateDirectory(primary);
                var file = Path.Combine(primary, key + ".html");
                File.WriteAllText(file, html, Encoding.UTF8);
                return file;
            }
            catch (Exception ex)
            {
                Log.LogStep($"OfficeSupportTool.SaveGeneratedTemplate: cannot write '{primary}': {ex.Message}; falling back to workspace");
                var fallback = Path.Combine(Setup.DocumentsPath, "_templates");
                Directory.CreateDirectory(fallback);
                var file = Path.Combine(fallback, key + ".html");
                File.WriteAllText(file, html, Encoding.UTF8);
                return file;
            }
        }

        /// <summary>Shared LLM-facing rules for template creation and editing: the English-language
        /// requirement, the category-color rule (the template must use the colors of the design set
        /// matching its primary purpose), the icon mechanism, the essential rules and the design
        /// guidelines. Used by both GenerateTemplate and GenerateModifiedTemplate so create and
        /// modify keep the same template rules.</summary>
        private static string TemplateRulesPrompt => $$"""
            The template MUST be written in English: all labels, text and placeholders in English (English document and field names normalize the behavior across sessions).
            The template belongs to one of the 7 document categories defined in the design guidelines (Financial & Accounting, Commercial & Trade, Human Resources, Legal & Corporate, Operations & Procedures, Communication & Marketing, Safety & Compliance). Determine the category of this document type by its PRIMARY purpose and use ONLY that category's colors — its pastel, ink and accent rule from the manifest — for the badge, section bars, table headers, totals and accent rules; never invent other colors for graphic elements.
            In the template you can place SVG icons with dynamic content substitution:
            - Use square SVG icons with a self-explanatory file name that can encode size and color: <icon-name>.<size>.<rrggbb>.svg (these files will be auto-generated based on the name you give them). Usage example: disc.32.aa0000.svg (a disc icon, 32x32 px, hex color #aa0000) → <img src="disc.32.aa0000.svg" alt="disc">
            Follow these rules:
            ```text
            {{EssentialGuidelines}}
            ```
            Follow these design guidelines:
            ```markdown
            {{TemplateGuidelines}}
            ```
            Write the template in English.
            {{OnlyOutputAnswer}}
            """;

        /// <summary>Asks the LLM (no history, no distracting context) to create a brand-new template
        /// for a document type that has none, following the shared template rules. Returns the template
        /// HTML or null when the LLM fails.</summary>
        private static string? GenerateTemplate(string key, string type)
        {
            var prompt = $$"""
                Today's date: {{DateTime.Now:yyyy-MM-dd}}

                Create a template for this document type: {{type}}.
                The template must be professionally valid.
                {{TemplateRulesPrompt}}
                """;
            Log.LogStep($"OfficeSupportTool.GenerateTemplate: type='{type}' key='{key}'");
            return GenerateHtml(prompt);
        }

        /// <summary>Asks the LLM (no history) to modify an existing template with the requested
        /// changes, keeping the shared template rules (English, category colors, placeholder format,
        /// design guidelines). Returns the updated template HTML or null when the LLM fails.</summary>
        private static string? GenerateModifiedTemplate(string key, string type, string currentTemplate, string changes)
        {
            var prompt = $$"""
                Today's date: {{DateTime.Now:yyyy-MM-dd}}

                Modify the existing template for this document type: {{type}}.
                Requested changes:
                ```text
                {{changes}}
                ```
                - Apply the changes LITERALLY: the exact strings in the changes request must appear verbatim in the template.
                - Change ONLY what the changes request; keep the rest of the template identical (structure, comments, category colors, existing placeholders).
                - Keep every {{PlaceholderFormat}} the template already uses; add new placeholders only where the changes need them, in the same {{PlaceholderFormat}} format.
                Current template:
                ```html
                {{currentTemplate}}
                ```
                {{TemplateRulesPrompt}}
                """;
            Log.LogStep($"OfficeSupportTool.GenerateModifiedTemplate: type='{type}' key='{key}'");
            return GenerateHtml(prompt);
        }

        /// <summary>Design guidelines for template creation, loaded from the packed asset
        /// (assets/design-guidelines.md); falls back to a compact built-in copy when missing.</summary>
        private static readonly Lazy<string> TemplateGuidelinesLazy = new(() =>
        {
            try
            {
                var file = Path.Combine(AppContext.BaseDirectory, "assets", "design-guidelines.md");
                if (File.Exists(file)) return File.ReadAllText(file, Encoding.UTF8);
            }
            catch { }
            return BuiltInGuidelines;
        }, LazyThreadSafetyMode.ExecutionAndPublication);

        private static string TemplateGuidelines => TemplateGuidelinesLazy.Value;

        /// <summary>Minimal LLM-facing rules for creating/editing a document's HTML, loaded from the
        /// packed asset (assets/essential-guidelines.md); attached to the create/update prompts as
        /// "Follow these rules:" with a ```text fence. Falls back to a compact built-in copy.</summary>
        private static readonly Lazy<string> EssentialGuidelinesLazy = new(() =>
        {
            try
            {
                var file = Path.Combine(AppContext.BaseDirectory, "assets", "essential-guidelines.md");
                if (File.Exists(file)) return File.ReadAllText(file, Encoding.UTF8);
            }
            catch { }
            return BuiltInEssential;
        }, LazyThreadSafetyMode.ExecutionAndPublication);

        private static string EssentialGuidelines => EssentialGuidelinesLazy.Value;

        /// <summary>Compact fallback copy of the essential rules (the full version lives in
        /// Assets/ESSENTIAL-GUIDELINES.md and is packed with the package).</summary>
        private const string BuiltInEssential = """
            You are writing HTML code that will be converted into a Word (DOCX) document. Follow these rules strictly.
            Allowed tags: a abbr acronym b blockquote body br cite del dfn div dl dt dd em figure figcaption font h1 h2 h3 h4 h5 h6 hr i img li ol p pre q s section span strike strong sub sup table caption col colgroup thead tbody tfoot tr th td time u ul svg
            - Inline CSS only (style="..." attributes): text-align, color, background-color, text-decoration, font-style, font-weight, font-size, font-family, font-variant, text-indent, line-height, margin, padding, border(-style/-width/-color), page-break-before/after (always), break-before/after (page). Forbidden: <style>, external CSS, <script>, flexbox/grid, display, position, float, external image URLs.
            - Put background-color on td or tr ONLY, never on <table> (it is ignored and the background disappears).
            - Every svg needs width/height with an explicit unit (e.g. 46px); bare numbers produce an invisible 0x0 image.
            - Do NOT nest HTML comments (a comment must never contain another <!-- or --> inside it); keep marker comments like <!-- SLA-ROW --> outside any banner comment.
            - Replace every {{ placeholder_name }} with real data; keep all inline styles; money values keep the {{ currency }} placeholder; images must be inline/data-URI.
            - Be faithful to the existing document/template: keep structure and inline styles; change ONLY what is requested; the exact strings of the change request MUST appear verbatim.
            - Output ONLY full HTML code. No opening or closing comments, no fences, no explanations.
            """;

        /// <summary>Compact fallback copy of the design guidelines (the full version lives in
        /// Assets/DESIGN-GUIDELINES.md and is packed with the package).</summary>
        private const string BuiltInGuidelines = """
            Design tokens: font 'Calibri','Carlito',sans-serif 10pt; body #1F2937; headings #111827;
            secondary #4B5563; labels #6B7280; muted #9CA3AF; neutral fill #F8FAFC; footer band #F3F4F6.
            Layout: header table (46x46 SVG badge + company_name 15pt bold + tagline 8pt; document title
            18-22pt small-caps on the right, 2px accent rule below) → meta strip (#F8FAFC, 3-4 label/value
            cells) → parties/intro boxes → sections (section bar + content: paragraphs, tables, lists) →
            totals block → notes → signatures (2-3 columns, hr, "Signature & date") → footer band.
            Page breaks before signatures: <p style="page-break-before:always; font-size:2pt; color:#FFFFFF; margin:0;">&nbsp;</p>.
            Placeholders: {{ placeholder_name }} (lowercase snake_case) for every variable field; keep all
            inline styles; money values append {{ currency }}.
            Icons: header badge = inline SVG circle (pastel fill, ink strokes); extra icons via
            <icon-name>.<size>.<rrggbb>.svg placeholders (auto-generated, e.g. disc.32.aa0000.svg).
            Only inline CSS: text-align, color, background-color, text-decoration, font-style, font-weight,
            font-size, font-family, font-variant, text-indent, line-height, margin, padding, border*,
            page-break-before/after, break-before/after, writing-mode, page-orientation (body only).
            Forbidden: <style>, external CSS, <script>, flexbox/grid, display, position, float, external URLs, webp.
            Use pastel fills only for graphic elements (badge, section bars, table headers, totals); keep
            body text neutral; zebra rows #F8FAFC on data tables; small-caps for titles/labels.
            Start each file with a 3-comment header (title, LLM instructions, design line) and add LLM
            comments marking duplicable rows. The template must be professionally valid and self-contained.
            """;

        // ---------- LLM ----------

        /// <summary>Resolves the optional image files to host paths, validating existence and type.
        /// A path that does not exist is retried as a bare file name (searched across the whole
        /// workspace) — agents often pass just the file name instead of the full Unix-style path. Returns (images, null) on success or
        /// (null, error message) on the first bad entry.</summary>
        private static (List<string>? Images, string? Error) ResolveImages(string[]? imageFiles)
        {
            var images = new List<string>();
            if (imageFiles == null) return (images, null);
            foreach (var img in imageFiles.Where(f => !string.IsNullOrWhiteSpace(f)))
            {
                string imgHost;
                try { imgHost = SandboxPath.Resolve(img); }
                catch (UnauthorizedAccessException ex) { return (null, $"Error: {ex.Message}"); }
                if (!File.Exists(imgHost))
                {
                    // bare or mis-prefixed path: search the workspace by file name
                    imgHost = Directory.GetFiles(Setup.DocumentsPath, Path.GetFileName(img), SearchOption.AllDirectories)
                        .FirstOrDefault() ?? imgHost;
                }
                if (!File.Exists(imgHost))
                    return (null, $"Error: image file '{img}' not found in the workspace. Pass the file's Unix-style workspace path (e.g. '/images/coffee.png').");
                if (MimeFor(imgHost) == null) return (null, $"Error: unsupported image type for '{img}'. Use png, jpg, gif, bmp, svg or webp.");
                images.Add(imgHost);
            }
            return (images, null);
        }

        /// <summary>Asks the LLM (no history) whether the given context material contains all the
        /// information needed to fill the template; returns the JSON verdict or null when the LLM fails.</summary>
        private static MaterialVerdict? AskMaterialSufficient(string template, string context)
        {
            using var llm = new LLMUtility(Setup.ProviderConfig.ProviderName);
            var prompt = $$"""
                Today's date: {{DateTime.Now:yyyy-MM-dd}}

                You validate whether the provided context material contains ALL the information needed to fill this specific document from the template below. Evaluate the template field by field (every placeholder, label, table, list, value, date): the document can be generated only when every field of this template is fillable from the context.
                Rules:
                - "sufficient" = true when every template field has its data somewhere in the context, even when the context formats it differently (address layout, lists, wording are rendering details you decide when generating the document).
                - "sufficient" = false ONLY when a template field genuinely has NO data in the context.

                Template:
                ```html
                {{template}}
                ```

                Context:
                ```text
                {{(string.IsNullOrWhiteSpace(context) ? "(none provided)" : context)}}
                ```

                Does the context material contain all the information needed to fill this specific document from the template?

                Respond with ONLY JSON (no fences, no commentary):
                {"sufficient": true|false, "missing": ["what is missing or unclear", ...]}
                - "sufficient" = false when the context lacks information needed to fill the template fields (company data, data of the parties involved, tables, lists, values to be entered in the fields, dates, and everything that is required for this type of document).
                - "missing" lists the concrete missing items when "sufficient" is false; empty array otherwise.
                """;
            var (response, hResult) = llm.SendQuery(prompt, useHistory: false, role: LLMUtility.SystemRole.DocumentPreparer,
                forceJsonResponse: true);
            if (hResult != null || string.IsNullOrWhiteSpace(response)) return null;
            var verdict = TryParseJson<MaterialVerdict>(response);
            if (verdict == null) Log.LogStep($"OfficeSupportTool.AskMaterialSufficient: unparseable JSON response");
            else Log.LogStep($"OfficeSupportTool.AskMaterialSufficient: sufficient={verdict.Sufficient} missing={verdict.Missing?.Count ?? 0}");
            return verdict;
        }

        /// <summary>Asks the LLM (no history) whether the requested changes (and optional context)
        /// are clear enough to apply. The provided images are made known to the evaluator: a request
        /// that refers to an available image (e.g. "add the logo") is clear — without this the gate
        /// would reject it as ambiguous. Returns the JSON verdict or null when the LLM fails.</summary>
        private static ChangeVerdict? AskChangesClear(string changes, string? contextText, List<string> images)
        {
            using var llm = new LLMUtility(Setup.ProviderConfig.ProviderName);
            var prompt = $$"""
                Today's date: {{DateTime.Now:yyyy-MM-dd}}

                You are about to edit an existing office document. Validate that the requested changes are clear enough to apply.
                Rules:
                - Decide sensible details yourself (e.g. which value is "the total", which clause to extend) — these do NOT make the request unclear.
                - Reject ONLY when the changes are genuinely unusable: empty, meaningless or contradictory requests.
                {{(!string.IsNullOrWhiteSpace(contextText) ? "Additional context: " + contextText : "")}}
                {{(images.Count > 0 ? "Available images (the request may refer to them by file name): " + string.Join(", ", images.Select(Path.GetFileName)) : "")}}

                Requested changes: {{changes}}

                Respond with ONLY JSON (no fences, no commentary):
                {"clear": true|false, "explanation": ["what is missing or unclear", ...]}
                - "clear" = false ONLY when the changes cannot be interpreted at all.
                - "explanation" lists what is unclear when "clear" is false; empty array otherwise.
                """;
            var (response, hResult) = llm.SendQuery(prompt, useHistory: false, role: LLMUtility.SystemRole.DocumentPreparer,
                forceJsonResponse: true);
            if (hResult != null || string.IsNullOrWhiteSpace(response)) return null;
            var verdict = TryParseJson<ChangeVerdict>(response);
            if (verdict == null) Log.LogStep($"OfficeSupportTool.AskChangesClear: unparseable JSON response");
            else Log.LogStep($"OfficeSupportTool.AskChangesClear: clear={verdict.Clear}");
            return verdict;
        }

        /// <summary>Asks the LLM (no history) whether the requested template changes are feasible
        /// and make sense for the given template HTML (the template obeys strict DOCX conversion
        /// rules: inline CSS only, no scripts, no external content). Returns the JSON verdict or
        /// null when the LLM fails.</summary>
        private static TemplateChangeVerdict? AskTemplateChangesFeasible(string template, string changes)
        {
            using var llm = new LLMUtility(Setup.ProviderConfig.ProviderName);
            var prompt = $$"""
                Today's date: {{DateTime.Now:yyyy-MM-dd}}

                You are about to modify an HTML document template that will be converted into a Word (DOCX) document. The template obeys strict rules: only inline CSS, no <style>, no scripts, no external URLs or dynamic content, no flexbox/grid. Validate that the requested changes are feasible and make sense for the given template.
                Rules:
                - "feasible" = true when the changes can be applied to this template: they refer to existing elements/placeholders or add new ones coherently, without violating the template constraints.
                - "feasible" = false ONLY when the changes are genuinely inapplicable: they contradict the template's structure or document type, or require features the DOCX conversion forbids (scripts, external/dynamic content).
                - Decide sensible details yourself; wording, layout and new field additions do NOT make the changes infeasible.

                Template:
                ```html
                {{template}}
                ```

                Requested changes: {{changes}}

                Respond with ONLY JSON (no fences, no commentary):
                {"feasible": true|false, "explanation": ["why the changes are infeasible", ...]}
                - "explanation" lists why the changes are infeasible when "feasible" is false; empty array otherwise.
                """;
            var (response, hResult) = llm.SendQuery(prompt, useHistory: false, role: LLMUtility.SystemRole.DocumentPreparer,
                forceJsonResponse: true);
            if (hResult != null || string.IsNullOrWhiteSpace(response)) return null;
            var verdict = TryParseJson<TemplateChangeVerdict>(response);
            if (verdict == null) Log.LogStep($"OfficeSupportTool.AskTemplateChangesFeasible: unparseable JSON response");
            else Log.LogStep($"OfficeSupportTool.AskTemplateChangesFeasible: feasible={verdict.Feasible}");
            return verdict;
        }

        /// <summary>Generates HTML via the LLM (no history), validates it as HTML5 and retries up to
        /// <see cref="MaxHtmlAttempts"/> times, feeding back the validation errors. Returns null when all attempts fail.</summary>
        private static string? GenerateHtml(string prompt)
        {
            using var llm = new LLMUtility(Setup.ProviderConfig.ProviderName);
            for (int attempt = 1; attempt <= MaxHtmlAttempts; attempt++)
            {
                Log.LogStep($"OfficeSupportTool.GenerateHtml: attempt {attempt}/{MaxHtmlAttempts}");
                var (response, hResult) = llm.SendQuery(prompt, useHistory: false, role: LLMUtility.SystemRole.DocumentPreparer);
                if (hResult != null)
                {
                    Log.LogStep($"OfficeSupportTool.GenerateHtml: LLM error hResult={hResult} on attempt {attempt} — retrying");
                    continue;
                }
                if (string.IsNullOrWhiteSpace(response)) continue;
                var html = response;
                if (!Utility.RemoveFencesEncapsulationAndFixTrim(ref html, false))
                {
                    Log.LogStep($"OfficeSupportTool.GenerateHtml: malformed fences on attempt {attempt}");
                    continue;
                }
                if (IsValidHtml5(html, out var errors))
                {
                    var violations = new List<string>();
                    if (HasNestedComments(html))
                        violations.Add("NESTED HTML comments: a `<!-- ... -->` comment whose body contains another `<!--`. In HTML the inner `-->` closes the outer comment and the remaining text becomes visible in the document; move every marker comment OUTSIDE the banner, directly above the element it marks.");
                    if (HasBareSvgSizes(html))
                        violations.Add("inline <svg> width/height attributes without an explicit unit (e.g. width=\"46\"): every svg width/height MUST carry a unit (width=\"46px\" height=\"46px\"); bare numbers are rejected and produce an invisible image.");
                    if (HasTableBackground(html))
                        violations.Add("background-color/bgcolor on a <table> tag: the converter ignores table-level backgrounds (the fill disappears); move every background-color to the <td> or <tr> elements.");
                    if (violations.Count > 0)
                    {
                        Log.LogStep($"OfficeSupportTool.GenerateHtml: {violations.Count} rule violation(s) on attempt {attempt}");
                        if (attempt == MaxHtmlAttempts) break;
                        prompt = $"""
                            The previous HTML code violates these rules:
                            - {string.Join("\n                            - ", violations)}
                            Fix ALL of them in the corrected version.
                            Here is the code that failed:
                            ```html
                            {html}
                            ```
                            {OnlyOutputAnswer}
                            """;
                        continue;
                    }
                    Log.LogStep($"OfficeSupportTool.GenerateHtml: valid HTML on attempt {attempt}");
                    return html;
                }
                Log.LogStep($"OfficeSupportTool.GenerateHtml: invalid HTML on attempt {attempt} ({errors.Count} errors: {string.Join(" | ", errors.Take(6))})");
                if (attempt == MaxHtmlAttempts) break;
                var errorFeedback = string.Join("\n", errors.Take(8).Select(e => $"  - {e}"));
                var missingEnd = !html.TrimEnd().EndsWith("</html>", StringComparison.OrdinalIgnoreCase);
                var truncated = missingEnd && html.Length > 20000;
                prompt = $"""
                    The previous HTML5 code you provided was not valid.
                    Here is the code that failed:
                    ```html
                    {html}
                    ```
                    Validation errors:
                    {errorFeedback}
                    {(truncated ? "The output was truncated by the output limit: the closing tags are missing at the end of the document. Produce a MORE COMPACT version that fits in the output limit: keep ALL the content but compress the markup (shorter inline styles, fewer wrapper elements). The output MUST end with the closing </html> tag." : missingEnd ? "The document does not end with the closing </html> tag." : "")}

                    Please fix ALL the errors above and provide a corrected, valid HTML5 version.
                    {OnlyOutputAnswer}
                    """;
            }
            return null;
        }

        private static string BuildCreatePrompt(string type, string note, List<string> contextFiles, string context, List<string> images, string template, string? outputTwoLetterLanguage = null)
        {
            // auto detect language if not specified: from the context, else from the first context file
            if (string.IsNullOrWhiteSpace(outputTwoLetterLanguage))
            {
                outputTwoLetterLanguage = Utility.DetectLanguage(context);
                if (outputTwoLetterLanguage == null && contextFiles.Count > 0)
                {
                    try { outputTwoLetterLanguage = Utility.DetectLanguage(ReadTextCapped(SandboxPath.Resolve(contextFiles[0]), 60_000)); }
                    catch (UnauthorizedAccessException) { }
                }
                outputTwoLetterLanguage ??= "en";
            }
            string languageName;
            try { languageName = new CultureInfo(outputTwoLetterLanguage).EnglishName; }
            catch (CultureNotFoundException) { languageName = "English"; }

            var sb = new StringBuilder();
            sb.AppendLine("Today's date: " + DateTime.Now.ToString("yyyy-MM-dd"));
            sb.AppendLine($"Create an office document in {languageName} using the HTML template below.");
            sb.AppendLine("Document type: " + type);
            sb.AppendLine("Document note:");
            sb.AppendLine("```text");
            sb.AppendLine(note);
            sb.AppendLine("```");
            if (contextFiles.Count > 0)
            {
                sb.AppendLine("Context documents (workspace paths):");
                foreach (var p in contextFiles) sb.AppendLine("- " + p);
            }
            if (!string.IsNullOrWhiteSpace(context))
            {
                sb.AppendLine("Context content:");
                sb.AppendLine("```text");
                sb.AppendLine(context.TrimEnd());
                sb.AppendLine("```");
            }
            sb.AppendLine();
            sb.AppendLine("Follow these rules:");
            sb.AppendLine("```text");
            sb.AppendLine(EssentialGuidelines.TrimEnd());
            sb.AppendLine("```");
            sb.AppendLine("- Follow the template's HTML comments (e.g. duplicate the marked rows once per line item).");
            sb.AppendLine("- Only inline CSS is supported: no <style>, no external CSS, no scripts, no flexbox/grid.");
            if (images.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Available images (reference by file name only, e.g. <img src=\"logo.png\">):");
                sb.AppendLine(FileManager.GetFilesInfo(images.Select(SandboxPath.ToAgent)));
                sb.AppendLine("- Use each image once");
                sb.AppendLine();
            }
            sb.AppendLine("- You may add SVG icons with a self-explanatory file name that can encode size and color: <icon-name>.<size>.<rrggbb>.svg (these files will be auto-generated based on the name you give them). Usage example: disc.32.aa0000.svg (a disc icon, 32x32 px, hex color #aa0000) → <img src=\"disc.32.aa0000.svg\" alt=\"disc\">");
            sb.AppendLine("HTML template:");
            sb.AppendLine("```html");
            sb.AppendLine(template);
            sb.AppendLine("```");
            sb.AppendLine("- Write the content in the language of the note.");
            sb.AppendLine("- The output MUST be in HTML format");
            sb.AppendLine("- Check before output");
            sb.AppendLine(OnlyOutputAnswer);
            return sb.ToString();
        }

        private static string BuildUpdatePrompt(string currentHtml, string changes, string? contextText, List<string> images)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Today's date: " + DateTime.Now.ToString("yyyy-MM-dd"));
            sb.AppendLine("Edit an existing office document (DOCX). The document is stored as HTML that reproduces an HTML template.");
            sb.AppendLine("Apply the requested changes LITERALLY to the HTML below:");
            sb.AppendLine("- The exact strings in the changes request (titles, labels, text, values) MUST appear verbatim in the output — do not reword or replace them.");
            sb.AppendLine("- Change ONLY what the changes request; keep the rest of the content, wording and structure identical.");
            sb.AppendLine("Follow these rules:");
            sb.AppendLine("```text");
            sb.AppendLine(EssentialGuidelines.TrimEnd());
            sb.AppendLine("```");
            sb.AppendLine();
            sb.AppendLine("Current document HTML:");
            sb.AppendLine("```html");
            sb.AppendLine(currentHtml);
            sb.AppendLine("```");
            sb.AppendLine();
            sb.AppendLine("Requested changes: " + changes);
            if (images.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Available images (reference by file name only, e.g. <img src=\"logo.png\">):");
                sb.AppendLine(FileManager.GetFilesInfo(images.Select(SandboxPath.ToAgent)));
                sb.AppendLine("- Use each image once");
                sb.AppendLine();
            }
            if (!string.IsNullOrWhiteSpace(contextText)) sb.AppendLine("Additional context: " + contextText);
            sb.AppendLine("You may add SVG icons with a minimalist self-descriptive file name, such as <icon-name>.svg (these files will be auto-generated based on the minimalist name you give them).");
            sb.AppendLine("Write the content in the language of the document.");
            sb.AppendLine(OnlyOutputAnswer);
            return sb.ToString();
        }

        private static string Reasons(List<string>? explanation, string fallback) =>
            explanation is { Count: > 0 }
                ? string.Join(" ", explanation.Select(e => "- " + e))
                : "- " + fallback;

        // ---------- HTML post-processing ----------

        /// <summary>Guarantees every provided image is used: any image whose file name does not
        /// appear in the generated HTML is forced in through the update flow (the change request
        /// instructs to insert it where most appropriate). The LLM often drops images — without this
        /// the omission would be silent.</summary>
        private static string EnsureImagesUsed(string html, List<string> images)
        {
            if (images.Count == 0) return html;
            var missing = images.Where(i => !html.Contains(Path.GetFileName(i), StringComparison.OrdinalIgnoreCase)).ToList();
            if (missing.Count == 0)
            {
                Log.LogStep("OfficeSupportTool.EnsureImagesUsed: all provided images referenced");
                return html;
            }
            Log.LogStep($"OfficeSupportTool.EnsureImagesUsed: forcing {missing.Count} missing image(s): " +
                        string.Join(", ", missing.Select(Path.GetFileName)));
            const string changes = "These images are part of the document and should be inserted where most appropriate.";
            return GenerateHtml(BuildUpdatePrompt(html, changes, null, images)) ?? html;
        }

        /// <summary>Replaces every src reference to a provided image with an inline data URI, so the
        /// document is self-contained. The reference may be a bare file name or a path ending with it
        /// (src="logo.png", src="./img/logo.png", src="/img/logo.png").</summary>
        private static string EmbedImages(string html, List<string> images)
        {
            foreach (var img in images)
            {
                var name = Path.GetFileName(img);
                var dataUri = "data:" + MimeFor(img) + ";base64," + Convert.ToBase64String(File.ReadAllBytes(img));
                html = Regex.Replace(html,
                    $@"src=[""'](?:[^""'/]*/)*{Regex.Escape(name)}[""']",
                    m => $"src=\"{dataUri}\"", RegexOptions.IgnoreCase);
            }
            return html;
        }

        /// <summary>Replaces every "&lt;icon-name&gt;[.&lt;size&gt;].[.&lt;rrggbb&gt;].svg" img placeholder with the
        /// matching icon from the host assets, encoded as a data URI (shared logic in
        /// Utility.EmbedSvgIcons — same pipeline as the presentation/document tools).</summary>
        internal static string EmbedSvgIcons(string html)
        {
            var iconsPath = Path.Combine(AppContext.BaseDirectory, "assets", "icons");
            return Utility.EmbedSvgIcons(html, iconsPath);
        }

        private static string? MimeFor(string path) =>
            Path.GetExtension(path).ToLowerInvariant() switch
            {
                ".png" => "image/png",
                ".jpg" or ".jpeg" => "image/jpeg",
                ".gif" => "image/gif",
                ".bmp" => "image/bmp",
                ".svg" => "image/svg+xml",
                ".webp" => "image/webp",
                _ => null
            };

        private static readonly HashSet<HtmlParseErrorCode> CriticalErrors = new()
        {
            HtmlParseErrorCode.TagNotClosed, HtmlParseErrorCode.TagNotOpened,
            HtmlParseErrorCode.EndTagNotRequired, HtmlParseErrorCode.EndTagInvalidHere
        };

        private static bool IsValidHtml5(string html, out List<string> errors)
        {
            errors = new List<string>();
            if (string.IsNullOrWhiteSpace(html))
            {
                errors.Add("HTML is empty.");
                return false;
            }
            var doc = new HtmlDocument();
            doc.LoadHtml(html);
            foreach (var e in doc.ParseErrors)
            {
                if (CriticalErrors.Contains(e.Code))
                    errors.Add($"Line {e.Line}, Pos {e.LinePosition}: {e.Reason}");
            }
            return errors.Count == 0;
        }

        /// <summary>Detects invalid nested HTML comments: a comment whose body contains another
        /// "&lt;!--" (e.g. a banner comment embedding a row marker). In HTML the inner "-->" closes
        /// the outer comment, leaking the remaining text as visible content in the DOCX.</summary>
        internal static bool HasNestedComments(string html) =>
            Regex.IsMatch(html, @"<!--(?:(?!-->)[\s\S])*?<!--");

        /// <summary>Detects inline &lt;svg&gt; width/height attributes without an explicit unit
        /// (e.g. width="46"): HtmlToOpenXml rejects bare numbers and renders a 0x0 image.</summary>
        internal static bool HasBareSvgSizes(string html) =>
            Regex.IsMatch(html, @"<svg[^>]*\s(?:width|height)\s*=\s*""\d+(?!px)""");

        /// <summary>Detects a background set on a &lt;table&gt; (style="background-color:..." or
        /// bgcolor attribute): HtmlToOpenXml ignores table-level backgrounds, the fill disappears.
        /// Backgrounds must live on &lt;td&gt;/&lt;tr&gt;.</summary>
        internal static bool HasTableBackground(string html) =>
            Regex.IsMatch(html, @"<table[^>]*(?:background(?:-color)?\s*[:=]|bgcolor\s*=)", RegexOptions.IgnoreCase);

        // ---------- DOCX conversion + metadata ----------

        /// <summary>Converts the document HTML to a DOCX (fresh in-memory document, inline CSS/SVG
        /// supported by HtmlToOpenXml) and stores the source HTML as hidden metadata (custom XML
        /// part, base64) so UpdateDocument can rebuild the document from it. The custom XML part is
        /// injected via System.IO.Packaging (the OpenXML SDK typed API does not expose CustomXmlPart).</summary>
        internal static byte[] ConvertToDocx(string html)
        {
            // HtmlToOpenXml's Unit.Parse requires an explicit unit on svg width/height attributes
            // (bare numbers like width="46" are rejected and produce a 0x0 image): normalize the
            // render copy, while the metadata keeps the original document HTML.
            var renderHtml = NormalizeSvgSizes(html);
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, autoSave: true))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                var converter = new HtmlConverter(mainPart);
                converter.ParseBody(renderHtml).GetAwaiter().GetResult();
                SetPageMargins(mainPart.Document.Body!);
                mainPart.Document.Save();
            }
            ms.Position = 0;
            using (var pkg = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite))
                WriteHtmlMetadata(pkg, html);
            return ms.ToArray();
        }

        /// <summary>Sets 1 cm page margins (567 twips per side) — the converted documents adapt
        /// better to further conversions (PDF, print); HtmlToOpenXml leaves the default setup.</summary>
        private static void SetPageMargins(Body body)
        {
            var sectPr = body.Elements<SectionProperties>().LastOrDefault()
                ?? body.AppendChild(new SectionProperties());
            sectPr.RemoveAllChildren<PageMargin>();
            sectPr.AppendChild(new PageMargin { Top = 567, Bottom = 567, Left = 567, Right = 567 });
        }

        /// <summary>HtmlToOpenXml's Unit.Parse requires a unit on the svg width/height attributes:
        /// bare numbers (width="46") are rejected and produce a 0x0 image. Appends "px" to bare
        /// numbers on every inline &lt;svg&gt; and inside svg data-URI &lt;img&gt; srcs (icon placeholders).</summary>
        internal static string NormalizeSvgSizes(string html)
        {
            // inline <svg> opening tags (width/height may appear in any order)
            html = Regex.Replace(html, @"<svg(?=[\s>])[^>]*>",
                m => Regex.Replace(m.Value, @"(width|height)\s*=\s*""(\d+(?:\.\d+)?)""",
                    mm => $"{mm.Groups[1].Value}=\"{mm.Groups[2].Value}px\""));
            // svg data-URI <img> srcs (base64-encoded content)
            html = Regex.Replace(html, @"src=""data:image/svg\+xml;base64,([^""]+)""",
                m =>
                {
                    var svg = Encoding.UTF8.GetString(Convert.FromBase64String(m.Groups[1].Value));
                    svg = Regex.Replace(svg, @"(?<![\w-])(width|height)\s*=\s*""(\d+(?:\.\d+)?)""",
                        mm => $"{mm.Groups[1].Value}=\"{mm.Groups[2].Value}px\"");
                    return $"src=\"data:image/svg+xml;base64,{Convert.ToBase64String(Encoding.UTF8.GetBytes(svg))}\"";
                });
            return html;
        }

        /// <summary>Stores the document HTML in the package custom XML part /customXml/htmlData.xml
        /// (root <see cref="HtmlDataRoot"/>, base64 payload — robust against any character, including
        /// XML-invalid control chars). A fresh document never has the part yet.</summary>
        private static void WriteHtmlMetadata(Package pkg, string html)
        {
            var uri = PackUriHelper.CreatePartUri(new Uri("/customXml/htmlData.xml", UriKind.Relative));
            var part = pkg.PartExists(uri) ? pkg.GetPart(uri) : pkg.CreatePart(uri, "application/xml");
            using (var writer = new StreamWriter(part.GetStream(FileMode.Create), Encoding.UTF8))
                writer.Write($"<{HtmlDataRoot} encoding=\"base64\">{Convert.ToBase64String(Encoding.UTF8.GetBytes(html))}</{HtmlDataRoot}>");
            const string relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
            if (!pkg.GetRelationshipsByType(relType).Any(r => r.TargetUri == uri))
                pkg.CreateRelationship(uri, TargetMode.Internal, relType);
        }

        /// <summary>Reads the document HTML back from the package custom XML metadata part; null when
        /// the document carries no such part (it was not created by CreateDocument).</summary>
        internal static string? ReadStoredHtml(string hostPath)
        {
            using var pkg = Package.Open(hostPath, FileMode.Open, FileAccess.Read);
            var uri = PackUriHelper.CreatePartUri(new Uri("/customXml/htmlData.xml", UriKind.Relative));
            if (!pkg.PartExists(uri)) return null;
            var part = pkg.GetPart(uri);
            string xml;
            using (var reader = new StreamReader(part.GetStream(), Encoding.UTF8))
                xml = reader.ReadToEnd();
            try
            {
                var root = XElement.Parse(xml);
                return root.Name.LocalName == HtmlDataRoot
                    ? Encoding.UTF8.GetString(Convert.FromBase64String(root.Value))
                    : null;
            }
            catch (Exception) { return null; }
        }

        // ---------- JSON ----------

        private static T? TryParseJson<T>(string raw) where T : class
        {
            var start = raw.IndexOf('{');
            var end = raw.LastIndexOf('}');
            if (start < 0 || end <= start) return null;
            try { return JsonSerializer.Deserialize<T>(raw.Substring(start, end - start + 1), JsonOpts); }
            catch (JsonException) { return null; }
        }

        private sealed class MaterialVerdict
        {
            public bool Sufficient { get; set; }
            public List<string>? Missing { get; set; }
        }

        private sealed class ChangeVerdict
        {
            public bool Clear { get; set; }
            public List<string>? Explanation { get; set; }
        }

        private sealed class TemplateChangeVerdict
        {
            public bool Feasible { get; set; }
            public List<string>? Explanation { get; set; }
        }

        // ---------- File helpers ----------

        private static string ReadTextCapped(string path, int maxChars)
        {
            var text = File.ReadAllText(path);
            return text.Length <= maxChars ? text : text[..maxChars] + "\n…[truncated]";
        }

        private static string Truncate(string value, int max) =>
            value.Length <= max ? value : value[..max] + "…";
    }
}
