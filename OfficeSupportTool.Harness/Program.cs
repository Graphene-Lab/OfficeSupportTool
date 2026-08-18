using System.Text;
using System.Text.RegularExpressions;
using System.IO.Packaging;
using AIOrchestrator;
using AIOrchestrator.API;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;

namespace OfficeSupportToolHarness;

/// <summary>
/// OfficeSupportTool test harness: behavioral LLM test (create + update + template generation)
/// plus a deterministic offline self-test (--selftest). Strategy mirrors DocumentTool.Tests /
/// OfficeTool.Tests / PresentationPlugin.Harness (AGENT_TOOLS_GUIDE "Testing Agent Tools").
///   - provider: --provider NAME (default DeepSeekBridge on 127.0.0.1:8787; falls back to
///     the local Ollama qwen3.5:4b runtime provider with --provider Ollama_Qwen)
/// Workspace lives in %TEMP% on purpose: the repo sits under OneDrive and test files
/// written under the repo got cloud-synced on every write (historical slow runs).
/// </summary>
static class Program
{
    private static int _failures;
    private static string _workspace = "";
    private static string _providerName = "DeepSeekBridge";
    private static readonly string ResultsFile = Path.Combine(Path.GetTempPath(), "officesupporttool_test_results.txt");

    static int Main(string[] args)
    {
        var idx = Array.IndexOf(args, "--provider");
        if (idx >= 0 && idx + 1 < args.Length) _providerName = args[idx + 1];
        if (Array.IndexOf(args, "--selftest") >= 0) return RunSelfTest();
        EnsureProvider();

        Log.IsEnabled = true;
        _workspace = Path.Combine(Path.GetTempPath(), "OfficeSupportTool.Tests-workspace");
        try
        {
            if (Directory.Exists(_workspace)) Directory.Delete(_workspace, recursive: true);
        }
        catch (Exception)
        {
        }
        Directory.CreateDirectory(_workspace);
        Setup.SkipIndexingOnStartup = true;
        Setup.DocumentsPath = _workspace;
        Setup.ProviderConfig = ProviderConfigs.Get(_providerName);
        StageIcons();
        StageImages();

        File.WriteAllText(ResultsFile, $"RUN {DateTime.Now:HH:mm:ss} provider={_providerName}\n");
        WriteResult("STARTED");

        Console.WriteLine("══════════ OfficeSupportTool LLM test ══════════");
        Console.WriteLine($"provider: {_providerName}");
        Log.LogStep($"=== OfficeSupportTool LLM test (provider {_providerName}) ===");

        try
        {
            var tool = new OfficeSupportTool();

            // T1 — create from an existing template (balance sheet), material check must pass
            const string context = "Fiori Coffee S.r.l. — via Roma 12, 20121 Milano, Italy. VAT IT01234567890, " +
                "phone +39 02 12345678, email info@fioricoffee.it, website www.fioricoffee.it. " +
                "Balance sheet as of 31 December 2025, prepared in accordance with IFRS, currency EUR. " +
                "Assets: cash and cash equivalents 45,200; accounts receivable 128,500; inventory 96,300; " +
                "prepaid expenses 12,400; property, plant and equipment 342,800; intangible assets 58,100. " +
                "Total assets 683,300. Liabilities: accounts payable 87,600; short-term debt 40,000; " +
                "accrued expenses 18,900; long-term debt 210,000. Equity: share capital 100,000; " +
                "retained earnings 226,800. Total liabilities and equity 683,300. " +
                "Prepared by Maria Rossi, Chief Accountant; approved by Luca Bianchi, CEO.";
            var r1 = tool.CreateDocument(
                "balance sheet",
                "Balance sheet of Fiori Coffee S.r.l. as of 31 December 2025: statement of financial position " +
                "with assets, liabilities and equity in EUR, prepared in accordance with IFRS, plus notes.",
                contextText: context,
                saveFullNameFile: "/balance-sheet.docx");
            Console.WriteLine($"  T1 CreateDocument → {r1}");
            var host1 = Path.Combine(_workspace, "balance-sheet.docx");
            if (!r1.StartsWith("Document created at") || !File.Exists(host1))
            { Fail("T1-create", $"create failed: {r1}"); return 1; }
            var html1 = OfficeSupportTool.ReadStoredHtml(host1);
            if (html1 == null || !html1.Contains("Balance", StringComparison.OrdinalIgnoreCase))
            { Fail("T1-create", "stored HTML metadata missing or empty"); return 1; }
            if (!DocxTextContains(host1, "Fiori Coffee"))
            { Fail("T1-create", "converted DOCX text does not contain the company name"); return 1; }
            Pass("T1-create");

            // T1b — overwriting the same path must report the created backup (Save() pattern)
            var r1b = tool.CreateDocument(
                "balance sheet",
                "Balance sheet of Fiori Coffee S.r.l. as of 31 December 2025 (second version: higher retained earnings).",
                draft: true, contextText: context,
                saveFullNameFile: "/balance-sheet.docx");
            Console.WriteLine($"  T1b CreateDocument(overwrite) → {r1b}");
            if (!r1b.Contains("backed up as", StringComparison.OrdinalIgnoreCase)
                || !File.Exists(Path.Combine(_workspace, "balance-sheet.001.bak")))
            { Fail("T1-overwrite-backup", $"expected backup report: {r1b}"); return 1; }
            Pass("T1-overwrite-backup");

            // T2 — update the document via the embedded HTML metadata
            var r2 = tool.UpdateDocument(
                "/balance-sheet.docx",
                "Change the balance sheet date to 31 March 2026 and increase accounts receivable to 150,000, " +
                "updating total assets and the accounting identity consistently.");
            Console.WriteLine($"  T2 UpdateDocument → {r2}");
            if (!r2.StartsWith("Document updated at") || !r2.Contains(".bak"))
            { Fail("T2-update", $"update failed: {r2}"); return 1; }
            var html2 = OfficeSupportTool.ReadStoredHtml(host1);
            if (html2 == null || !html2.Contains("150,000"))
            { Fail("T2-update", "updated HTML metadata does not contain the new value"); return 1; }
            Pass("T2-update");

            // T3 — unknown type: template generated by the LLM and reused
            var r3 = tool.CreateDocument(
                "service level agreement",
                "Service level agreement between Fiori Coffee S.r.l. and CloudHost GmbH: uptime, response times, penalties.",
                draft: true,
                contextText: "CloudHost GmbH provides hosting for Fiori Coffee. SLA: 99.9% uptime, response within 4 hours, penalty 5% of monthly fee.",
                saveFullNameFile: "/sla.docx");
            Console.WriteLine($"  T3 CreateDocument(unknown type, draft) → {r3}");
            var host3 = Path.Combine(_workspace, "sla.docx");
            if (!r3.StartsWith("Document created at") || !File.Exists(host3))
            { Fail("T3-template-gen", $"create failed: {r3}"); return 1; }
            if (OfficeSupportTool.ResolveTemplate("service-level-agreement") == null)
            { Fail("T3-template-gen", "generated template not persisted for reuse"); return 1; }
            Pass("T3-template-gen");

            // T3b — inspect the LLM-generated template (saved next to the shipped ones in the
            // executable's assets/templates) for conformity with the design rules, and check how
            // the SVG icons it uses survive the DOCX conversion.
            var genTpl = Path.Combine(AppContext.BaseDirectory, "assets", "templates", "service-level-agreement.html");
            if (!File.Exists(genTpl))
            { Fail("T3-template-inspection", $"generated template not found at {genTpl}"); return 1; }
            var tplIssues = InspectTemplate(genTpl, "LLM-generated template");
            var html3 = OfficeSupportTool.ReadStoredHtml(host3);
            if (html3 == null)
            { Fail("T3-template-inspection", "stored HTML metadata missing on generated docx"); return 1; }
            Console.WriteLine($"    converted docx: {Regex.Matches(html3, "src=\"data:image/(?:png|svg\\+xml);base64,").Count} data-URI image(s) in stored HTML, {DocxImageCount(host3)} image part(s) in the DOCX");
            if (tplIssues > 0) { Fail("T3-template-inspection", $"{tplIssues} conformity issue(s)"); return 1; }
            Pass("T3-template-inspection");

            // T4 — images: the document requires a logo + a product photo, provided via imageFiles
            const string t4Context = "Fiori Coffee S.r.l. — via Roma 12, 20121 Milano, Italy. VAT IT01234567890, phone +39 02 12345678, email info@fioricoffee.it, website www.fioricoffee.it. Company tagline: 'Specialty coffee since 1998'. " +
                "Letter date: 18 August 2026. Letter reference: FI-2026-077. " +
                "Recipient: Maria Conti, Procurement Director, Nova S.p.A., via Torino 45, 10122 Torino, Italy. " +
                "Subject: Partnership renewal for 2026. Salutation: 'Dear Ms. Conti'. " +
                "Body: Nova S.p.A. has been a partner since 2019; the collaboration produced 120 new clients in 2025. We thank you for the partnership and look forward to expanding the cooperation in 2026 with a new product line launch in September. " +
                "Closing: 'Best regards'. Sender: Luca Bianchi, CEO. Enclosures: product catalogue 2026. CC: Roberto Verdi, Sales Director, Nova S.p.A.";
            var r4 = tool.CreateDocument(
                "business letter",
                "Business letter from Fiori Coffee S.r.l. to Nova S.p.A. thanking them for the partnership. " +
                "Place the company logo (logo.png) at the top and the product photo (coffee.png) in the body.",
                contextText: t4Context,
                imageFiles: new[] { "/images/logo.png", "/images/coffee.png" },
                saveFullNameFile: "/letter.docx");
            Console.WriteLine($"  T4 CreateDocument(images) → {r4}");
            var host4 = Path.Combine(_workspace, "letter.docx");
            if (!r4.StartsWith("Document created at") || !File.Exists(host4))
            { Fail("T4-images", $"create failed: {r4}"); return 1; }
            var html4 = OfficeSupportTool.ReadStoredHtml(host4);
            var pngUris = html4 == null ? 0 : Regex.Matches(html4, "src=\"data:image/png;base64,").Count;
            if (html4 == null || pngUris != 2)
            { Fail("T4-images", $"expected both images embedded once (2 data URIs), found {pngUris}"); return 1; }
            if (DocxImageCount(host4) < 2)
            { Fail("T4-images", "converted DOCX has fewer than 2 image parts"); return 1; }
            Pass("T4-images");

            // T5 — material gate: draft=false with no context → deterministic rejection + draft hint
            var r5 = tool.CreateDocument("invoice", "Invoice for ACME Corp, March 2026, itemized line items, VAT and total.");
            Console.WriteLine($"  T5 material gate (no context) → {r5}");
            if (!r5.StartsWith("Error:") || !r5.Contains("draft", StringComparison.OrdinalIgnoreCase))
            { Fail("T5-material-gate", $"expected rejection with draft hint: {r5}"); return 1; }
            Pass("T5-material-gate");

            // T6 — update a document without HTML metadata → deterministic error
            var foreign = Path.Combine(_workspace, "foreign.docx");
            File.WriteAllBytes(foreign, OfficeSupportTool.ConvertToDocx("<html><body><p>plain</p></body></html>"));
            using (var pkg = Package.Open(foreign, FileMode.Open, FileAccess.ReadWrite))
            {
                var uri = PackUriHelper.CreatePartUri(new Uri("/customXml/htmlData.xml", UriKind.Relative));
                if (pkg.PartExists(uri))
                {
                    const string relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
                    var rel = pkg.GetRelationshipsByType(relType).FirstOrDefault(r => r.TargetUri == uri);
                    if (rel != null) pkg.DeleteRelationship(rel.Id);
                    pkg.DeletePart(uri);
                }
            }
            var r6 = tool.UpdateDocument("/foreign.docx", "change the greeting");
            Console.WriteLine($"  T6 UpdateDocument(no metadata) → {r6}");
            if (!r6.StartsWith("Error:") || !r6.Contains("no embedded HTML metadata"))
            { Fail("T6-foreign-docx", $"expected metadata error: {r6}"); return 1; }
            Pass("T6-foreign-docx");

            // T7 — full rollback flow at agent level: create → user change → user asks rollback
            Console.WriteLine("  T7 rollback flow (agent-level, may take several minutes)...");
            var (res7, trace7) = RunAgent(
                "TASK 1: Create a business letter from 'Fiori Coffee S.r.l.' to 'Nova S.p.A.' thanking them for the partnership, save it as /rollback.docx. " +
                "Supporting material to pass as contextText: Fiori Coffee S.r.l. — via Roma 12, 20121 Milano, Italy, VAT IT01234567890, email info@fioricoffee.it. " +
                "Letter date 18 August 2026, reference FI-2026-077. Recipient: Maria Conti, Procurement Director, Nova S.p.A., via Torino 45, 10122 Torino, Italy. " +
                "Subject 'Partnership renewal', salutation 'Dear Ms. Conti'. Body: Nova S.p.A. has been a partner since 2019; the collaboration produced 120 new clients in 2025; " +
                "we thank you for the partnership started in 2019 and look forward to 2026. Closing 'Best regards', sender Luca Bianchi (CEO), enclosures: product catalogue 2026. " +
                "TASK 2: The user asks to change the letter: replace the sentence about the partnership with 'thank you for the 2025 campaign results' " +
                "and add the closing 'We look forward to 2026.' " +
                "TASK 3: The user now wants to UNDO the change of TASK 2 (rollback to the version before the change), using the backup the tool reported. " +
                "Report what you did in each task.",
                maxIterations: 50);
            WriteResult("T7 trace: " + string.Join(" | ", trace7));
            Console.WriteLine($"  T7 tool calls: {string.Join(" → ", trace7)}");
            if (res7.Error != null) { Fail("T7-rollback", $"agent error: {res7.Error}"); return 1; }
            var rollbackFile = Path.Combine(_workspace, "rollback.docx");
            var html7 = File.Exists(rollbackFile) ? OfficeSupportTool.ReadStoredHtml(rollbackFile) : null;
            if (html7 == null) { Fail("T7-rollback", "rollback.docx not created or no metadata"); return 1; }
            if (html7.Contains("2025 campaign results", StringComparison.OrdinalIgnoreCase))
            { Fail("T7-rollback", "document still contains the TASK 2 change — restore did not revert"); return 1; }
            if (!html7.Contains("partnership", StringComparison.OrdinalIgnoreCase))
            { Fail("T7-rollback", "original wording missing after restore"); return 1; }
            if (!trace7.Any(c => c.Contains("restore", StringComparison.OrdinalIgnoreCase)))
            { Fail("T7-rollback", $"agent never called Restore — trace: {string.Join(" | ", trace7)}"); return 1; }
            if (Directory.GetFiles(_workspace, "rollback.*.bak").Length < 2)
            { Fail("T7-rollback", "backup chain shorter than expected (update + restore swap)"); return 1; }
            Pass("T7-rollback");

            Console.WriteLine();
            Console.WriteLine(_failures == 0 ? "  ALL LLM TESTS PASSED" : $"  {_failures} LLM TEST FAILURES");
            WriteResult(_failures == 0 ? "DONE PASS" : $"DONE FAIL ({_failures})");
            return _failures == 0 ? 0 : 1;
        }
        catch (Exception ex)
        {
            Fail("main", $"CRASH {ex.GetType().Name}: {ex.Message}");
            WriteResult($"DONE FAIL (crash {ex.GetType().Name})");
            return 1;
        }
    }

    static void Pass(string id) { Console.WriteLine($"  ✓ {id} PASS"); WriteResult($"{id} PASS"); }
    static void Fail(string id, string problem) { _failures++; Console.WriteLine($"  ✗ {id} FAIL: {problem}"); WriteResult($"{id} FAIL: {problem}"); }
    static void WriteResult(string line) => File.AppendAllText(ResultsFile, line + Environment.NewLine);

    /// <summary>Extracts the plain text of a DOCX (paragraphs + table cells) to verify the conversion.</summary>
    static bool DocxTextContains(string path, string expected)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var text = new StringBuilder();
        foreach (var p in doc.MainDocumentPart!.Document!.Descendants<Paragraph>())
            text.AppendLine(p.InnerText);
        foreach (var c in doc.MainDocumentPart.Document.Descendants<TableCell>())
            text.Append(' ').Append(c.InnerText);
        return text.ToString().Contains(expected, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Counts the image parts of a DOCX (rendered images from &lt;img&gt; / SVG conversion).</summary>
    static int DocxImageCount(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        return doc.MainDocumentPart?.ImageParts.Count() ?? 0;
    }

    /// <summary>Runs a behavioral agent-level test: natural-language prompt, "OfficeSupportTool"
    /// registered, tool-call trace collected from the AgentProgress events. Returns the agent
    /// result and the ordered call sequence (what the agent actually did).</summary>
    static (AgentResult Result, List<string> Trace) RunAgent(string prompt, int maxIterations = 50)
    {
        using var orch = new AgentHarness(_providerName);
        var calls = new List<string>();
        orch.AgentProgress += (_, e) =>
        {
            if (e.State == AgentHarness.AgentState.Iteration && !string.IsNullOrWhiteSpace(e.MethodName))
                calls.Add($"#{e.Iteration}:{e.MethodName}");
        };
        var result = orch.ExecuteAction(prompt, new[] { "OfficeSupportTool" }, maxIterations: maxIterations);
        return (result, calls);
    }

    /// <summary>Prints a conformity report for a template (shipped or LLM-generated) against the
    /// essential design rules and returns the number of hard issues found. Also reports which SVG
    /// mechanism the template uses (inline svg / icon-name placeholders / data URIs) — that is what
    /// EmbedSvgIcons + NormalizeSvgSizes must handle when the template becomes a document.</summary>
    static int InspectTemplate(string path, string label)
    {
        var tpl = File.ReadAllText(path, Encoding.UTF8);
        var issues = 0;
        Console.WriteLine($"  ── {label}: {path} ({tpl.Length} chars)");
        void Check(bool ok, string what)
        {
            if (ok) Console.WriteLine($"    ✓ {what}");
            else { issues++; Console.WriteLine($"    ✗ {what}"); }
        }
        Check(!OfficeSupportTool.HasNestedComments(tpl), "no nested HTML comments");
        Check(!OfficeSupportTool.HasTableBackground(tpl), "no background-color/bgcolor on <table> (use tr/td)");
        Check(!OfficeSupportTool.HasBareSvgSizes(tpl), "svg width/height always with explicit unit (px)");
        Check(!Regex.IsMatch(tpl, @"<(?:style|script)\b|(?:src|href)\s*=\s*[""']https?://|url\([""']?https?://", RegexOptions.IgnoreCase),
            "no <style>/<script>/external URLs");
        Check(!Regex.IsMatch(tpl, @"\b(?:display|position|float)\s*:|\bflex\b", RegexOptions.IgnoreCase),
            "no flexbox/grid/position/float CSS");
        Check(!Regex.IsMatch(tpl, @"[A-Za-z]:\\|/home/|/mnt/|AIOrchestrator[\\/]|assets[\\/]icons", RegexOptions.IgnoreCase),
            "no local/host paths in the template");
        var inlineSvg = Regex.Matches(tpl, @"<svg\b").Count;
        var iconPlaceholders = Regex.Matches(tpl, @"<img\b[^>]*src\s*=\s*""([^""]*\.svg)""")
            .Select(m => m.Groups[1].Value).Where(s => !s.StartsWith("data:", StringComparison.OrdinalIgnoreCase)).ToList();
        var svgDataUris = Regex.Matches(tpl, @"data:image/svg\+xml;base64,").Count;
        Console.WriteLine($"    svg mechanism: {inlineSvg} inline <svg>, {iconPlaceholders.Count} icon-name placeholder(s) [{string.Join(", ", iconPlaceholders)}], {svgDataUris} svg data-URI(s)");
        Console.WriteLine(tpl.Contains("{{")
            ? "    ✓ placeholders ({{...}}) present"
            : "    ⚠ no {{ placeholder }} found (the LLM may have filled concrete values)");
        return issues;
    }

    // ---------- deterministic self-test (no LLM, no network) ----------

    /// <summary>Runs the offline deterministic checks: type normalization, template resolution,
    /// HTML→DOCX conversion + metadata round-trip, SVG icon embedding. Exit code 0 = all green.</summary>
    static int RunSelfTest()
    {
        Console.WriteLine("══════════ OfficeSupportTool deterministic self-test ══════════");
        var failures = 0;

        failures += Test("normalize: case/space/dash insensitive", () =>
        {
            if (OfficeSupportTool.NormalizeType("Balance Sheet") != "balance-sheet") return "Balance Sheet";
            if (OfficeSupportTool.NormalizeType("balance-sheet") != "balance-sheet") return "balance-sheet";
            if (OfficeSupportTool.NormalizeType("BALANCE  SHEET") != "balance-sheet") return "BALANCE  SHEET";
            if (OfficeSupportTool.NormalizeType("  Invoice  ") != "invoice") return "  Invoice  ";
            if (OfficeSupportTool.NormalizeType("non-disclosure agreement") != "non-disclosure-agreement") return "non-disclosure agreement";
            return null;
        });

        failures += Test("templates: shipped set resolves", () =>
        {
            var tpl = OfficeSupportTool.ResolveTemplate("balance-sheet");
            if (tpl == null) return "balance-sheet.html not found in harness output";
            if (!tpl.Contains("{{ company_name }}")) return "template placeholders missing";
            return null;
        });

        failures += Test("docx: conversion + metadata round-trip", () =>
        {
            var html = "<!DOCTYPE html><html lang=\"en\"><head><meta charset=\"utf-8\"/></head><body style=\"font-family:'Calibri',sans-serif; font-size:10pt;\">" +
                "<h1>Round-trip &amp; test</h1><p>Hello <strong>World</strong> — € 1.234,56 &lt;tag&gt; 'quotes' \"dq\" \uD83D\uDE80</p>" +
                "<table width=\"100%\" cellspacing=\"0\" cellpadding=\"0\"><tr><td style=\"padding:4px;\">A</td><td style=\"padding:4px;\">B</td></tr></table>" +
                "</body></html>";
            var file = Path.Combine(Path.GetTempPath(), "ost-selftest-" + Guid.NewGuid().ToString("N") + ".docx");
            try
            {
                var bytes = OfficeSupportTool.ConvertToDocx(html);
                File.WriteAllBytes(file, bytes);
                var back = OfficeSupportTool.ReadStoredHtml(file);
                if (back == null) return "metadata not readable";
                if (back != html) return "metadata round-trip mismatch (length " + back.Length + " vs " + html.Length + ")";
                if (!DocxTextContains(file, "Round-trip")) return "converted DOCX text missing";
                if (!DocxTextContains(file, "€")) return "converted DOCX text missing entity character";
                return null;
            }
            finally { try { File.Delete(file); } catch { } }
        });

        failures += Test("docx: no metadata on foreign docx", () =>
        {
            var file = Path.Combine(Path.GetTempPath(), "ost-selftest-" + Guid.NewGuid().ToString("N") + ".docx");
            try
            {
                var bytes = OfficeSupportTool.ConvertToDocx("<html><body><p>x</p></body></html>");
                File.WriteAllBytes(file, bytes);
                // simulate a foreign docx by removing the htmlData custom XML part via OPC
                using (var pkg = Package.Open(file, FileMode.Open, FileAccess.ReadWrite))
                {
                    var uri = PackUriHelper.CreatePartUri(new Uri("/customXml/htmlData.xml", UriKind.Relative));
                    if (pkg.PartExists(uri))
                    {
                        const string relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
                        var rel = pkg.GetRelationshipsByType(relType).FirstOrDefault(r => r.TargetUri == uri);
                        if (rel != null) pkg.DeleteRelationship(rel.Id);
                        pkg.DeletePart(uri);
                    }
                }
                if (OfficeSupportTool.ReadStoredHtml(file) != null) return "foreign docx reported metadata";
                return null;
            }
            finally { try { File.Delete(file); } catch { } }
        });

        failures += Test("html: nested comment detection", () =>
        {
            var nested = "<!-- banner\n# - table: <!-- SLA-ROW --> marks rows\n-->\n<p>x</p>";
            if (!OfficeSupportTool.HasNestedComments(nested)) return "nested comment not detected";
            var legit = "<!-- banner -->\n<!-- SLA-ROW -->\n<tr><td>x</td></tr>\n<!-- end -->";
            if (OfficeSupportTool.HasNestedComments(legit)) return "legit flat comments flagged";
            if (OfficeSupportTool.HasNestedComments("<p>plain text <!-- a --> only</p>")) return "single comment flagged";
            return null;
        });

        failures += Test("html: bare svg size detection", () =>
        {
            if (!OfficeSupportTool.HasBareSvgSizes("<svg width=\"46\" height=\"46\" viewBox=\"0 0 24 24\"><g/></svg>")) return "bare width not detected";
            if (OfficeSupportTool.HasBareSvgSizes("<svg width=\"46px\" height=\"46px\" viewBox=\"0 0 24 24\"><g/></svg>")) return "px units wrongly flagged";
            if (OfficeSupportTool.HasBareSvgSizes("<svg viewBox=\"0 0 24 24\"><g/></svg>")) return "no-size svg wrongly flagged";
            if (OfficeSupportTool.HasBareSvgSizes("<svg width=\"46.5px\" height=\"46px\"><g/></svg>")) return "decimal px wrongly flagged";
            return null;
        });

        failures += Test("inspect: table background style-form detected", () =>
        {
            var dir = Path.Combine(Path.GetTempPath(), "ost-inspect-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                var f = Path.Combine(dir, "bad.html");
                File.WriteAllText(f, "<html><body><table style=\"background-color:#FFFFFF;\"><tr><td>x</td></tr></table></body></html>");
                if (InspectTemplate(f, "selftest") == 0) return "style-form background on <table> not flagged";
                File.WriteAllText(f, "<html><body><table bgcolor=\"#FFFFFF\"><tr><td>x</td></tr></table></body></html>");
                if (InspectTemplate(f, "selftest") == 0) return "bgcolor attribute on <table> not flagged";
                File.WriteAllText(f, "<html><body><table><tr style=\"background-color:#F8FAFC;\"><td>x</td></tr></table></body></html>");
                if (InspectTemplate(f, "selftest") > 0) return "legit tr background wrongly flagged";
                return null;
            }
            finally { try { Directory.Delete(dir, true); } catch { } }
        });

        failures += Test("html: table background detection", () =>
        {
            if (!OfficeSupportTool.HasTableBackground("<table style=\"background-color:#FFF;\"><tr><td>x</td></tr></table>")) return "style-form not detected";
            if (!OfficeSupportTool.HasTableBackground("<table bgcolor=\"#FFF\"><tr><td>x</td></tr></table>")) return "bgcolor not detected";
            if (OfficeSupportTool.HasTableBackground("<table><tr style=\"background-color:#F8FAFC;\"><td>x</td></tr></table>")) return "tr background wrongly flagged";
            return null;
        });

        failures += Test("html: svg size normalization (bare -> px)", () =>
        {
            var html = "<svg width=\"46\" height=\"46\" viewBox=\"0 0 24 24\"><g/></svg>" +
                       "<img src=\"data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjQiIGhlaWdodD0iMjQiLz4=\">" +
                       "<img src=\"chart.png\">";
            var outHtml = OfficeSupportTool.NormalizeSvgSizes(html);
            if (!outHtml.Contains("width=\"46px\" height=\"46px\"")) return "inline svg not normalized";
            var b64 = Regex.Match(outHtml, "base64,([^\"]+)\"").Groups[1].Value;
            var svg = Encoding.UTF8.GetString(Convert.FromBase64String(b64));
            if (!svg.Contains("width=\"24px\" height=\"24px\"")) return $"data-uri svg not normalized: {svg}";
            if (!outHtml.Contains("src=\"chart.png\"")) return "non-svg img must stay untouched";
            return null;
        });

        failures += Test("icons: placeholder embedding (size + color + paths)", () =>
        {
            var iconsDir = Path.Combine(Path.GetTempPath(), "ost-selftest-icons-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(iconsDir);
            try
            {
                File.WriteAllText(Path.Combine(iconsDir, "disc.svg"),
                    "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"24\" height=\"24\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><circle cx=\"12\" cy=\"12\" r=\"10\"/></svg>");
                var sample = "<img src=\"disc.32.aa0000.svg\" alt=\"disc\">" +
                             "<img src=\"/icons/disc.24.svg\">" +
                             "<img src=\"zzzz.16.123456.svg\">";
                var outHtml = Utility.EmbedSvgIcons(sample, iconsDir);
                if (Regex.Matches(outHtml, "src=\"data:image/svg\\+xml;base64,").Count != 2)
                    return $"expected 2 data-URI srcs, got: {outHtml}";
                if (!outHtml.Contains("zzzz.16.123456.svg")) return "unknown icon must stay unresolved";
                var first = Regex.Match(outHtml, "src=\"data:image/svg\\+xml;base64,([^\"]+)\"");
                var svg = Encoding.UTF8.GetString(Convert.FromBase64String(first.Groups[1].Value));
                if (!svg.Contains("width=\"32px\"")) return $"size 32 not applied: {svg}";
                if (!svg.Contains("stroke=\"#aa0000\"")) return $"color aa0000 not applied: {svg}";
                return null;
            }
            finally { try { Directory.Delete(iconsDir, true); } catch { } }
        });

        failures += Test("restore: named backup + swap (current becomes a backup)", () =>
        {
            var dir = Path.Combine(Path.GetTempPath(), "ost-restore-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            var saved = Setup.DocumentsPath;
            Setup.DocumentsPath = dir;
            try
            {
                var f = Path.Combine(dir, "invoice.docx");
                File.WriteAllBytes(f, OfficeSupportTool.ConvertToDocx("<html><body><h1>Version 1</h1></body></html>"));
                File.Copy(f, Path.Combine(dir, "invoice.001.bak"));
                File.WriteAllBytes(f, OfficeSupportTool.ConvertToDocx("<html><body><h1>Version 2</h1></body></html>"));
                var r = new OfficeSupportTool().Restore("invoice.001.bak");
                if (!r.StartsWith("Document restored at") || !r.Contains("invoice.001.bak")) return $"restore result: {r}";
                if (!DocxTextContains(f, "Version 1")) return "restored content is not Version 1";
                if (!File.Exists(Path.Combine(dir, "invoice.002.bak"))) return "swap backup invoice.002.bak not created";
                if (!DocxTextContains(Path.Combine(dir, "invoice.002.bak"), "Version 2")) return "swap backup does not hold Version 2";
                return null;
            }
            finally { Setup.DocumentsPath = saved; try { Directory.Delete(dir, true); } catch { } }
        });

        failures += Test("restore: no-arg picks the newest backup in the workspace", () =>
        {
            var dir = Path.Combine(Path.GetTempPath(), "ost-restore-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(dir, "sub"));
            var saved = Setup.DocumentsPath;
            Setup.DocumentsPath = dir;
            try
            {
                var fa = Path.Combine(dir, "a.docx");
                File.WriteAllBytes(fa, OfficeSupportTool.ConvertToDocx("<html><body><p>A1</p></body></html>"));
                File.Copy(fa, Path.Combine(dir, "a.001.bak"));
                File.SetLastWriteTimeUtc(Path.Combine(dir, "a.001.bak"), new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc));
                var fb = Path.Combine(dir, "sub", "b.docx");
                File.WriteAllBytes(fb, OfficeSupportTool.ConvertToDocx("<html><body><p>B1</p></body></html>"));
                File.Copy(fb, Path.Combine(dir, "sub", "b.001.bak"));
                File.SetLastWriteTimeUtc(Path.Combine(dir, "sub", "b.001.bak"), new DateTime(2026, 1, 2, 0, 0, 0, DateTimeKind.Utc));
                File.WriteAllBytes(fb, OfficeSupportTool.ConvertToDocx("<html><body><p>B2</p></body></html>"));
                var r = new OfficeSupportTool().Restore();
                if (!r.Contains("b.001.bak")) return $"expected b.001.bak, got: {r}";
                if (!DocxTextContains(fb, "B1")) return "b.docx not restored to B1";
                return null;
            }
            finally { Setup.DocumentsPath = saved; try { Directory.Delete(dir, true); } catch { } }
        });

        failures += Test("restore: error paths (no backups, missing file, non-backup name)", () =>
        {
            var dir = Path.Combine(Path.GetTempPath(), "ost-restore-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            var saved = Setup.DocumentsPath;
            Setup.DocumentsPath = dir;
            try
            {
                var t = new OfficeSupportTool();
                var r1 = t.Restore();
                if (!r1.StartsWith("Error: no backup file")) return $"r1: {r1}";
                var f = Path.Combine(dir, "x.docx");
                File.WriteAllBytes(f, OfficeSupportTool.ConvertToDocx("<html><body><p>x</p></body></html>"));
                File.Copy(f, Path.Combine(dir, "x.001.bak"));
                var r2 = t.Restore("zzz.001.bak");
                if (!r2.StartsWith("Error: backup file 'zzz.001.bak' not found")) return $"r2: {r2}";
                var r3 = t.Restore("x.docx");
                if (!r3.StartsWith("Error: cannot derive")) return $"r3: {r3}";
                return null;
            }
            finally { Setup.DocumentsPath = saved; try { Directory.Delete(dir, true); } catch { } }
        });

        Console.WriteLine(failures == 0 ? "  ALL SELF-TESTS PASSED" : $"  {failures} SELF-TEST FAILURES");
        return failures == 0 ? 0 : 1;
    }

    static int Test(string id, Func<string?> run)
    {
        try
        {
            var problem = run();
            if (problem == null) { Console.WriteLine($"  ✓ {id} PASS"); return 0; }
            Console.WriteLine($"  ✗ {id} FAIL: {problem}");
            return 1;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ✗ {id} CRASH: {ex.GetType().Name}: {ex.Message}");
            return 1;
        }
    }

    // ---------- fixtures ----------

    /// <summary>Ensures the requested provider exists. DeepSeekBridge is preconfigured;
    /// 'Ollama_Qwen' (local qwen3.5:4b) is registered at runtime when requested but absent,
    /// keeping providers.json untouched.</summary>
    static void EnsureProvider()
    {
        if (ProviderConfigs.TryGet(_providerName, out _)) return;
        if (!string.Equals(_providerName, "Ollama_Qwen", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException($"Unknown provider '{_providerName}'. Use --provider with a configured provider or 'Ollama_Qwen'.");
        ProviderConfigs.Add(new ProviderConfig
        {
            ProviderName = "Ollama_Qwen",
            Protocol = ProviderProtocol.OpenAI,
            CacheType = ProviderCacheType.PrefixCache,
            ModelName = "qwen3.5:4b",
            BaseAddress = new Uri("http://localhost:11434/"),
            EndPoint = "v1/chat/completions",
            Timeout = TimeSpan.FromMinutes(40),
            PauseBetweenRequests = TimeSpan.Zero,
            ContextWindow = 262144,
        }, persist: false);
    }

    /// <summary>Stages a small icon set in the harness output so the plugin-level icon embedding
    /// (AppContext.BaseDirectory/assets/icons) works during tests.</summary>
    static void StageIcons()
    {
        var iconsDir = Path.Combine(AppContext.BaseDirectory, "assets", "icons");
        Directory.CreateDirectory(iconsDir);
        var icons = new Dictionary<string, string>
        {
            ["disc"] = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"24\" height=\"24\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><circle cx=\"12\" cy=\"12\" r=\"10\"/></svg>",
            ["file"] = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"24\" height=\"24\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><path d=\"M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z\"/><polyline points=\"14 2 14 8 20 8\"/></svg>",
            ["users"] = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"24\" height=\"24\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><path d=\"M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2\"/><circle cx=\"9\" cy=\"7\" r=\"4\"/><path d=\"M23 21v-2a4 4 0 0 0-3-3.87\"/><path d=\"M16 3.13a4 4 0 0 1 0 7.75\"/></svg>",
        };
        foreach (var (name, svg) in icons)
            File.WriteAllText(Path.Combine(iconsDir, name + ".svg"), svg);
    }

    /// <summary>Stages two small solid-color PNGs in the workspace so the images flow can be
    /// exercised (CreateDocument with imageFiles: the LLM must place each image once).</summary>
    static void StageImages()
    {
        var dir = Path.Combine(_workspace, "images");
        Directory.CreateDirectory(dir);
        WritePng(Path.Combine(dir, "logo.png"), 64, 64, "#8B1A1A");
        WritePng(Path.Combine(dir, "coffee.png"), 72, 48, "#6F4E37");
    }

    static void WritePng(string path, int w, int h, string hex)
    {
        using var img = new Image<Rgba32>(w, h);
        img.Mutate(x => x.BackgroundColor(SixLabors.ImageSharp.Color.ParseHex(hex)));
        img.SaveAsPng(path);
    }
}
