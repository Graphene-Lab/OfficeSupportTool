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
using UISupportGeneric;

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
        if (Array.IndexOf(args, "--utests") >= 0) return RunUTests();
        EnsureProvider();

        InitWorkspace();

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
            if (!r1b.Contains("New version:", StringComparison.OrdinalIgnoreCase))
            { Fail("T1-overwrite-backup", $"expected version report: {r1b}"); return 1; }
            Pass("T1-overwrite-backup");

            // T2 — update the document via the embedded HTML metadata
            var r2 = tool.UpdateDocument(
                "/balance-sheet.docx",
                "Change the balance sheet date to 31 March 2026 and increase accounts receivable to 150,000, " +
                "updating total assets and the accounting identity consistently.");
            Console.WriteLine($"  T2 UpdateDocument → {r2}");
            if (!r2.StartsWith("Document updated at") || !r2.Contains("New version:"))
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
                "TASK 3: The user now wants to UNDO the change of TASK 2 (rollback to the version before the change), using GitTool.restore with the version the tool reported (history() lists all versions). " +
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
            { Fail("T7-rollback", $"agent never called GitTool.restore — trace: {string.Join(" | ", trace7)}"); return 1; }
            if (GitSupport.History(rollbackFile).Count < 3)
            { Fail("T7-rollback", "version chain shorter than expected (create + update + restore swap)"); return 1; }
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

    /// <summary>Fresh %TEMP% workspace + staging (icons, images) for a test run. The repo sits
    /// under OneDrive: test files must never be written under the repo (cloud-synced on every
    /// write — historical slow runs).</summary>
    static void InitWorkspace()
    {
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
    }

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

    /// <summary>Multi-turn variant: each element is a separate user request on the SAME
    /// conversation (history shared), mirroring how a user asks a follow-up in one chat. The trace
    /// accumulates the tool calls of every turn.</summary>
    static (AgentResult Result, List<string> Trace) RunAgentConversation(string[] turns, int maxIterations = 40)
    {
        using var orch = new AgentHarness(_providerName);
        var calls = new List<string>();
        orch.AgentProgress += (_, e) =>
        {
            if (e.State == AgentHarness.AgentState.Iteration && !string.IsNullOrWhiteSpace(e.MethodName))
                calls.Add($"#{e.Iteration}:{e.MethodName}");
        };
        AgentResult result = null!;
        foreach (var turn in turns)
            result = orch.ExecuteAction(turn, new[] { "OfficeSupportTool" }, maxIterations: maxIterations);
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

        failures += Test("templates: ALL shipped templates conform (svg px, no table bg, flat comments)", () =>
        {
            var dir = Path.Combine(AppContext.BaseDirectory, "assets", "templates");
            if (!Directory.Exists(dir)) return "templates dir missing";
            var files = Directory.GetFiles(dir, "*.html");
            if (files.Length < 40) return $"expected ~49 shipped templates, found {files.Length}";
            var issues = 0;
            foreach (var f in files)
            {
                var tpl = File.ReadAllText(f);
                if (OfficeSupportTool.HasNestedComments(tpl)) { issues++; Console.WriteLine($"    nested comments: {Path.GetFileName(f)}"); }
                if (OfficeSupportTool.HasTableBackground(tpl)) { issues++; Console.WriteLine($"    table background: {Path.GetFileName(f)}"); }
                if (OfficeSupportTool.HasBareSvgSizes(tpl)) { issues++; Console.WriteLine($"    bare svg size: {Path.GetFileName(f)}"); }
                if (!tpl.Contains("{{")) Console.WriteLine($"    ⚠ no {{placeholder}}: {Path.GetFileName(f)}");
            }
            return issues == 0 ? null : $"{issues} shipped template(s) with violations";
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

        failures += Test("definitions: dynamic template list resolves", () =>
        {
            var defs = Analyzer.GeToolDefinitions(typeof(OfficeSupportTool));
            if (defs.Contains("[[available_templates]]")) return "placeholder not resolved";
            if (!defs.Contains("balance-sheet", StringComparison.OrdinalIgnoreCase)) return "resolved list missing a shipped template";
            return null;
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

        failures += Test("docx: icon-name placeholder end-to-end (embed + convert to image part)", () =>
        {
            var iconsDir = Path.Combine(Path.GetTempPath(), "ost-ico-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(iconsDir);
            try
            {
                File.WriteAllText(Path.Combine(iconsDir, "disc.svg"),
                    "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"24\" height=\"24\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><circle cx=\"12\" cy=\"12\" r=\"10\"/></svg>");
                var html = "<!DOCTYPE html><html lang=\"en\"><head><meta charset=\"utf-8\"/></head><body>" +
                           "<p style=\"font-size:10pt;\">Test <img src=\"disc.32.aa0000.svg\" alt=\"disc\"></p></body></html>";
                var embedded = Utility.EmbedSvgIcons(html, iconsDir);
                if (!embedded.Contains("data:image/svg+xml;base64,")) return "icon placeholder not embedded as data URI";
                var file = Path.Combine(Path.GetTempPath(), "ost-ico-" + Guid.NewGuid().ToString("N") + ".docx");
                try
                {
                    var bytes = OfficeSupportTool.ConvertToDocx(embedded);
                    File.WriteAllBytes(file, bytes);
                    if (DocxImageCount(file) < 1) return "embedded icon not converted to an image part in the docx";
                    return null;
                }
                finally { try { File.Delete(file); } catch { } }
            }
            finally { try { Directory.Delete(iconsDir, true); } catch { } }
        });

        failures += Test("versioning: snapshot + rollback on .docx (swap preserved, unknown version rejected)", () =>
        {
            var dir = Path.Combine(Path.GetTempPath(), "ost-restore-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            var saved = Setup.DocumentsPath;
            Setup.DocumentsPath = dir;
            try
            {
                var f = Path.Combine(dir, "invoice.docx");
                File.WriteAllBytes(f, OfficeSupportTool.ConvertToDocx("<html><body><h1>Version 1</h1></body></html>"));
                var v1 = GitSupport.Snapshot(f, "V1");
                File.WriteAllBytes(f, OfficeSupportTool.ConvertToDocx("<html><body><h1>Version 2</h1></body></html>"));
                GitSupport.Snapshot(f, "V2");
                if (GitSupport.History(f).Count != 2) return $"history count: {GitSupport.History(f).Count}";
                // Pending edit (NOT yet versioned): the restore swap must capture it before overwriting.
                File.WriteAllBytes(f, OfficeSupportTool.ConvertToDocx("<html><body><h1>Version 3</h1></body></html>"));
                var r = GitSupport.Restore(v1, f);
                if (!r.StartsWith("Restored")) return $"restore result: {r}";
                if (!DocxTextContains(f, "Version 1")) return "restored content is not Version 1";
                var v3 = GitSupport.History(f)[0].VersionId;   // the swap snapshot of the pending V3
                if (GitSupport.History(f).Count != 3) return "swap snapshot missing after restore";
                GitSupport.Restore(v3, f);                      // rollback of the rollback
                if (!DocxTextContains(f, "Version 3")) return "swap snapshot not restorable";
                var failed = false;
                try { GitSupport.Restore("deadbeef", f); } catch (InvalidOperationException) { failed = true; }
                if (!failed) return "unknown version not rejected";
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

    // ---------- untested-flows campaign (--utests) ----------

    /// <summary>Behavioral tests for usage sequences never covered by the T-series: a template that
    /// expects images (product catalog) compiled with real images, update with images, rollback of a
    /// SPECIFIC document among several with backups, create from a context file, template reuse
    /// across calls, chained updates + restore to the first backup, and deterministic error paths.
    /// Each agent-level test records the tool-call trace (what the agent actually did).</summary>
    static int RunUTests()
    {
        EnsureProvider();
        InitWorkspace();
        File.WriteAllText(ResultsFile, $"RUN-UTEST {DateTime.Now:HH:mm:ss} provider={_providerName}\n");
        WriteResult("STARTED");

        Console.WriteLine("══════════ OfficeSupportTool untested-flows test ══════════");
        Console.WriteLine($"provider: {_providerName}");
        Log.LogStep($"=== OfficeSupportTool untested-flows (provider {_providerName}) ===");

        var tool = new OfficeSupportTool();
        try
        {
            // U1 — a template that expects images (product catalog) compiled with product photos
            Console.WriteLine("  U1 product catalog + images (agent-level)...");
            var (r1, t1) = RunAgent(
                "Create a product catalog for 'Fiori Coffee S.r.l.' (via Roma 12, 20121 Milano, Italy, VAT IT01234567890, email info@fioricoffee.it) with 3 products " +
                "in one category 'Coffee': 'Espresso Blend' (FB-01, 250g, EUR 12.50), 'Arabica Single Origin' (FB-02, 250g, EUR 15.00), 'Decaf' (FB-03, 250g, EUR 13.00). " +
                "The catalog must include the product photo of each product: coffee.png, beans.png and cup.png (they are in /images/, pass them as imageFiles, e.g. '/images/coffee.png', and place each photo in its product row). " +
                "Additional material: company tagline 'Specialty coffee since 1998'; subtitle 'Autumn collection 2026'; introduction 'A selection of our finest single-origin and blended coffees'; " +
                "per-product description and specifications; ordering info 'orders@fioricoffee.it, 2-week lead time, minimum order 10 units'; payment terms '30 days net'; sales contact 'Anna Verdi, Sales Manager'; catalog version 1.0. " +
                "If any template field is still missing from the material, set the draft parameter and generate a sensible value. Save as /catalog.docx.",
                maxIterations: 40);
            WriteResult("U1 trace: " + string.Join(" | ", t1));
            Console.WriteLine($"  U1 tool calls: {string.Join(" → ", t1)}");
            if (r1.Error != null) { Fail("U1-catalog-images", $"agent error: {r1.Error}"); return 1; }
            var host1 = Path.Combine(_workspace, "catalog.docx");
            var html1 = File.Exists(host1) ? OfficeSupportTool.ReadStoredHtml(host1) : null;
            if (html1 == null) { Fail("U1-catalog-images", "catalog.docx not created"); return 1; }
            var pngCount1 = Regex.Matches(html1, "src=\"data:image/png;base64,").Count;
            if (pngCount1 != 3) { Fail("U1-catalog-images", $"expected 3 embedded product images, found {pngCount1}"); return 1; }
            if (DocxImageCount(host1) < 3) { Fail("U1-catalog-images", "converted DOCX has fewer than 3 image parts"); return 1; }
            if (!DocxTextContains(host1, "Espresso Blend")) { Fail("U1-catalog-images", "catalog content missing"); return 1; }
            Pass("U1-catalog-images");

            // U2 — update WITH images: the user asks, in a follow-up turn, to add the logo to the
            // document created in the previous turn (realistic multi-turn usage)
            Console.WriteLine("  U2 update + images (agent-level, multi-turn)...");
            var (r2, t2) = RunAgentConversation(new[]
            {
                "Create a business letter from 'Fiori Coffee S.r.l.' to 'Nova S.p.A.' thanking them for the partnership, save as /update-img.docx. " +
                "Supporting material to pass as contextText: Fiori Coffee S.r.l. — via Roma 12, 20121 Milano, Italy, VAT IT01234567890, email info@fioricoffee.it, website www.fioricoffee.it. " +
                "Letter date 18 August 2026, reference FI-2026-077. Recipient: Maria Conti, Procurement Director, Nova S.p.A., via Torino 45, 10122 Torino, Italy. " +
                "Subject 'Partnership renewal', salutation 'Dear Ms. Conti'. Body: Nova S.p.A. has been a partner since 2019; we thank you for the partnership. " +
                "Closing 'Best regards', sender Luca Bianchi (CEO), enclosures: product catalogue 2026.",
                "The user now asks: add the company logo (logo.png, in /images/) at the top of the letter — the document is /update-img.docx, modify it (pass the image via imageFiles)."
            });
            WriteResult("U2 trace: " + string.Join(" | ", t2));
            Console.WriteLine($"  U2 tool calls: {string.Join(" → ", t2)}");
            if (r2.Error != null) { Fail("U2-update-images", $"agent error: {r2.Error}"); return 1; }
            var host2 = Path.Combine(_workspace, "update-img.docx");
            var html2 = File.Exists(host2) ? OfficeSupportTool.ReadStoredHtml(host2) : null;
            if (html2 == null) { Fail("U2-update-images", "update-img.docx not created"); return 1; }
            if (Regex.Matches(html2, "src=\"data:image/png;base64,").Count < 1) { Fail("U2-update-images", "logo not embedded after update"); return 1; }
            if (GitSupport.History(host2).Count < 2) { Fail("U2-update-images", "update did not create a new version"); return 1; }
            Pass("U2-update-images");

            // U3 — rollback of a SPECIFIC document among several documents with backups. The setup
            // (create + update of both letters) is done with direct tool calls so the test isolates
            // the agent's rollback behavior — the real target — from the flaky create/material handling.
            Console.WriteLine("  U3 rollback on a specific document (agent-level, deterministic setup)...");
            var sA1 = tool.CreateDocument("business letter", "Business letter from Fiori Coffee to Nova S.p.A. thanking them for the partnership.", draft: true,
                contextText: "Fiori Coffee S.r.l., via Roma 12, 20121 Milano, VAT IT01234567890; recipient Maria Conti, Nova S.p.A., via Torino 45, Torino; " +
                "date 18 August 2026; subject 'Partnership renewal'; salutation 'Dear Ms. Conti'; body 'we thank you for the partnership started in 2019'; " +
                "closing 'Best regards'; sender Luca Bianchi (CEO).",
                saveFullNameFile: "/letterA.docx");
            if (!sA1.StartsWith("Document created at")) { Fail("U3-setup", $"letterA create: {sA1}"); return 1; }
            var sA2 = tool.UpdateDocument("/letterA.docx", "Replace 'started in 2019' with 'started in 2020'.");
            if (!sA2.StartsWith("Document updated at")) { Fail("U3-setup", $"letterA update: {sA2}"); return 1; }
            var sB1 = tool.CreateDocument("business letter", "Business letter from Fiori Coffee to Beta S.r.l. confirming the order.", draft: true,
                contextText: "Fiori Coffee S.r.l., via Roma 12, 20121 Milano, VAT IT01234567890; recipient Giulia Neri, Beta S.r.l., via Firenze 8, Roma; " +
                "date 18 August 2026; salutation 'Dear Ms. Neri'; body 'we confirm the order for 500 kg of coffee'; closing 'Best regards'; sender Luca Bianchi (CEO).",
                saveFullNameFile: "/letterB.docx");
            if (!sB1.StartsWith("Document created at")) { Fail("U3-setup", $"letterB create: {sB1}"); return 1; }
            var sB2 = tool.UpdateDocument("/letterB.docx", "Replace '500 kg' with '750 kg'.");
            if (!sB2.StartsWith("Document updated at")) { Fail("U3-setup", $"letterB update: {sB2}"); return 1; }
            var (r3, t3) = RunAgent(
                "The user asks: undo the change made to /letterA.docx — rollback it to the version before it was modified, using GitTool.restore with the version the tool reported (history() lists the versions). Do NOT touch /letterB.docx.",
                maxIterations: 30);
            WriteResult("U3 trace: " + string.Join(" | ", t3));
            Console.WriteLine($"  U3 tool calls: {string.Join(" → ", t3)}");
            if (r3.Error != null) { Fail("U3-rollback-specific", $"agent error: {r3.Error}"); return 1; }
            var htmlA = OfficeSupportTool.ReadStoredHtml(Path.Combine(_workspace, "letterA.docx"));
            var htmlB = OfficeSupportTool.ReadStoredHtml(Path.Combine(_workspace, "letterB.docx"));
            if (htmlA == null || htmlB == null) { Fail("U3-rollback-specific", "documents missing"); return 1; }
            if (htmlA.Contains("started in 2020")) { Fail("U3-rollback-specific", "letterA still contains the change — rollback failed"); return 1; }
            if (!htmlA.Contains("started in 2019")) { Fail("U3-rollback-specific", "letterA original wording missing"); return 1; }
            if (!htmlB.Contains("750 kg")) { Fail("U3-rollback-specific", "letterB was wrongly touched (must keep its change)"); return 1; }
            Pass("U3-rollback-specific");

            // U4 — create from a context FILE
            Console.WriteLine("  U4 create with contextFile (agent-level)...");
            var ctxDir = Path.Combine(_workspace, "context");
            Directory.CreateDirectory(ctxDir);
            File.WriteAllText(Path.Combine(ctxDir, "company.md"),
                "Fiori Coffee S.r.l. — via Roma 12, 20121 Milano, Italy. VAT IT01234567890, phone +39 02 12345678, email info@fioricoffee.it, website www.fioricoffee.it. Company tagline: 'Specialty coffee since 1998'. " +
                "Income statement for the fiscal year ended 31 December 2025, prepared on 18 August 2026 by Maria Rossi (Chief Accountant), approved by Luca Bianchi (CEO). Status: final. Currency: EUR. " +
                "Revenue 1,240,000; cost of goods sold 610,000; gross profit 630,000; operating expenses 320,000 (of which: salaries 180,000, rent 60,000, marketing 50,000, utilities 30,000); " +
                "operating income 310,000; other income 5,000; interest expense 8,000; income before taxes 307,000; taxes 85,000; net income 222,000. " +
                "Previous year 2024: revenue 1,050,000; net income 175,000. Key ratios: gross margin 50.8%; net margin 17.9%. " +
                "Notes: figures in EUR, prepared in accordance with IFRS.");
            var (r4, t4) = RunAgent(
                "Create an income statement for Fiori Coffee for the year 2025. The company data and figures are in the workspace file /context/company.md — pass it as contextFile. " +
                "If any template field is still missing, set the draft parameter and generate a sensible value. Save as /income-statement.docx.",
                maxIterations: 30);
            WriteResult("U4 trace: " + string.Join(" | ", t4));
            Console.WriteLine($"  U4 tool calls: {string.Join(" → ", t4)}");
            if (r4.Error != null) { Fail("U4-contextFile", $"agent error: {r4.Error}"); return 1; }
            var html4 = File.Exists(Path.Combine(_workspace, "income-statement.docx")) ? OfficeSupportTool.ReadStoredHtml(Path.Combine(_workspace, "income-statement.docx")) : null;
            if (html4 == null) { Fail("U4-contextFile", "income-statement.docx not created"); return 1; }
            if (!html4.Contains("1,240,000")) { Fail("U4-contextFile", "income statement data missing"); return 1; }
            Pass("U4-contextFile");

            // U5 — template reuse across calls: the second create of the same unknown type must reuse
            // the template saved by the first (no regeneration), visible in the log
            Console.WriteLine("  U5 template reuse across calls (agent-level)...");
            var (r5, t5) = RunAgent(
                "Complete ALL tasks in order; do not stop early. " +
                "TASK 1: Create a 'vendor risk register' document for Fiori Coffee (draft ok) listing risk categories Financial, Operational, Compliance — save as /risk1.docx. " +
                "TASK 2: Create ANOTHER 'vendor risk register' document for Fiori Coffee (draft ok) — the same document type — with risk categories Financial, Operational, Reputational — save as /risk2.docx.",
                maxIterations: 40);
            WriteResult("U5 trace: " + string.Join(" | ", t5));
            Console.WriteLine($"  U5 tool calls: {string.Join(" → ", t5)}");
            if (r5.Error != null) { Fail("U5-template-reuse", $"agent error: {r5.Error}"); return 1; }
            if (!File.Exists(Path.Combine(_workspace, "risk2.docx"))) { Fail("U5-template-reuse", "risk2.docx not created"); return 1; }
            if (Log.CurrentLogFile == null || !File.Exists(Log.CurrentLogFile)) { Fail("U5-template-reuse", "log file unavailable"); return 1; }
            var logText = File.ReadAllText(Log.CurrentLogFile);
            var genCount = Regex.Matches(logText, "no template for 'vendor-risk-register'").Count;
            if (genCount != 1) { Fail("U5-template-reuse", $"expected exactly 1 template generation, found {genCount} (second create did not reuse the saved template)"); return 1; }
            Pass("U5-template-reuse");

            // U6 — chained updates + restore to the FIRST version. The memo is created and updated
            // twice with direct tool calls; the agent handles only the restore-to-first-backup turn.
            Console.WriteLine("  U6 chained updates + restore to first version (agent-level, deterministic setup)...");
            var sM1 = tool.CreateDocument("memorandum", "Memorandum about the holiday schedule.", draft: true,
                contextText: "Fiori Coffee S.r.l., via Roma 12, 20121 Milano, VAT IT01234567890; sender Luca Bianchi (CEO); to all staff; " +
                "subject 'Holiday schedule'; body 'The office will close from 24 December to 2 January.'; date 18 August 2026.",
                saveFullNameFile: "/memo.docx");
            if (!sM1.StartsWith("Document created at")) { Fail("U6-setup", $"memo create: {sM1}"); return 1; }
            var sM2 = tool.UpdateDocument("/memo.docx", "Change the closing period to '24 December to 6 January'.");
            if (!sM2.StartsWith("Document updated at")) { Fail("U6-setup", $"memo update 1: {sM2}"); return 1; }
            var sM3 = tool.UpdateDocument("/memo.docx", "Change the subject to 'Holiday schedule 2026'.");
            if (!sM3.StartsWith("Document updated at")) { Fail("U6-setup", $"memo update 2: {sM3}"); return 1; }
            var (r6, t6) = RunAgent(
                "The user asks: go back to the ORIGINAL version of /memo.docx (before the first modification) — restore it using GitTool.restore with the FIRST version of the file (history() lists them oldest last).",
                maxIterations: 30);
            WriteResult("U6 trace: " + string.Join(" | ", t6));
            Console.WriteLine($"  U6 tool calls: {string.Join(" → ", t6)}");
            if (r6.Error != null) { Fail("U6-chained-restore", $"agent error: {r6.Error}"); return 1; }
            var host6 = Path.Combine(_workspace, "memo.docx");
            var html6 = OfficeSupportTool.ReadStoredHtml(host6);
            if (html6 == null) { Fail("U6-chained-restore", "memo.docx not created"); return 1; }
            if (!html6.Contains("24 December to 2 January")) { Fail("U6-chained-restore", "memo not restored to the original period"); return 1; }
            if (html6.Contains("6 January")) { Fail("U6-chained-restore", "memo still contains the first change"); return 1; }
            if (GitSupport.History(host6).Count < 4) { Fail("U6-chained-restore", "version chain shorter than expected (create + 2 updates + restore swap)"); return 1; }
            Pass("U6-chained-restore");

            // U7 — deterministic error paths (direct calls, no LLM)
            File.WriteAllText(Path.Combine(_workspace, "images", "bad.txt"), "not an image");
            var u7a = tool.CreateDocument("invoice", "test", draft: true, contextText: "x", imageFiles: new[] { "/images/logo.png", "/images/bad.txt" });
            if (!u7a.StartsWith("Error:") || !u7a.Contains("unsupported image type")) { Fail("U7-error-image-type", $"unexpected: {u7a}"); return 1; }
            Pass("U7-error-image-type");
            var u7b = tool.CreateDocument("invoice", "test", draft: true, contextText: "x", contextFile: "/nope/missing.md");
            if (!u7b.StartsWith("Error:") || !u7b.Contains("not found")) { Fail("U7-error-context-file", $"unexpected: {u7b}"); return 1; }
            Pass("U7-error-context-file");
            var u7c = tool.CreateDocument("invoice", "test", draft: true, contextText: "x", saveFullNameFile: "/out/foo.txt");
            if (!u7c.StartsWith("Error:") || !u7c.Contains(".docx")) { Fail("U7-error-extension", $"unexpected: {u7c}"); return 1; }
            Pass("U7-error-extension");
            var u7d = tool.UpdateDocument("/missing.docx", "change x");
            if (!u7d.StartsWith("Error:") || !u7d.Contains("not found")) { Fail("U7-error-update-missing", $"unexpected: {u7d}"); return 1; }
            Pass("U7-error-update-missing");

            Console.WriteLine();
            Console.WriteLine(_failures == 0 ? "  ALL UNTESTED-FLOW TESTS PASSED" : $"  {_failures} UTEST FAILURES");
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
        WritePng(Path.Combine(dir, "beans.png"), 64, 64, "#3E7C17");
        WritePng(Path.Combine(dir, "cup.png"), 56, 64, "#2563EB");
    }

    static void WritePng(string path, int w, int h, string hex)
    {
        using var img = new Image<Rgba32>(w, h);
        img.Mutate(x => x.BackgroundColor(SixLabors.ImageSharp.Color.ParseHex(hex)));
        img.SaveAsPng(path);
    }
}
