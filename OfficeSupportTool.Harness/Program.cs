using System.Text;
using System.Text.RegularExpressions;
using System.IO.Packaging;
using AIOrchestrator;
using AIOrchestrator.API;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
}
