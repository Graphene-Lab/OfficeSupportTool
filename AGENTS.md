# OfficeSupportTool — notes for coding agents

> Read this file first. This repo is one tool plugin of the AIOffice solution: it is compiled on
> its own and loaded by the hosts from `Tools/OfficeSupportTool/` at runtime.

## What this is

DOCX documents built and updated from HTML templates through `HtmlToOpenXml`.

## Read before you edit

| If you touch… | Read first |
|---|---|
| anything | [PLUGIN-GUARDRAILS](../AIOrchestrator/docs-dev/PLUGIN-GUARDRAILS.md) — the contract shared by the 12 plugins; the file lives in the sibling AIOrchestrator repo (private), reachable from the solution checkout |
| the HTML→DOCX conversion | [HtmlToOpenXml-Guide.md](HtmlToOpenXml-Guide.md) |
| a tool method or its XML docs | [AGENT_TOOLS_GUIDE](../AIOrchestrator/API/AGENT_TOOLS_GUIDE.md) |

## Non-negotiables

1. The class derives from `BaseAgentTool` and keeps
   `[ToolPurpose(ToolPurpose.DocumentTemplates)]`.
2. Public methods and their XML docs **are** the agent's surface: no public method without docs,
   and a public method is never renamed (the model calls it by that name).
3. AnyCPU, no `RuntimeIdentifier`, `GenerateDocumentationFile` stays on.
4. The plugin folder never carries `AIOrchestrator.dll`; no state is written into `Tools/`.
5. A new metadata goes in a class attribute (or a `virtual` member with a default) — never an
   `abstract` member: the dlls already shipped must keep loading.

## Where things are

- `OfficeSupportTool.cs` — the tool class; its public methods are what the model can call
- `Assets/` — the background/template assets (SVG/VML templates live here)
- `OfficeSupportTool.Harness/` — nested harness, excluded from the library compile
- `HtmlToOpenXml-Guide.md` — the conversion rules of the vendored library

## Shipping

Publishes a GitHub release zip **and** a NuGet package; AgentBridge ships it in every release.
AIOffice does not build it into its dev `Tools/`.
