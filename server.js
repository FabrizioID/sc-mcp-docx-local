"use strict";

const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");
const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} = require("@modelcontextprotocol/sdk/types.js");

function ensureDocxPath(inputPath) {
  if (!inputPath || typeof inputPath !== "string") {
    throw new Error("`path` is required.");
  }

  const resolved = path.resolve(inputPath);
  if (!resolved.toLowerCase().endsWith(".docx")) {
    throw new Error("Only .docx files are supported.");
  }
  if (!fs.existsSync(resolved)) {
    throw new Error(`File not found: ${resolved}`);
  }
  return resolved;
}

function defaultOutputPath(inputPath) {
  const parsed = path.parse(inputPath);
  return path.join(parsed.dir, `${parsed.name}.edited${parsed.ext}`);
}

function runPowerShell(script, args = {}) {
  const payload = { script, args };
  const runner = `
$ErrorActionPreference = 'Stop'
$payload = @'
${JSON.stringify(payload)}
'@
$request = $payload | ConvertFrom-Json -Depth 10
$scriptBlock = [ScriptBlock]::Create($request.script)
& $scriptBlock $request.args
`;
  const encoded = Buffer.from(runner, "utf16le").toString("base64");
  const result = spawnSync(
    "powershell.exe",
    ["-NoProfile", "-NonInteractive", "-EncodedCommand", encoded],
    { encoding: "utf8", maxBuffer: 10 * 1024 * 1024 }
  );

  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    const stderr = (result.stderr || "").trim();
    const stdout = (result.stdout || "").trim();
    throw new Error(stderr || stdout || `PowerShell failed with exit code ${result.status}`);
  }
  const output = (result.stdout || "").trim();
  if (!output) {
    return {};
  }
  try {
    return JSON.parse(output);
  } catch {
    return { raw: output };
  }
}

function readDocx(docxPath) {
  return runPowerShell(
    `
param($args)
$word = $null
$doc = $null
try {
  $resolved = (Resolve-Path -LiteralPath $args.path).Path
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($resolved, $false, $true)
  $result = [PSCustomObject]@{
    path = $resolved
    text = $doc.Content.Text
    paragraphs = $doc.Paragraphs.Count
  }
  $result | ConvertTo-Json -Compress -Depth 10
}
finally {
  if ($doc) { $doc.Close([ref]$false) }
  if ($word) { $word.Quit() }
}
`,
    { path: docxPath }
  );
}

function replaceTextInDocx(docxPath, replacements, outputPath) {
  return runPowerShell(
    `
param($args)
$word = $null
$doc = $null
try {
  $resolved = (Resolve-Path -LiteralPath $args.path).Path
  $target = $args.output_path
  if (-not $target) {
    throw "output_path is required"
  }

  $targetDir = Split-Path -Parent $target
  if ($targetDir -and -not (Test-Path -LiteralPath $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
  }

  Copy-Item -LiteralPath $resolved -Destination $target -Force

  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($target)

  $summary = @()
  foreach ($replacement in $args.replacements) {
    $find = [string]$replacement.find
    $replace = [string]$replacement.replace
    $originalText = $doc.Content.Text
    $matches = ([regex]::Matches($originalText, [regex]::Escape($find))).Count

    $range = $doc.Content
    $range.Find.ClearFormatting() | Out-Null
    $range.Find.Replacement.ClearFormatting() | Out-Null
    $null = $range.Find.Execute(
      $find,
      $false,
      $false,
      $false,
      $false,
      $false,
      $true,
      1,
      $false,
      $replace,
      2
    )

    $summary += [PSCustomObject]@{
      find = $find
      replace = $replace
      matches_before = $matches
    }
  }

  $doc.Save()
  $result = [PSCustomObject]@{
    source_path = $resolved
    output_path = $target
    replacements = $summary
  }
  $result | ConvertTo-Json -Compress -Depth 10
}
finally {
  if ($doc) { $doc.Close() }
  if ($word) { $word.Quit() }
}
`,
    {
      path: docxPath,
      output_path: outputPath,
      replacements,
    }
  );
}

function asTextContent(value) {
  return {
    content: [
      {
        type: "text",
        text: typeof value === "string" ? value : JSON.stringify(value, null, 2),
      },
    ],
  };
}

const server = new Server(
  {
    name: "docx-editor-local",
    version: "0.1.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    {
      name: "read_docx",
      description: "Read plain text content from a local DOCX file.",
      inputSchema: {
        type: "object",
        properties: {
          path: {
            type: "string",
            description: "Path to a local .docx file.",
          },
        },
        required: ["path"],
      },
    },
    {
      name: "replace_text_docx",
      description:
        "Create an edited DOCX copy by replacing text occurrences in a local DOCX file.",
      inputSchema: {
        type: "object",
        properties: {
          path: {
            type: "string",
            description: "Path to the source .docx file.",
          },
          output_path: {
            type: "string",
            description:
              "Optional output path. Defaults to a sibling file ending in .edited.docx.",
          },
          replacements: {
            type: "array",
            description: "List of text replacements to apply.",
            items: {
              type: "object",
              properties: {
                find: { type: "string" },
                replace: { type: "string" },
              },
              required: ["find", "replace"],
            },
            minItems: 1,
          },
        },
        required: ["path", "replacements"],
      },
    },
  ],
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args = {} } = request.params;

  if (name === "read_docx") {
    const docxPath = ensureDocxPath(args.path);
    return asTextContent(readDocx(docxPath));
  }

  if (name === "replace_text_docx") {
    const docxPath = ensureDocxPath(args.path);
    if (!Array.isArray(args.replacements) || args.replacements.length === 0) {
      throw new Error("`replacements` must be a non-empty array.");
    }

    const outputPath = args.output_path
      ? path.resolve(args.output_path)
      : defaultOutputPath(docxPath);

    return asTextContent(
      replaceTextInDocx(
        docxPath,
        args.replacements.map((item) => ({
          find: String(item.find),
          replace: String(item.replace),
        })),
        outputPath
      )
    );
  }

  throw new Error(`Unknown tool: ${name}`);
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error.message}\n`);
  process.exit(1);
});
