#!/usr/bin/env node
/*
 * Regenerates prompts/*.md from assets/prompts.js (the single source of truth
 * for the setup guide). Run from the repository root:
 *
 *   node tools/build-prompts.js
 *
 * No dependencies. Placeholders stay as [BRACKETS] in the Markdown copies;
 * the setup page fills them in from each person's details.
 */
"use strict";
const fs = require("fs");
const path = require("path");
const P = require(path.join(__dirname, "..", "assets", "prompts.js"));

const outDir = path.join(__dirname, "..", "prompts");
fs.mkdirSync(outDir, { recursive: true });

const fence = (text) => {
  // Use a fence longer than any run of backticks in the text.
  const longest = Math.max(2, ...(text.match(/`+/g) || []).map((m) => m.length));
  const f = "`".repeat(longest + 1);
  return f + "text\n" + text + "\n" + f;
};

const index = [
  "# Setup prompts",
  "",
  "Copy-ready prompts for building your own COS Communications Dashboard in Cowork, in order.",
  "The easiest way to use them is the setup guide (`setup.html`), which fills in your details for you.",
  "Here, anything in [BRACKETS] is a placeholder to replace before you paste.",
  "",
  "| Step | Prompt | What it does |",
  "|---|---|---|",
];

for (const p of P.PROMPTS) {
  const text = P.fill(p.text, {});
  const after = p.after ? P.fill(p.after, {}) : "";
  const body = [
    "# " + p.step + ": " + p.title,
    "",
    p.why,
    "",
    "**Where:** " + p.where,
    "",
    fence(text),
    "",
  ];
  if (after) body.push("**After Cowork answers:** " + after, "");
  fs.writeFileSync(path.join(outDir, p.file), body.join("\n"));
  index.push("| " + p.step + " | [" + p.title + "](" + p.file + ") | " + p.why.replace(/\|/g, "\\|") + " |");
}
index.push("", "The automation prompts (Steps 4 to 7) include the exact feed format the dashboard reads; see `docs/feed-format.md`.", "");
fs.writeFileSync(path.join(outDir, "README.md"), index.join("\n"));
console.log("Wrote " + P.PROMPTS.length + " prompts to " + path.relative(process.cwd(), outDir));
