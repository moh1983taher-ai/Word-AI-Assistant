"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const uiSource = fs.readFileSync(
  path.join(root, "src/taskpane/mcp-research-assistant.js"),
  "utf8"
);
const taskpaneHtml = fs.readFileSync(
  path.join(root, "src/taskpane/taskpane.html"),
  "utf8"
);

assert.match(taskpaneHtml, /id="scope-btn"/, "The existing plus trigger must remain present.");
assert.match(taskpaneHtml, /id="scope-panel"/, "The existing plus menu must remain present.");
assert.match(taskpaneHtml, /class="scope-panel-content"/, "The plus menu content container must remain present.");

assert.match(
  uiSource,
  /document\.querySelector\("#scope-panel \.scope-panel-content"\)/,
  "The MCP entry must target the existing plus-menu content."
);
assert.match(
  uiSource,
  /scopeMenu\.appendChild\(entry\)/,
  "The source-search option must remain inside the plus menu."
);
assert.match(
  uiSource,
  /mcp-research-mode-option/,
  "The source-search option must have a stable, distinct control."
);
assert.match(
  uiSource,
  /mcp-research-settings-option/,
  "The provider-status panel must remain reachable from the plus-menu entry."
);
assert.doesNotMatch(
  uiSource,
  /document\.querySelector\("\.input-tools"\)|toolbar\.insertBefore\(controls/,
  "The MCP UI must not inject controls into the message composer toolbar."
);
assert.match(
  uiSource,
  /if \(!state\.enabled\) return;\s*const target = event\.target[\s\S]*?closest\("\#send-btn"\)/,
  "Normal message submission must remain unchanged unless source-search mode is enabled."
);

assert.match(
  uiSource,
  /const AI_SETTINGS_KEY = ["']AI_SETTINGS["']/,
  "Cloud research must reuse the existing AI_SETTINGS credential store."
);
assert.match(
  uiSource,
  /localStorage\.getItem\(AI_SETTINGS_KEY\)/,
  "Cloud research must read the currently saved provider, key and model."
);
assert.doesNotMatch(
  uiSource,
  /AR_MCP_RESEARCH_SETTINGS|BRIDGE_URL|callLocalResearch|127\.0\.0\.1|DuckAI|Duck\.ai/,
  "The source-search feature must not retain a separate key store or local DuckAI route."
);

console.log("Cloud MCP research menu and provider integration tests passed.");
