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
const bridgeSource = fs.readFileSync(
  path.join(root, "runtime/DuckAI-MCP-Bridge.ps1"),
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
  "The MCP entry must be appended inside the plus menu."
);
assert.match(
  uiSource,
  /mcp-research-mode-option/,
  "The source-search option must have a stable, distinct control."
);
assert.match(
  uiSource,
  /mcp-research-settings-option/,
  "Settings must be reachable from the plus-menu entry."
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
  /requestOptions\\.targetAddressSpace = ["']loopback["']/, 
  "Local bridge fetches must explicitly target the loopback address space."
);
assert.match(
  uiSource,
  /["']Content-Type["']:\\s*["']text\\/plain["']/, 
  "Research POST must use a CORS-safelisted content type to avoid an OPTIONS preflight."
);
assert.match(
  uiSource,
  /client:\\s*["']word-addin["']/, 
  "The client marker must travel in the JSON payload, not a custom request header."
);
assert.doesNotMatch(
  uiSource,
  /X-ResearchTools-Client/,
  "Research fetch must not add a non-safelisted custom header."
);
assert.match(
  bridgeSource,
  /function IsAllowedResearchClient[\\s\\S]*?AllowedOrigin[\\s\\S]*?word-addin/,
  "The bridge must validate the exact page origin and client marker."
);
assert.match(
  bridgeSource,
  /IsAllowedResearchClient \\$origin \\$inputData/,
  "The research endpoint must enforce the simple-request origin/client validation."
);

console.log("MCP research menu placement tests passed.");
