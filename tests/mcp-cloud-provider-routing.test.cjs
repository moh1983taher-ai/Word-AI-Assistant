"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const source = fs.readFileSync(
  path.join(root, "src/taskpane/mcp-research-assistant.js"),
  "utf8"
);

assert.match(source, /const AI_SETTINGS_KEY = ["']AI_SETTINGS["']/);
assert.match(source, /localStorage\.getItem\(AI_SETTINGS_KEY\)/);
assert.match(source, /function loadSettings\(\)[\s\S]*?provider:[\s\S]*?key:[\s\S]*?model:/);
assert.doesNotMatch(source, /AR_MCP_RESEARCH_SETTINGS|AR_MCP_RESEARCH_THREAD.*key/);

assert.match(source, /https:\/\/api\.groq\.com\/openai\/v1\/responses/);
assert.match(source, /type: "mcp"/);
assert.match(source, /server_label: server\.id/);
assert.match(source, /server_url: server\.url/);
assert.match(source, /require_approval: "never"/);
assert.match(source, /allowed_tools: server\.tools\.slice\(\)/);
assert.match(source, /openai\/gpt-oss-20b/);
assert.match(source, /openai\/gpt-oss-120b/);
assert.match(source, /function callGroqResearch[\s\S]*?tools: getGroqTools\(\)/);
assert.match(source, /GROQ_REMOTE_MCP_MODELS\.indexOf\(settings\.model\)/);

assert.match(source, /https:\/\/generativelanguage\.googleapis\.com\/v1beta\/interactions/);
assert.match(source, /antigravity-preview-09-2026/);
assert.match(source, /gemini-4-hydrogen/);
assert.match(source, /type: "mcp_server"/);
assert.match(source, /name: server\.id/);
assert.match(source, /url: server\.url/);
assert.match(
  source,
  /function getGeminiTools\(\)[\s\S]*?allowed_tools: \[\{\s*mode: "auto",\s*tools: server\.tools\.slice\(\)\s*\}\]/
);
assert.match(source, /environment: "remote"/);
assert.match(source, /Antigravity managed agents reject store:false/);
assert.doesNotMatch(source, /\n\s*store\s*:\s*false,/);
assert.match(source, /يتطلب مسار Gemini حفظ التفاعل لدى Google/);
assert.match(source, /function callGeminiResearch[\s\S]*?tools: getGeminiTools\(\)/);

assert.match(source, /https:\/\/aljam3\.com\/mcp/);
assert.match(source, /https:\/\/mcp\.shamela\.ws/);
[
  "search", "list_books", "get_book", "get_page",
  "list_authors", "list_categories", "list_libraries",
  "shamela_find", "shamela_find_scoped", "shamela_open", "shamela_open_many"
].forEach(function (tool) {
  assert.ok(source.includes('"' + tool + '"'), "Read-only MCP tool must be allowlisted: " + tool);
});

assert.match(source, /mcp_server_tool_call/);
assert.match(source, /mcp_server_tool_result/);
assert.match(source, /resultCallId !== callId/);
assert.match(source, /function validateResearchResult[\s\S]*?toolCalls\.length === 0[\s\S]*?أدوات MCP/);
assert.match(source, /async function sendMessage[\s\S]*?callCloudResearch\(state\.messages, settings\)/);
assert.doesNotMatch(source, /127\.0\.0\.1|DuckAI|Duck\.ai|callLocalResearch|BRIDGE_URL/);

console.log("Cloud MCP provider routing tests passed.");
