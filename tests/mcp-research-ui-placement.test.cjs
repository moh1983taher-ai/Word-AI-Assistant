"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const uiSource = fs.readFileSync(
  path.join(root, "src/taskpane/mcp-research-assistant.js"),
  "utf8"
);
const appSource = fs.readFileSync(
  path.join(root, "src/taskpane/taskpane.js"),
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
  "The source-search option must remain inside the existing plus menu."
);
assert.match(uiSource, /mcp-research-mode-option/, "The source switch must retain its stable control ID.");
assert.match(uiSource, /mcp-research-settings-option/, "The provider-status settings must remain reachable.");
assert.doesNotMatch(
  uiSource,
  /document\.querySelector\("\.input-tools"\)|toolbar\.insertBefore\(controls/,
  "MCP controls must not be injected into the message composer toolbar."
);

assert.match(uiSource, /const AI_SETTINGS_KEY = ["']AI_SETTINGS["']/);
assert.match(uiSource, /localStorage\.getItem\(AI_SETTINGS_KEY\)/);
assert.doesNotMatch(
  uiSource,
  /AR_MCP_RESEARCH_SETTINGS|AR_MCP_RESEARCH_THREAD.*key/,
  "MCP search must reuse the existing cloud provider key and settings."
);
assert.match(uiSource, /const AI_SETTINGS_KEY = ["']AI_SETTINGS["']/);
assert.doesNotMatch(
  uiSource,
  /LOCAL_MCP_HOST_URL|127\.0\.0\.1:18081|callLocalResearch|fetchLocalHostJson|checkLocalHost|DuckAI المحلي|mcp-research-test-local|RESEARCH_ENGINE_KEY|loadResearchEngine|saveResearchEngine|mcp-research-engine|طريقة تنفيذ البحث|تلقائي \(المحلي أولاً\)|mcp-research-panel select/,
  "MCP source search must expose only the cloud route."
);
assert.match(uiSource, /بحث سحابي مباشر عبر Groq أو Gemini/);
assert.match(uiSource, /heading\.textContent = "إعدادات البحث"/);
assert.match(uiSource, /sourceTitle\.textContent = "مصادر البحث"/);
assert.match(uiSource, /fontSize = "12px"/);
assert.doesNotMatch(uiSource, /mcp-research-panel select|تلقائي \(المحلي أولاً\)/);
assert.doesNotMatch(uiSource, /المكتبات التي سيبحث فيها النموذج|فحص التوليد الفعلي يتم عند تنفيذ البحث/);


assert.match(
  uiSource,
  /const bridge = window\.WordAIChatBridge[\s\S]*?bridge\.appendUserMessage\(prompt\)[\s\S]*?bridge\.getMessages\(chatId\)/,
  "MCP questions must enter the currently selected ordinary chat and use its conversation history."
);
assert.match(
  uiSource,
  /bridge\.appendAssistantMessage\(chatId, answer,[\s\S]*?sources: sources,[\s\S]*?provider:[\s\S]*?model:/,
  "MCP answers and tool provenance must be saved as assistant messages in the same chat."
);
assert.match(
  uiSource,
  /bridge\.addPendingMessage\(\s*chatId[\s\S]*?bridge\.removePendingMessage\(pending\)/,
  "The pending indicator must be rendered in the native chat without creating another transcript."
);
assert.match(
  uiSource,
  /function toggleMode\(\)[\s\S]*?state\.enabled = !state\.enabled;[\s\S]*?updateModeButton\(\);/,
  "The source switch must only toggle the processing mode."
);

const toggleMode = uiSource.slice(
  uiSource.indexOf("function toggleMode()"),
  uiSource.indexOf("async function initialize()")
);
assert.doesNotMatch(
  toggleMode,
  /renderChat|renderThread|appendUserMessage|appendAssistantMessage|saveThread|localStorage/,
  "Toggling sources must not create, open, clear or persist a separate conversation."
);

assert.match(uiSource, /const LEGACY_THREAD_KEY = ["']AR_MCP_RESEARCH_THREAD["']/);
assert.match(
  uiSource,
  /bridge\.importLegacyThread\(legacyThread\)/,
  "The old source-only transcript must be migrated into the normal chat history."
);
assert.doesNotMatch(
  uiSource,
  /function saveThread\(|state\.messages|renderThread\(|mcp-research-transcript|resetThread\(/,
  "The MCP module must not maintain or render a parallel chat transcript."
);
assert.doesNotMatch(
  uiSource,
  /newChat\.addEventListener/,
  "MCP must not override normal new-chat behavior."
);

assert.match(
  appSource,
  /window\.WordAIChatBridge = Object\.freeze\(/,
  "The main application must expose a single bridge to its native chat model."
);
assert.match(
  appSource,
  /function appendMcpUserMessage\(text\)[\s\S]*?createNewChat\(\)[\s\S]*?finalizeCurrentChat\(prompt\)[\s\S]*?addUserMessage\(currentChat, prompt\)/,
  "A source-search question must use the same create/finalize/save path as an ordinary user message."
);
assert.match(
  appSource,
  /function getMcpChatMessages\(chatId\)[\s\S]*?role: message\.role === "ai" \? "assistant" : "user"/,
  "The model must receive the selected chat's existing messages in provider-neutral roles."
);
assert.match(
  appSource,
  /function appendMcpAssistantMessage\(chatId, text, metadata\)[\s\S]*?role: "ai"[\s\S]*?mcpSources: sources[\s\S]*?saveChat\(chat\)/,
  "MCP answers and source details must be stored in the ordinary chat object."
);
assert.match(
  appSource,
  /function importLegacyMcpThread\(record\)[\s\S]*?mcp-legacy-source-thread[\s\S]*?saveChatsToIndexedDB\(chats\)/,
  "The previously separate source thread must be migrated into the ordinary chat list."
);
assert.match(appSource, /normalizedMessage\.mcpSources/);
assert.match(appSource, /normalizedMessage\.mcpProvider/);
assert.match(appSource, /normalizedMessage\.mcpModel/);
assert.match(
  appSource,
  /window\.dispatchEvent\(new Event\("word-ai-chat-ready"\)\)/,
  "The MCP control must initialize only after the native chat bridge is ready."
);

console.log("Cloud MCP and native chat history tests passed.");
