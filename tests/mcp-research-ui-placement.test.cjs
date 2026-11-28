"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const uiSource = fs.readFileSync(path.join(root, "src/taskpane/mcp-research-assistant.js"), "utf8");
const appSource = fs.readFileSync(path.join(root, "src/taskpane/taskpane.js"), "utf8").replace(/\r\n/g, "\n");
const taskpaneHtml = fs.readFileSync(path.join(root, "src/taskpane/taskpane.html"), "utf8");

assert.match(taskpaneHtml, /id="user-input"/, "The native message composer must remain the interaction point.");
assert.match(taskpaneHtml, /id="send-btn"/, "The native send button must remain unchanged.");
assert.match(uiSource, /function updateSourceTokenRanges\(previousValue, currentValue, sourceTokens\)/);
assert.match(uiSource, /function getMentionContext\(value, cursorPosition\)/);
assert.match(uiSource, /function createMentionMenu\(inputArea, input\)/);
assert.match(uiSource, /function insertSourceMention\(input, server\)/);
assert.match(uiSource, /const tokenLabel = server\.mentionLabel \|\| server\.label/);
assert.doesNotMatch(uiSource, /@الجامع |@الشاملة /);
assert.match(uiSource, /input\.placeholder = "اكتب سؤالك، أو اكتب @ لاختيار مصدر\.\.\."/);
assert.match(uiSource, /if \(!shouldRouteToSources\(\)\) return;/);
assert.match(uiSource, /function shouldRouteToSources\(\)/);
assert.match(uiSource, /getSelectedServerIds\(state\.sourceTokens\)/);

assert.doesNotMatch(
  uiSource,
  /mcp-research-entry|mcp-research-mode-option|mcp-research-settings-option|mcp-research-settings-panel|toggleMode|state\.enabled|updateModeButton|loadSelectedServerIds|saveSelectedServerIds/,
  "There must be no separate source-search button, toggle, or persistent source-selection panel."
);
assert.doesNotMatch(
  uiSource,
  /LOCAL_MCP_HOST_URL|127\.0\.0\.1:18081|callLocalResearch|fetchLocalHostJson|checkLocalHost|DuckAI المحلي|targetAddressSpace/,
  "Source mentions must call only cloud MCP servers."
);
assert.doesNotMatch(
  uiSource,
  /document\.querySelector\("#scope-panel \.scope-panel-content"\)|scopeMenu\.appendChild/,
  "The source picker must not add another item to the plus menu."
);
assert.match(uiSource, /localStorage\.removeItem\("AR_MCP_SELECTED_SERVERS"\)/);
assert.match(uiSource, /const AI_SETTINGS_KEY = ["']AI_SETTINGS["']/);
assert.match(uiSource, /localStorage\.getItem\(AI_SETTINGS_KEY\)/);
assert.doesNotMatch(uiSource, /AR_MCP_RESEARCH_SETTINGS|AR_MCP_RESEARCH_THREAD.*key/);

assert.match(
  uiSource,
  /const chatId = bridge\.appendUserMessage\(visiblePrompt\)[\s\S]*?const history = bridge\.getMessages\(chatId\)[\s\S]*?callCloudResearch\(history, settings, cloudServerIds\)/,
  "A source mention must keep the user's message in the active chat and pass its history to cloud MCP."
);
assert.match(
  uiSource,
  /bridge\.appendAssistantMessage\(chatId, answer,[\s\S]*?sources: allSources,[\s\S]*?provider:[\s\S]*?model:/,
  "MCP answers and source provenance must be saved as assistant messages in the same chat."
);
assert.match(
  uiSource,
  /bridge\.addPendingMessage\([\s\S]*?bridge\.removePendingMessage\(pending\)/,
  "The pending indicator must be rendered in the native chat."
);
assert.match(uiSource, /window\.addEventListener\("click"[\s\S]*?shouldRouteToSources\(\)/);
assert.match(uiSource, /event\.key !== "Enter"[\s\S]*?shouldRouteToSources\(\)/);

assert.match(uiSource, /const LEGACY_THREAD_KEY = ["']AR_MCP_RESEARCH_THREAD["']/);
assert.match(uiSource, /bridge\.importLegacyThread\(legacyThread\)/);
assert.doesNotMatch(
  uiSource,
  /function saveThread\(|state\.messages|renderThread\(|mcp-research-transcript|resetThread\(/,
  "The research feature must not maintain a parallel transcript."
);

assert.match(appSource, /window\.WordAIChatBridge = Object\.freeze\(/);
assert.match(
  appSource,
  /function appendMcpUserMessage\(text\)[\s\S]*?createNewChat\(\)[\s\S]*?finalizeCurrentChat\(prompt\)[\s\S]*?addUserMessage\(currentChat, prompt\)/
);
assert.match(
  appSource,
  /function getMcpChatMessages\(chatId\)[\s\S]*?role: message\.role === "ai" \? "assistant" : "user"/
);
assert.match(
  appSource,
  /function appendMcpAssistantMessage\(chatId, text, metadata\)[\s\S]*?role: "ai"[\s\S]*?mcpSources: sources[\s\S]*?saveChat\(chat\)/
);
assert.match(
  appSource,
  /function importLegacyMcpThread\(record\)[\s\S]*?mcp-legacy-source-thread[\s\S]*?saveChatsToIndexedDB\(chats\)/
);
assert.match(appSource, /normalizedMessage\.mcpSources/);
assert.match(appSource, /normalizedMessage\.mcpProvider/);
assert.match(appSource, /normalizedMessage\.mcpModel/);
assert.match(appSource, /window\.dispatchEvent\(new Event\("word-ai-chat-ready"\)\)/);

console.log("Cloud MCP @-mention routing and native chat history tests passed.");
