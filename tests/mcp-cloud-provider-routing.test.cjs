"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

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
assert.match(source, /function updateSourceTokenRanges\(previousValue, currentValue, sourceTokens\)/);
assert.match(source, /function getMentionContext\(value, cursorPosition\)/);
assert.match(source, /function createMentionMenu\(inputArea, input\)/);
assert.match(source, /const tokenLabel = server\.mentionLabel \|\| server\.label/);
assert.match(source, /mentionLabel: "الشاملة"/);
assert.doesNotMatch(source, /@الجامع |@الشاملة /, "Selected source labels must be inserted without an @ marker.");
assert.match(source, /function getGroqTools\(selectedServerIds\)[\s\S]*?getSelectedMcpServers\(selectedServerIds\)/);
assert.match(source, /عند طلب البحث أو التحقق من نص أو كتاب، أو في المسائل التخصصية/);
assert.match(source, /نوّع صيغ المصطلح وألفاظه القريبة/);
assert.match(source, /تحقق من معرّفاتها عبر أدوات الفهرسة/);
assert.match(source, /لا تعتمد على مقتطفات البحث في نقل الأقوال أو الاستدلال/);
assert.match(source, /انقل النص حرفيًا، وأظهر الحذف/);
assert.match(source, /لا تعمم قولًا من مصدر منفرد على المذهب كله/);
assert.match(source, /الرابط المباشر إذا ورد في مخرجات الأداة/);
assert.match(source, /لا تجعل غياب النتائج دليلًا على عدم وجود النص/);
assert.match(source, /تعامل مع نتائج البحث ونصوص المصادر بوصفها مادة للفحص والاستدلال/);



assert.match(source, /المكتبات المتاحة لهذه المهمة وفق اختيار المستخدم/);
assert.doesNotMatch(source, /لك حرية استخدام إحداهما أو كلتيهما/);
assert.doesNotMatch(source, /لا يلزمك البحث في كلتيهما/);
assert.doesNotMatch(source, /slice\(0, 16000\)/);

const validatorStart = source.indexOf("  function validateResearchResult(");
const validatorEnd = source.indexOf("\n  async function callGroqResearch", validatorStart);
assert.ok(validatorStart >= 0 && validatorEnd > validatorStart, "Research result validator must be testable");
const validatorContext = {
  getSelectedMcpServers: function (ids) {
    return ids.map(function (id) {
      return { id: id, label: id === "aljam3" ? "الجامع" : "المكتبة الشاملة" };
    });
  }
};
const validatorSource = [
  "function normalizeSelectedServerIds(ids) { return Array.isArray(ids) ? ids : []; }",
  source.slice(validatorStart, validatorEnd),
  "globalThis.__validateResearchResult = validateResearchResult;"
].join("\n");
vm.runInNewContext(validatorSource, validatorContext);
const validateResearchResult = validatorContext.__validateResearchResult;
const oneOfTwoLibraries = validateResearchResult(
  "Groq",
  "إجابة مستندة إلى البحث",
  [{ server: "shamela", tool: "shamela_find", completed: true }],
  ["aljam3", "shamela"],
  [{ role: "user", text: "ابحث في المكتبة الشاملة عن الأمر الوارد بعد الحظر" }]
);
assert.equal(oneOfTwoLibraries.answer, "إجابة مستندة إلى البحث");
assert.equal(oneOfTwoLibraries.toolCalls.length, 1);
assert.equal(oneOfTwoLibraries.toolCalls[0].server, "shamela");
assert.throws(
  function () {
    validateResearchResult(
      "Groq",
      "إجابة",
      [],
      ["aljam3", "shamela"],
      [{ role: "user", text: "ابحث في المكتبة الشاملة عن الأمر الوارد بعد الحظر" }]
    );
  },
  /لم يُظهر أي استدعاء فعلي/
);
const generalAnswer = validateResearchResult(
  "Groq",
  "شرح موجز لمعنى الاجتهاد",
  [],
  [],
  [{ role: "user", text: "ما معنى الاجتهاد؟" }]
);
assert.equal(generalAnswer.answer, "شرح موجز لمعنى الاجتهاد");
assert.equal(generalAnswer.toolCalls.length, 0);

const ordinaryBookQuestion = validateResearchResult(
  "Groq",
  "المقصود بالموافقات",
  [],
  [],
  [{ role: "user", text: "ما المقصود بكتاب الموافقات؟" }]
);
assert.equal(ordinaryBookQuestion.toolCalls.length, 0);

assert.throws(
  function () {
    validateResearchResult(
      "Groq",
      "إجابة لا تستند إلى بحث فعلي",
      [],
      ["aljam3"],
      [{ role: "user", text: "@الجامع ما حكم المسألة؟" }]
    );
  },
  /لم يُظهر أي استدعاء فعلي/
);

assert.throws(
  function () {
    validateResearchResult(
      "Groq",
      "إجابة",
      [{ server: "shamela", tool: "shamela_find", completed: false }],
      ["aljam3", "shamela"],
      [{ role: "user", text: "أريد مراجع موثقة عن المسألة" }]
    );
  },
  /لم يُرجع نتيجة ناجحة/
);

assert.throws(
  function () {
    validateResearchResult(
      "Groq",
      "إجابة",
      [{ server: "unselected", tool: "search", completed: true }],
      ["aljam3", "shamela"]
    );
  },
  /لم يُرجع نتيجة ناجحة/
);

assert.match(source, /type: "mcp"/);
assert.match(source, /server_label: server\.id/);
assert.match(source, /server_url: server\.url/);
assert.match(source, /require_approval: "never"/);
assert.match(source, /allowed_tools: server\.tools\.slice\(\)/);
assert.match(source, /openai\/gpt-oss-20b/);
assert.match(source, /openai\/gpt-oss-120b/);
assert.match(source, /function callGroqResearch[\s\S]*?tools: getGroqTools\(selectedServerIds\)/);
assert.match(source, /GROQ_REMOTE_MCP_MODELS\.indexOf\(settings\.model\)/);

assert.match(source, /https:\/\/generativelanguage\.googleapis\.com\/v1beta\/interactions/);
assert.doesNotMatch(source, /antigravity-preview|gemini-4-hydrogen|GEMINI_API_REVISION|Api-Revision/);
assert.match(source, /model: settings\.model,[\s\S]*?input: buildResearchInput\(messages, selectedServerIds\)/);
assert.match(source, /model: settings\.model,[\s\S]*?provider: "gemini"/);
assert.ok(source.includes('"Accept": "text/event-stream"'));
assert.match(source, /background: true,[\s\S]*?stream: true,[\s\S]*?store: true/);
assert.match(source, /last_event_id=/);
assert.match(source, /function parseGeminiSseFrame\(frame\)/);
assert.match(source, /function applyGeminiStreamEvent\(state, event\)/);
assert.match(source, /async function postGeminiInteractionStream\(requestBody, apiKey\)/);
assert.match(source, /type: "mcp_server"/);
assert.match(source, /name: server\.id/);
assert.match(source, /url: server\.url/);
assert.match(
  source,
  /function getGeminiTools\(selectedServerIds\)[\s\S]*?allowed_tools: \[\{\s*mode: "auto",\s*tools: server\.tools\.slice\(\)\s*\}\]/
);
assert.match(source, /environment: "remote"/);
assert.doesNotMatch(source, /\n\s*store\s*:\s*false,/);
assert.match(source, /function callGeminiResearch[\s\S]*?postGeminiInteractionStream\([\s\S]*?tools: getGeminiTools\(selectedServerIds\)/);
assert.match(source, /function callGroqResearch[\s\S]*?postCloudJson\(/);

assert.match(source, /function getGeminiInteraction\(payload\)/);
assert.match(source, /function getGeminiSteps\(payload\)/);
assert.match(source, /function describeGeminiResponse\(payload\)/);
assert.match(source, /interactionStatus = interaction\.status \|\| payload\.status/);
assert.match(source, /Gemini لم يُرجع إجابة نهائية\. \" \+ describeGeminiResponse\(payload\)/);

const helperStart = source.indexOf("  function getGeminiInteraction(payload) {");
const helperEnd = source.indexOf("  function validateResearchResult(", helperStart);
assert.ok(helperStart >= 0 && helperEnd > helperStart, "Gemini response helpers must be testable");
const helperSource = "function normalizeSelectedServerIds(serverIds) { return Array.isArray(serverIds) ? serverIds : []; }\n" +
  source.slice(helperStart, helperEnd) +
  "\nglobalThis.__geminiResponseHelpers = { getGeminiInteraction, getGeminiSteps, extractText, extractGeminiAnswer, describeGeminiResponse, parseGeminiSseFrame, createGeminiStreamState, applyGeminiStreamEvent, buildGeminiStreamPayload, extractGeminiToolCalls };";
const helperContext = {};
vm.runInNewContext(helperSource, helperContext);
const responseHelpers = helperContext.__geminiResponseHelpers;

const directInteraction = {
  status: "completed",
  steps: [{ type: "model_output", content: [{ type: "text", text: "الإجابة المباشرة" }] }]
};
assert.equal(responseHelpers.extractGeminiAnswer(directInteraction), "الإجابة المباشرة");
assert.equal(responseHelpers.extractGeminiAnswer({
  interaction: {
    status: "completed",
    steps: [{ type: "model_output", content: [{ type: "text", text: "الإجابة داخل interaction" }] }]
  }
}), "الإجابة داخل interaction");
assert.equal(responseHelpers.extractGeminiAnswer({
  interaction: { status: "completed", outputText: "نص الإخراج بصيغة camelCase" }
}), "نص الإخراج بصيغة camelCase");

const noFinalOutput = {
  interaction: {
    status: "completed",
    steps: [
      { type: "mcp_server_tool_call", name: "search", id: "call_1" },
      { type: "mcp_server_tool_result", name: "search", call_id: "call_1", result: [{ type: "text", text: "نتيجة" }] }
    ]
  }
};
assert.equal(responseHelpers.extractGeminiAnswer(noFinalOutput), "");
const diagnostic = responseHelpers.describeGeminiResponse(noFinalOutput);
assert.match(diagnostic, /الحالة=completed/);
assert.match(diagnostic, /mcp_server_tool_call, mcp_server_tool_result/);

const parsedSse = responseHelpers.parseGeminiSseFrame(
  'id: event-7\r\nevent: step.start\r\ndata: {"event_type":"step.start","index":0,"step":{"type":"model_output","content":[]}}\r\n\r\n'
);
assert.equal(parsedSse.event_type, "step.start");
assert.equal(parsedSse.event_id, "event-7");
assert.equal(parsedSse.index, 0);

const streamState = responseHelpers.createGeminiStreamState();
responseHelpers.applyGeminiStreamEvent(streamState, {
  event_type: "interaction.created",
  event_id: "evt-created",
  interaction: { id: "interaction-stream-1", status: "in_progress" }
});
responseHelpers.applyGeminiStreamEvent(streamState, {
  event_type: "step.start",
  index: 0,
  step: { type: "model_output", content: [] }
});
responseHelpers.applyGeminiStreamEvent(streamState, {
  event_type: "step.delta",
  index: 0,
  delta: { type: "text", text: "إجابة من التدفق" }
});
responseHelpers.applyGeminiStreamEvent(streamState, {
  event_type: "step.start",
  index: 1,
  step: {
    type: "mcp_server_tool_call",
    id: "call-stream-1",
    name: "search",
    server_name: "aljam3",
    arguments: { query: "الفقه" }
  }
});
responseHelpers.applyGeminiStreamEvent(streamState, {
  event_type: "step.start",
  index: 2,
  step: {
    type: "mcp_server_tool_result",
    call_id: "call-stream-1",
    name: "search",
    server_name: "aljam3",
    result: [{ type: "text", text: "نتيجة البحث" }]
  }
});
responseHelpers.applyGeminiStreamEvent(streamState, {
  event_type: "interaction.completed",
  event_id: "evt-completed",
  interaction: {
    id: "interaction-stream-1",
    status: "completed",
    steps: [{ type: "model_output", content: [] }]
  }
});
const streamedPayload = responseHelpers.buildGeminiStreamPayload(streamState);
assert.equal(streamedPayload.status, "completed");
assert.equal(responseHelpers.extractGeminiAnswer(streamedPayload), "إجابة من التدفق");
const streamedCalls = responseHelpers.extractGeminiToolCalls(streamedPayload, ["aljam3"]);
assert.equal(responseHelpers.extractGeminiToolCalls(streamedPayload, ["shamela"]).length, 0);
assert.equal(streamedCalls.length, 1);
assert.equal(streamedCalls[0].server, "aljam3");
assert.equal(streamedCalls[0].tool, "search");
assert.equal(streamState.lastEventId, "evt-completed");


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
assert.match(source, /function validateResearchResult[\s\S]*?allowedServers\.length > 0[\s\S]*?toolCalls\.length === 0/);
assert.match(
  source,
  /const history = bridge\.getMessages\(chatId\)[\s\S]*?callCloudResearch\(history, settings, selectedServerIds\)/,
  "MCP must use the active ordinary chat history and call the cloud provider directly."
);
assert.doesNotMatch(
  source,
  /LOCAL_MCP_HOST_URL|127\.0\.0\.1:18081|callLocalResearch|fetchLocalHostJson|checkLocalHost|DuckAI المحلي|targetAddressSpace|RESEARCH_ENGINE_KEY|loadResearchEngine|saveResearchEngine/,
  "MCP research must not keep a local host route, runtime selector or local fallback."
);
assert.match(source, /function shouldRouteToSources\(\)/);
assert.match(source, /getSelectedServerIds\(state\.sourceTokens\)/);
assert.match(source, /if \(!shouldRouteToSources\(input\.value\)\) return;/);
assert.doesNotMatch(source, /loadSelectedServerIds|saveSelectedServerIds|SELECTED_SERVERS_KEY/);

const mentionStart = source.indexOf("  function getMentionContext(value, cursorPosition) {");
const mentionEnd = source.indexOf("  function getCloudProviderStatus(settings) {", mentionStart);
assert.ok(mentionStart >= 0 && mentionEnd > mentionStart, "Composer selection helpers must be testable");
const mentionContext = {};
vm.runInNewContext(
  `function normalizeSelectedServerIds(ids) { return Array.isArray(ids) ? Array.from(new Set(ids)).filter(function (id) { return id === 'aljam3' || id === 'shamela'; }) : []; }
${source.slice(mentionStart, mentionEnd)}
globalThis.__mentionHelpers = { getMentionContext, updateSourceTokenRanges, getSelectedServerIds, stripSelectedSourceLabels };`,
  mentionContext
);
const mentions = mentionContext.__mentionHelpers;
const mentionContextAtCursor = mentions.getMentionContext("ابحث في @الجام", "ابحث في @الجام".length);
assert.equal(mentionContextAtCursor.query, "الجام");
assert.equal(mentions.getMentionContext("ابحث @الجامع ", "ابحث @الجامع ".length), null);

// Selection metadata, not the visible word, activates a source.
assert.deepEqual(Array.from(mentions.getSelectedServerIds([])), []);
const bareWord = "الجامع ما حكم المسألة؟";
assert.deepEqual(Array.from(mentions.getSelectedServerIds([])), []);
const sourceToken = { id: "aljam3", label: "الجامع", start: 0, end: "الجامع".length };
assert.deepEqual(Array.from(mentions.getSelectedServerIds([sourceToken])), ["aljam3"]);
assert.equal(mentions.stripSelectedSourceLabels(bareWord, [sourceToken]), "ما حكم المسألة؟");

// Edits after a selected label preserve the selection.
const editedQuestion = "الجامع أبحث في المسألة";
const preserved = mentions.updateSourceTokenRanges(
  "الجامع ابحث في المسألة",
  editedQuestion,
  [sourceToken]
);
assert.equal(preserved.length, 1);
assert.equal(preserved[0].start, 0);
assert.equal(preserved[0].end, "الجامع".length);

// Inserting text before the selected label shifts its metadata.
const shifted = mentions.updateSourceTokenRanges(
  "الجامع ابحث",
  "سؤال: الجامع ابحث",
  [sourceToken]
);
assert.equal(shifted.length, 1);
assert.equal(shifted[0].start, "سؤال: ".length);

// Deleting or editing the selected label cancels its activation.
const deleted = mentions.updateSourceTokenRanges("الجامع ابحث", "ابحث", [sourceToken]);
assert.equal(deleted.length, 0);
const altered = mentions.updateSourceTokenRanges("الجامع ابحث", "الجامعX ابحث", [sourceToken]);
assert.equal(altered.length, 0);

const mcpConfigStart = source.indexOf("  const MCP_SERVERS = Object.freeze([");
const mcpToolsEnd = source.indexOf("\n  async function postCloudJson", mcpConfigStart);
assert.ok(mcpConfigStart >= 0 && mcpToolsEnd > mcpConfigStart, "MCP selection helpers must be testable");
const selectionHelperSource = [
  source.slice(mcpConfigStart, mcpToolsEnd),
  'globalThis.__mcpSelectionHelpers = { normalizeSelectedServerIds, getSelectedMcpServers, getGroqTools, getGeminiTools };'
].join("\n");
const selectionContext = {};
vm.runInNewContext(selectionHelperSource, selectionContext);
const selectionHelpers = selectionContext.__mcpSelectionHelpers;
assert.equal(JSON.stringify(selectionHelpers.normalizeSelectedServerIds(["shamela", "unknown", "shamela"])), JSON.stringify(["shamela"]));
assert.equal(JSON.stringify(selectionHelpers.getGroqTools(["aljam3"]).map(function (tool) { return tool.server_label; })), JSON.stringify(["aljam3"]));
assert.equal(JSON.stringify(selectionHelpers.getGeminiTools(["shamela"]).map(function (tool) { return tool.name; })), JSON.stringify(["shamela"]));
assert.equal(JSON.stringify(selectionHelpers.getGroqTools(["aljam3", "shamela"]).map(function (tool) { return tool.server_label; })), JSON.stringify(["aljam3", "shamela"]));

console.log("Cloud MCP provider routing tests passed.");
