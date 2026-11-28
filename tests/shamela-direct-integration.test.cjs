"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const adapter = fs.readFileSync(path.join(root, "src/taskpane/shamela-word-bridge.js"), "utf8");
const routing = fs.readFileSync(path.join(root, "src/taskpane/mcp-research-assistant.js"), "utf8");
const app = fs.readFileSync(path.join(root, "src/taskpane/taskpane.js"), "utf8");
const webpack = fs.readFileSync(path.join(root, "webpack.config.js"), "utf8");

assert.match(adapter, /http:\/\/127\.0\.0\.1:18765/);
assert.match(adapter, /credentials:\s*"omit"/);
assert.match(adapter, /getStatus/);
assert.match(adapter, /askShamela/);
assert.match(adapter, /onProgress\(/);
assert.match(adapter, /result\.state === "completed"/);
assert.match(adapter, /result\.state === "failed"/);
assert.doesNotMatch(adapter, /document\.cookie|csrf|Authorization|localStorage\.setItem/i);

assert.match(routing, /id: "shamela_direct"/);
assert.match(routing, /transport: "local"/);
assert.match(routing, /getSelectedMcpServers\(serverIds\)[\s\S]*?server\.transport !== "local"/);
assert.match(routing, /window\.ShamelaWordBridge\.askShamela/);
assert.match(routing, /window\.ShamelaWordBridge\.getStatus/);
assert.match(routing, /bridge\.updatePendingMessage/);
assert.match(routing, /formatShamelaSources/);
assert.match(routing, /formatShamelaSourceAudit/);

assert.match(app, /function updateMcpPendingMessage\(pending, text\)/);
assert.match(app, /updatePendingMessage: updateMcpPendingMessage/);
assert.match(webpack, /\.\/src\/taskpane\/shamela-word-bridge\.js/);

console.log("Shamela direct Word bridge integration tests passed.");
