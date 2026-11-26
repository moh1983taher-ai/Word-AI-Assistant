"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const read = (file) => fs.readFileSync(path.join(root, file), "utf8");

const taskpane = read("src/taskpane/taskpane.js").replace(/\r\n/g, "\n");
const taskpaneHtml = read("src/taskpane/taskpane.html");
const runtimeSource = read("src/taskpane/orama-runtime.js");
const mcpSource = read("src/taskpane/mcp-research-assistant.js");
const pkg = JSON.parse(read("package.json"));
const lock = JSON.parse(read("package-lock.json"));

assert.match(taskpane, /require\(["']\.\/orama-runtime["']\)/);
assert.match(runtimeSource, /require\(["']@orama\/orama["']\)/);
assert.match(runtimeSource, /require\(["']@orama\/plugin-qps["']\)/);
assert.match(runtimeSource, /window\.Orama\s*=\s*\{/);
assert.match(runtimeSource, /pluginQPS/);

assert.doesNotMatch(taskpaneHtml, /cdn\.jsdelivr\.net\/npm\/@orama\//);
assert.doesNotMatch(taskpaneHtml, /<script\s+type=["']module["']/i);
assert.doesNotMatch(taskpaneHtml, /<script\s+src=["']taskpane\.js["']/i);
assert.match(read("webpack.config.js"), /taskpane:\s*\[[\s\S]*?taskpane\.js/);

assert.equal(pkg.dependencies["@orama/plugin-qps"], "^3.1.18");
assert.equal(lock.packages[""].dependencies["@orama/plugin-qps"], "^3.1.18");
assert.equal(lock.packages["node_modules/@orama/plugin-qps"].version, "3.1.18");
assert.equal(lock.packages["node_modules/@orama/plugin-qps"].dependencies["@orama/orama"], "3.1.18");

assert.match(mcpSource, /errorObject\.details/);
assert.match(mcpSource, /JSON\.stringify\(errorObject\.details\)/);

console.log("Orama/QPS Webpack runtime tests passed.");
