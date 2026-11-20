"use strict";

// Keep Orama and QPS in the same Webpack bundle as the task pane.
// This avoids runtime CommonJS chunks emitted by third-party CDN ESM transforms.
const {
  create,
  insert,
  insertMultiple,
  search
} = require("@orama/orama");
const { pluginQPS } = require("@orama/plugin-qps");

if (typeof window !== "undefined") {
  window.Orama = {
    create,
    insert,
    insertMultiple,
    search,
    pluginQPS
  };
}
