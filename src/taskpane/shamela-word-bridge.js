/* Shamela direct bridge adapter for the Word task pane.
 * Authentication stays in the already-open Shamela Chrome tab.
 * This module never reads or stores cookies, CSRF values, or passwords.
 */
(function attachShamelaWordBridge(root) {
  "use strict";

  if (root.ShamelaWordBridge) return;

  const BASE_URL = "http://127.0.0.1:18765";
  const SIGN_IN_URL = "https://shamela.ws/ai";
  const POLL_INTERVAL_MS = 700;

  async function request(path, options) {
    const response = await fetch(BASE_URL + path, Object.assign({
      mode: "cors",
      credentials: "omit",
      cache: "no-store",
      targetAddressSpace: "loopback"
    }, options || {}));
    const data = await response.json().catch(function () { return {}; });
    if (!response.ok) {
      const error = new Error(data.error || data.message || ("Shamela Bridge HTTP " + response.status));
      error.status = response.status;
      error.data = data;
      throw error;
    }
    return data;
  }

  function sleep(ms) {
    return new Promise(function (resolve) { setTimeout(resolve, ms); });
  }

  async function getStatus() {
    return request("/status", { method: "GET" });
  }

  async function askShamela(query, options) {
    const text = String(query || "").trim();
    if (!text) throw new Error("اكتب السؤال أولًا.");
    if (text.length > 11500) throw new Error("السؤال وسياقه أطول من الحد المسموح للاتصال بالشاملة.");

    const settings = options || {};
    const accepted = await request("/ask", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ query: text, threadId: String(settings.threadId || "").trim() || null })
    });
    if (!accepted.jobId) throw new Error("لم يُرجع الوسيط المحلي معرّف الطلب.");

    let lastSnapshot = "";
    while (true) {
      await sleep(POLL_INTERVAL_MS);
      const result = await request("/result?id=" + encodeURIComponent(accepted.jobId), {
        method: "GET"
      });
      const partialAnswer = typeof result.partialAnswer === "string" ? result.partialAnswer : "";
      const snapshot = [
        result.state || "",
        result.progress || "",
        partialAnswer.length,
        result.eventsReceived || 0,
        Array.isArray(result.sources) ? result.sources.length : 0
      ].join("|");

      if (snapshot !== lastSnapshot) {
        lastSnapshot = snapshot;
        if (typeof settings.onProgress === "function") {
          settings.onProgress({
            state: result.state || "running",
            progress: String(result.progress || "الشاملة تنفذ البحث..."),
            answer: partialAnswer,
            sources: Array.isArray(result.sources) ? result.sources : [],
            eventsReceived: Number(result.eventsReceived || 0)
          });
        }
      }

      if (result.state === "completed") {
        const finalResult = result.result || {};
        return {
          answer: String(finalResult.answer || ""),
          sources: Array.isArray(finalResult.sources) ? finalResult.sources : [],
          sourceAudit: finalResult.sourceAudit || null,
          runId: String(finalResult.runId || ""),
          terminal: String(finalResult.terminal || "run_completed"),
          threadId: String(finalResult.threadId || result.threadId || "").trim() || null
        };
      }

      if (result.state === "failed") {
        const error = new Error(String(result.error || "تعذر إكمال طلب الشاملة."));
        error.partialAnswer = partialAnswer;
        error.sources = Array.isArray(result.sources) ? result.sources : [];
        error.threadId = String(result.threadId || "").trim() || null;
        throw error;
      }
    }
  }

  root.ShamelaWordBridge = Object.freeze({
    baseUrl: BASE_URL,
    signInUrl: SIGN_IN_URL,
    getStatus: getStatus,
    askShamela: askShamela
  });
})(typeof window !== "undefined" ? window : globalThis);
