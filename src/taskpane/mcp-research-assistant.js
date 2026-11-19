/* global Office */
(function () {
  "use strict";

  const SETTINGS_KEY = "AR_MCP_RESEARCH_SETTINGS";
  const SESSION_SECRETS_KEY = "AR_MCP_RESEARCH_SECRETS";
  const THREAD_KEY = "AR_MCP_RESEARCH_THREAD";
  const BRIDGE_URL = "http://127.0.0.1:18081";
  const DEFAULT_TIMEOUT_MS = 180000;

  const state = {
    enabled: false,
    busy: false,
    messages: [],
    settings: null,
    panel: null,
    turns: null,
    modeButton: null,
    connection: null
  };

  function safeJson(value, fallback) {
    try {
      return value ? JSON.parse(value) : fallback;
    } catch (_) {
      return fallback;
    }
  }

  function loadSettings() {
    const saved = safeJson(localStorage.getItem(SETTINGS_KEY), {});
    const secrets = safeJson(sessionStorage.getItem(SESSION_SECRETS_KEY), {});
    const savedModel = String(saved.model || "").trim();
    return {
      model: /gpt-6\\.2-astra/i.test(savedModel) ? "" : savedModel,
      shamelaToken: String(secrets.shamelaToken || "").trim()
    };
  }

  function saveSettings(settings) {
    const value = { model: String(settings.model || "").trim() };
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(value));
    sessionStorage.setItem(SESSION_SECRETS_KEY, JSON.stringify({
      shamelaToken: String(settings.shamelaToken || "").trim()
    }));
    state.settings = {
      model: value.model,
      shamelaToken: String(settings.shamelaToken || "").trim()
    };
    return state.settings;
  }

  function loadThread() {
    const saved = safeJson(localStorage.getItem(THREAD_KEY), {});
    state.enabled = saved.enabled === true;
    state.messages = Array.isArray(saved.messages)
      ? saved.messages.filter(function (item) {
          return item && (item.role === "user" || item.role === "assistant") && typeof item.text === "string";
        }).slice(-40)
      : [];
  }

  function saveThread() {
    localStorage.setItem(THREAD_KEY, JSON.stringify({
      enabled: state.enabled,
      messages: state.messages.slice(-40)
    }));
  }

  function addStyles() {
    if (document.getElementById("mcp-research-styles")) return;
    const style = document.createElement("style");
    style.id = "mcp-research-styles";
    style.textContent = [
      ".mcp-research-toolbar{display:flex;align-items:center;gap:6px;}",
      ".mcp-research-mode,.mcp-research-settings{border:1px solid #d8dfd8;border-radius:8px;background:#fffdf8;color:#315957;min-height:30px;padding:4px 9px;font:12px 'Segoe UI',Arial,sans-serif;cursor:pointer;}",
      ".mcp-research-mode[aria-pressed='true']{background:#0b5553;color:#fff;border-color:#0b5553;}",
      ".mcp-research-mode:disabled{opacity:.65;cursor:wait;}",
      ".mcp-research-settings{width:30px;padding:3px;font-size:15px;}",
      ".mcp-research-panel{position:absolute;z-index:4500;right:12px;bottom:76px;width:min(370px,calc(100vw - 90px));padding:15px;background:#fffdf8;border:1px solid #d8dfd8;border-radius:12px;box-shadow:0 10px 35px rgba(5,45,43,.16);color:#173e3d;direction:rtl;}",
      ".mcp-research-panel[hidden]{display:none;}",
      ".mcp-research-panel h3{font-size:14px;margin:0 0 10px;font-weight:650;}",
      ".mcp-research-panel label{display:block;font-size:12px;margin:9px 0 4px;color:#405b58;}",
      ".mcp-research-panel input{box-sizing:border-box;width:100%;height:34px;border:1px solid #ccd7d1;border-radius:7px;padding:6px 9px;background:#fff;color:#183c3b;direction:ltr;text-align:left;}",
      ".mcp-research-panel .mcp-note{font-size:11px;line-height:1.7;color:#667772;margin:9px 0;}",
      ".mcp-research-panel .mcp-actions{display:flex;gap:7px;margin-top:12px;flex-wrap:wrap;}",
      ".mcp-research-panel button{border:1px solid #ccd7d1;border-radius:7px;background:#f7f8f3;padding:7px 12px;color:#244b48;cursor:pointer;font-size:12px;}",
      ".mcp-research-panel button.primary{background:#0b5553;color:white;border-color:#0b5553;}",
      ".mcp-research-panel button:disabled{opacity:.55;cursor:wait;}",
      ".mcp-research-health{font-size:11px;line-height:1.6;margin-top:9px;color:#667772;white-space:pre-wrap;}",
      ".mcp-research-transcript{display:flex;flex-direction:column;gap:12px;padding:12px 14px 24px;max-width:900px;margin:0 auto;direction:rtl;}",
      ".mcp-research-message{max-width:92%;border-radius:12px;padding:11px 14px;white-space:pre-wrap;overflow-wrap:anywhere;line-height:1.8;font-size:14px;}",
      ".mcp-research-message.user{align-self:flex-start;background:#edf3ef;color:#183e3b;border:1px solid #dce7e0;}",
      ".mcp-research-message.assistant{align-self:stretch;background:#fffdf8;color:#183e3b;border:1px solid #e7e4d9;}",
      ".mcp-research-message.status{font-size:12px;color:#6c7c75;background:transparent;border:0;padding:2px 4px;}",
      ".mcp-research-message a{color:#08766e;text-decoration:underline;}",
      ".mcp-research-caption{font-size:10px;color:#7b8981;margin-bottom:3px;}",
      ".mcp-research-panel input:focus,.mcp-research-mode:focus-visible,.mcp-research-settings:focus-visible{outline:2px solid #72a49b;outline-offset:1px;}"
    ].join("\n");
    document.head.appendChild(style);
  }

  function createButton(label, id, className, title) {
    const button = document.createElement("button");
    button.type = "button";
    button.id = id;
    button.className = className;
    button.textContent = label;
    button.title = title || label;
    return button;
  }

  function makeField(labelText, id, type, placeholder, value) {
    const wrapper = document.createElement("div");
    const label = document.createElement("label");
    label.htmlFor = id;
    label.textContent = labelText;
    const input = document.createElement("input");
    input.id = id;
    input.type = type;
    input.placeholder = placeholder || "";
    input.value = value || "";
    wrapper.appendChild(label);
    wrapper.appendChild(input);
    return { wrapper: wrapper, input: input };
  }

  function buildSettingsPanel(parent) {
    const panel = document.createElement("section");
    panel.className = "mcp-research-panel";
    panel.id = "mcp-research-settings-panel";
    panel.hidden = true;
    panel.setAttribute("aria-label", "إعدادات مصادر البحث");

    const heading = document.createElement("h3");
    heading.textContent = "إعدادات البحث بالمصادر المتصلة";
    panel.appendChild(heading);

    const settings = state.settings || loadSettings();
    const modelField = makeField(
      "معرّف نموذج DuckAI (اختياري)",
      "mcp-duckai-model",
      "text",
      "اتركه فارغاً لاختيار Luna تلقائياً",
      settings.model
    );
    const shamelaField = makeField(
      "رمز OAuth لوصلة الشاملة (اختياري)",
      "mcp-shamela-token",
      "password",
      "Access token",
      settings.shamelaToken
    );
    panel.appendChild(modelField.wrapper);
    panel.appendChild(shamelaField.wrapper);

    const note = document.createElement("p");
    note.className = "mcp-note";
    note.textContent = "يعتمد هذا المسار على DuckAI المحلي الموجود لديك، ولا يحتاج إلى مفتاح OpenAI أو اشتراك API. يتصل خادم الجامع تلقائياً، أما وصلة الشاملة فتحتاج رمز OAuth. يحفظ الرمز في جلسة Word فقط.";
    panel.appendChild(note);

    const health = document.createElement("div");
    health.id = "mcp-research-health";
    health.className = "mcp-research-health";
    health.setAttribute("role", "status");
    health.textContent = "لم يُختبر الاتصال بعد.";
    panel.appendChild(health);

    const actions = document.createElement("div");
    actions.className = "mcp-actions";
    const save = createButton("حفظ الإعدادات", "mcp-research-save-settings", "primary");
    const test = createButton("اختبار DuckAI المحلي", "mcp-research-test-connection", "");
    const close = createButton("إغلاق", "mcp-research-close-settings", "");
    actions.appendChild(save);
    actions.appendChild(test);
    actions.appendChild(close);
    panel.appendChild(actions);
    parent.appendChild(panel);

    save.addEventListener("click", function () {
      saveSettings({
        model: modelField.input.value,
        shamelaToken: shamelaField.input.value
      });
      panel.hidden = true;
      showNotice("حُفظت إعدادات مصادر البحث على هذا الجهاز.");
    });

    test.addEventListener("click", async function () {
      test.disabled = true;
      health.textContent = "يجري فحص الاتصال بالخادم المحلي…";
      try {
        const result = await checkBridge();
        state.connection = result;
        health.textContent = "الاتصال جاهز · النموذج: " + result.model +
          "\nالنماذج المتاحة: " + (result.models || []).join("، ");
      } catch (error) {
        health.textContent = "تعذر الاتصال: " + (error && error.message ? error.message : String(error));
      } finally {
        test.disabled = false;
      }
    });

    close.addEventListener("click", function () {
      panel.hidden = true;
    });
    state.panel = panel;
  }

  function ensureTurns() {
    const chatArea = document.getElementById("chat-area");
    if (!chatArea) return null;
    let turns = document.getElementById("mcp-research-transcript");
    if (!turns) {
      turns = document.createElement("div");
      turns.id = "mcp-research-transcript";
      turns.className = "mcp-research-transcript";
      chatArea.appendChild(turns);
    }
    state.turns = turns;
    return turns;
  }

  function appendTextWithLinks(container, value) {
    const text = String(value || "");
    const urlPattern = /https?:\/\/[^\s<>"']+/g;
    let cursor = 0;
    let match;
    while ((match = urlPattern.exec(text)) !== null) {
      const rawUrl = match[0];
      const trailing = rawUrl.match(/[)\].،؛,:!?]+$/);
      const url = trailing ? rawUrl.slice(0, -trailing[0].length) : rawUrl;
      if (match.index > cursor) {
        container.appendChild(document.createTextNode(text.slice(cursor, match.index)));
      }
      if (url) {
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.target = "_blank";
        anchor.rel = "noopener noreferrer";
        anchor.textContent = url;
        container.appendChild(anchor);
      }
      if (trailing) container.appendChild(document.createTextNode(trailing[0]));
      cursor = match.index + rawUrl.length;
    }
    if (cursor < text.length) {
      container.appendChild(document.createTextNode(text.slice(cursor)));
    }
  }

  function renderMessage(message, append) {
    const turns = ensureTurns();
    if (!turns) return null;
    const bubble = document.createElement("div");
    bubble.className = "mcp-research-message " + (message.role === "user" ? "user" : "assistant");
    if (message.status) bubble.className += " status";
    if (message.role === "assistant" && !message.status) {
      const caption = document.createElement("div");
      caption.className = "mcp-research-caption";
      caption.textContent = "أدوات الباحث · مصادر MCP عبر DuckAI";
      bubble.appendChild(caption);
    }
    const body = document.createElement("div");
    appendTextWithLinks(body, message.text);
    bubble.appendChild(body);
    if (append === false) turns.prepend(bubble);
    else turns.appendChild(bubble);
    const chatArea = document.getElementById("chat-area");
    if (chatArea) chatArea.scrollTop = chatArea.scrollHeight;
    return bubble;
  }

  function renderThread() {
    const turns = ensureTurns();
    if (!turns) return;
    turns.textContent = "";
    const welcome = document.querySelector("#chat-area .welcome");
    if (state.messages.length && welcome) welcome.style.display = "none";
    else if (welcome) welcome.style.display = "";
    state.messages.forEach(function (message) {
      renderMessage(message, true);
    });
  }

  function showNotice(text) {
    renderMessage({ role: "assistant", text: text }, true);
  }

  function updateModeButton() {
    if (!state.modeButton) return;
    state.modeButton.setAttribute("aria-pressed", state.enabled ? "true" : "false");
    state.modeButton.textContent = state.busy ? "جارٍ البحث…" : (state.enabled ? "بحث MCP مفعّل" : "بحث MCP");
    state.modeButton.title = state.enabled
      ? "البحث عبر DuckAI المحلي وخوادم المكتبات المتصلة؛ اضغط للعودة إلى المسار السابق"
      : "تفعيل البحث بالمصادر عبر DuckAI المحلي";
  }

  async function fetchJson(url, options, timeoutMs) {
    const controller = typeof AbortController !== "undefined" ? new AbortController() : null;
    let timer = null;
    if (controller) timer = window.setTimeout(function () { controller.abort(); }, timeoutMs || 15000);
    try {
      const requestOptions = Object.assign({}, options || {});
      if (controller) requestOptions.signal = controller.signal;
      const response = await fetch(url, requestOptions);
      const payload = await response.json().catch(function () { return {}; });
      if (!response.ok) {
        const detail = payload && (payload.error || payload.message);
        throw new Error(detail ? String(detail) : "HTTP " + response.status);
      }
      return payload;
    } catch (error) {
      if (error && error.name === "AbortError") {
        throw new Error("انتهت مهلة الاتصال بالخادم المحلي.");
      }
      if (error instanceof TypeError) {
        throw new Error("تعذر الوصول إلى DuckAI المحلي. تحقق من تشغيل Word وخدمة DuckAI المحلية.");
      }
      throw error;
    } finally {
      if (timer !== null) window.clearTimeout(timer);
    }
  }

  async function checkBridge() {
    const payload = await fetchJson(BRIDGE_URL + "/health", {
      method: "GET",
      mode: "cors",
      cache: "no-store",
      credentials: "omit"
    }, 12000);
    if (!payload || payload.ready !== true) {
      throw new Error(payload && payload.error ? payload.error : "الخادم المحلي غير جاهز.");
    }
    return {
      model: String(payload.model || "غير محدد"),
      models: Array.isArray(payload.models) ? payload.models.map(String) : []
    };
  }

  async function callLocalResearch(messages, settings) {
    const payload = await fetchJson(BRIDGE_URL + "/v1/research", {
      method: "POST",
      mode: "cors",
      cache: "no-store",
      credentials: "omit",
      headers: {
        "Content-Type": "application/json",
        "X-ResearchTools-Client": "word-addin"
      },
      body: JSON.stringify({
        messages: messages,
        model: settings.model || "",
        shamelaToken: settings.shamelaToken || ""
      })
    }, DEFAULT_TIMEOUT_MS);

    if (!payload || typeof payload.answer !== "string" || !payload.answer.trim()) {
      throw new Error("لم يُرجع محرك DuckAI إجابة نهائية صالحة.");
    }
    return payload;
  }

  function formatToolTrace(payload) {
    const calls = Array.isArray(payload.toolCalls) ? payload.toolCalls : [];
    if (!calls.length) return "";
    const labels = calls.map(function (call) {
      const server = call && call.label ? call.label : (call && call.server ? call.server : "مصدر");
      const tool = call && call.tool ? call.tool : "";
      return tool ? server + " (" + tool + ")" : server;
    });
    return "\n\nالمصادر التي جرى الرجوع إليها: " + labels.join("، ");
  }

  function formatConnectionWarnings(payload) {
    const warnings = Array.isArray(payload.connectionWarnings) ? payload.connectionWarnings : [];
    if (!warnings.length) return "";
    const labels = warnings.map(function (warning) {
      return String(warning).replace(/^aljam3:/, "الجامع:").replace(/^shamela_link:/, "وصلة الشاملة:");
    });
    return "\n\nتنبيه الاتصال بالمصادر: " + labels.join("؛ ");
  }

  async function sendMessage(rawText) {
    const prompt = String(rawText || "").trim();
    if (!prompt || state.busy) return;

    const settings = state.settings || loadSettings();
    const input = document.getElementById("user-input");
    if (input) {
      input.value = "";
      input.dispatchEvent(new Event("input", { bubbles: true }));
    }

    const welcome = document.querySelector("#chat-area .welcome");
    if (welcome) welcome.style.display = "none";

    state.messages.push({ role: "user", text: prompt });
    state.messages = state.messages.slice(-40);
    renderMessage({ role: "user", text: prompt }, true);
    saveThread();

    const status = renderMessage({
      role: "assistant",
      text: "يتصل بـ DuckAI المحلي، ويفحص أدوات المكتبات المتاحة عند الحاجة…",
      status: true
    }, true);

    state.busy = true;
    updateModeButton();

    try {
      const result = await callLocalResearch(state.messages, settings);
      const answer = result.answer.trim() + formatToolTrace(result) + formatConnectionWarnings(result);
      if (status && status.parentNode) status.parentNode.removeChild(status);
      state.messages.push({ role: "assistant", text: answer });
      state.messages = state.messages.slice(-40);
      renderMessage({ role: "assistant", text: answer }, true);
      saveThread();
    } catch (error) {
      if (status && status.parentNode) status.parentNode.removeChild(status);
      const message = error && error.message ? error.message : "حدث خطأ غير متوقع أثناء البحث.";
      state.messages.push({ role: "assistant", text: message });
      state.messages = state.messages.slice(-40);
      renderMessage({ role: "assistant", text: message }, true);
      saveThread();
    } finally {
      state.busy = false;
      updateModeButton();
    }
  }

  function toggleMode() {
    state.enabled = !state.enabled;
    saveThread();
    updateModeButton();
    if (state.enabled) {
      renderThread();
      showNotice("تم تفعيل البحث بالمصادر عبر DuckAI المحلي. لا يحتاج هذا المسار إلى مفتاح OpenAI API؛ والمسار السابق للمحادثة لم يتغير.");
    }
  }

  function resetThread() {
    state.messages = [];
    saveThread();
    if (state.turns) state.turns.remove();
    state.turns = null;
    const welcome = document.querySelector("#chat-area .welcome");
    if (welcome) welcome.style.display = "";
  }

  function initialize() {
    if (document.getElementById("mcp-research-mode")) return;
    const toolbar = document.querySelector(".input-tools");
    const inputArea = document.querySelector(".input-area");
    if (!toolbar || !inputArea || !document.getElementById("send-btn")) return;

    loadThread();
    state.settings = loadSettings();
    addStyles();

    const controls = document.createElement("div");
    controls.className = "mcp-research-toolbar";
    const modeButton = createButton("بحث MCP", "mcp-research-mode", "mcp-research-mode", "تفعيل البحث بالمصادر عبر DuckAI المحلي");
    modeButton.setAttribute("aria-pressed", "false");
    const settingsButton = createButton("⚙", "mcp-research-settings", "mcp-research-settings", "إعدادات مصادر البحث");
    controls.appendChild(modeButton);
    controls.appendChild(settingsButton);
    toolbar.insertBefore(controls, toolbar.firstChild);
    state.modeButton = modeButton;

    buildSettingsPanel(inputArea);
    modeButton.addEventListener("click", toggleMode);
    settingsButton.addEventListener("click", function () {
      const settings = state.settings || loadSettings();
      const modelInput = document.getElementById("mcp-duckai-model");
      const tokenInput = document.getElementById("mcp-shamela-token");
      if (modelInput) modelInput.value = settings.model || "";
      if (tokenInput) tokenInput.value = settings.shamelaToken || "";
      if (state.panel) state.panel.hidden = !state.panel.hidden;
    });

    updateModeButton();
    if (state.messages.length) renderThread();

    window.addEventListener("click", function (event) {
      if (!state.enabled) return;
      const target = event.target && event.target.closest ? event.target.closest("#send-btn") : null;
      if (!target) return;
      event.preventDefault();
      event.stopPropagation();
      if (event.stopImmediatePropagation) event.stopImmediatePropagation();
      const messageInput = document.getElementById("user-input");
      sendMessage(messageInput ? messageInput.value : "");
    }, true);

    window.addEventListener("keydown", function (event) {
      if (!state.enabled || event.key !== "Enter" || event.shiftKey || event.isComposing) return;
      if (!event.target || event.target.id !== "user-input") return;
      event.preventDefault();
      event.stopPropagation();
      if (event.stopImmediatePropagation) event.stopImmediatePropagation();
      sendMessage(event.target.value);
    }, true);

    const newChat = document.getElementById("new-chat-btn");
    if (newChat) {
      newChat.addEventListener("click", function () {
        if (state.enabled) window.setTimeout(resetThread, 0);
      }, true);
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialize, { once: true });
  } else {
    initialize();
  }
})();
