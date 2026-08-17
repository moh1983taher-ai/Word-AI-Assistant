/*
 * Word AI Assistant - Full Application
 * Consolidated and Fixed Version
 * 
 * Features:
 * - Projects & Documents Management
 * - Orama-based Search Engine (Arabic support)
 * - AI Integration (OpenRouter, Gemini, OpenAI, Groq)
 * - Streaming Responses with Citations
 * - IndexedDB Storage for Documents
 */

Office.onReady(function () {

    // =====================================================
    // DOM ELEMENTS
    // =====================================================

    const projectsBtn = document.getElementById("projects-btn");
    const projectsPopup = document.getElementById("projects-popup");
    const projectsList = document.getElementById("projects-list");
    const documentsList = document.getElementById("documents-list");
    const addDocumentBtn = document.getElementById("add-document-btn");
    const wordDocumentPicker = document.getElementById("word-document-picker");
    const newProjectBtn = document.getElementById("new-project-btn");
    const chatBtn = document.getElementById("chat-btn");
    const newChatBtn = document.getElementById("new-chat-btn");
    const expandedSidebar = document.getElementById("expanded-sidebar");
    const sidebarToggleBtn = document.getElementById("sidebar-toggle-btn");
    const expandedSidebarToggleSlot = document.getElementById("expanded-sidebar-toggle-slot");
    const input = document.getElementById("user-input");
    const sendBtn = document.getElementById("send-btn");
    const chatArea = document.getElementById("chat-area");
    const documentTitle = document.getElementById("document-title");

    // Settings Elements
    const settingsBtn = document.getElementById("settings-btn");
    const settingsWindow = document.getElementById("settings-window");
    const closeSettings = document.getElementById("close-settings");
    const showKey = document.getElementById("show-key");
    const apiKey = document.getElementById("api-key");
    const provider = document.getElementById("provider-select");
    const modelSelect = document.getElementById("model-select");
    const refreshModels = document.getElementById("refresh-models");
    const saveSettings = document.getElementById("save-settings");
    const testConnection = document.getElementById("test-connection");
    const settingsStatus = document.getElementById("settings-status");
    const providerInfo = document.getElementById("provider-info");

    // Chat Popup
    const chatPopup = document.getElementById("chat-popup");
    const recentChatList = document.getElementById("recent-chat-list");

    // Search Popup
    const searchPopup = document.getElementById("search-popup");
    const searchInput = document.getElementById("search-input");
    const searchResults = document.getElementById("search-results");
    const searchBtn = document.getElementById("search-btn");

    // Sidebar Placeholder
    let sidebarTogglePlaceholder = document.createComment("sidebar-toggle-placeholder");
    if (sidebarToggleBtn && sidebarToggleBtn.parentNode) {
        sidebarToggleBtn.parentNode.insertBefore(sidebarTogglePlaceholder, sidebarToggleBtn);
    }

    // =====================================================
    // STATE
    // =====================================================

    let projects = [];
    let documents = [];
    let chats = [];
    let currentProject = null;
    let currentDocument = null;
    let currentChat = null;
    let currentCitationSources = [];

    // Orama Cache
    let oramaRetrievalDb = null;
    let oramaRetrievalCacheKey = "";
    let oramaRetrievalDocumentId = null;

    // =====================================================
    // CONSTANTS
    // =====================================================

    const DOCUMENT_DB_NAME = "WORD_AI_DOCUMENT_STORAGE";
    const DOCUMENT_DB_VERSION = 4;
    const DOCUMENT_STORE_NAME = "files";
    const DOCUMENT_TEXT_STORE_NAME = "texts";
    const DOCUMENT_STRUCTURE_STORE_NAME = "structures";
    const INDEX_SCHEMA_VERSION = 5;
    const ORAMA_SCHEMA_VERSION = 1;

    // =====================================================
    // ICONS
    // =====================================================

    const projectIcon = `
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#000000" stroke-width="1" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path>
        </svg>
    `;

    const chatIcon = `
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#000000" stroke-width="1" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path>
        </svg>
    `;

    // =====================================================
    // HELPER FUNCTIONS
    // =====================================================

    function readStorageArray(key) {
        try {
            const value = JSON.parse(localStorage.getItem(key) || "[]");
            return Array.isArray(value) ? value : [];
        } catch {
            return [];
        }
    }

    function getSavedSettings() {
        try {
            return JSON.parse(localStorage.getItem("AI_SETTINGS")) || {};
        } catch {
            return {};
        }
    }

    function saveAISettings(data) {
        localStorage.setItem("AI_SETTINGS", JSON.stringify(data));
    }

    function normalizeGeminiModel(model) {
        return String(model || "").replace(/^models\//, "");
    }

    function getAPIError(result, fallback) {
        if (!result) return fallback;
        if (result.error) {
            if (typeof result.error === "string") return result.error;
            if (result.error.message) {
                let message = result.error.message;
                if (result.error.code) message += " | Code: " + result.error.code;
                return message;
            }
        }
        if (result.message) return String(result.message);
        return fallback;
    }

    async function readJSON(response) {
        try {
            return await response.json();
        } catch {
            return {};
        }
    }

    function estimateTokenCount(text) {
        return Math.ceil(String(text || "").length / 4);
    }

    // =====================================================
    // STORAGE FUNCTIONS
    // =====================================================

    function saveProjects() {
        localStorage.setItem("WORD_AI_PROJECTS", JSON.stringify(projects));
    }

    function saveDocuments() {
        localStorage.setItem("WORD_AI_DOCUMENTS", JSON.stringify(documents));
    }

    function saveChats() {
        localStorage.setItem("WORD_AI_CHATS", JSON.stringify(chats));
    }

    // =====================================================
    // LOAD DATA
    // =====================================================

    // Load Projects
    projects = readStorageArray("WORD_AI_PROJECTS")
        .filter(function (item) { return item && typeof item === "object"; })
        .map(function (item) {
            const now = new Date().toISOString();
            return {
                id: item.id || Date.now() + Math.random(),
                name: item.name || "مشروع جديد",
                createdAt: item.createdAt || now,
                updatedAt: item.updatedAt || now,
                documents: Array.isArray(item.documents) ? item.documents : [],
                references: Array.isArray(item.references) ? item.references : [],
                chatIds: Array.isArray(item.chatIds) ? item.chatIds : [],
                settings: item.settings && typeof item.settings === "object" ? item.settings : { citationStyle: "", notes: "" }
            };
        });

    // Load Documents
    documents = readStorageArray("WORD_AI_DOCUMENTS")
        .filter(function (item) { return item && typeof item === "object"; })
        .map(function (item) {
            return {
                ...item,
                indexTokenCount: Number(item.indexTokenCount || 0),
                indexUniqueTerms: Number(item.indexUniqueTerms || 0),
                indexUniqueFamilies: Number(item.indexUniqueFamilies || 0),
                indexSchemaVersion: Number(item.indexSchemaVersion || 0),
                indexStatus: item.indexStatus || "new",
                readStatus: item.readStatus || "new"
            };
        });

    // Load Chats
    chats = readStorageArray("WORD_AI_CHATS")
        .filter(function (item) { return item && typeof item === "object"; })
        .map(function (item) {
            return {
                ...item,
                messages: Array.isArray(item.messages) ? item.messages : []
            };
        });

    saveProjects();
    saveDocuments();
    saveChats();

    // =====================================================
    // INDEXEDDB FUNCTIONS
    // =====================================================

    function openDocumentDatabase() {
        return new Promise(function (resolve, reject) {
            const request = indexedDB.open(DOCUMENT_DB_NAME, DOCUMENT_DB_VERSION);
            request.onupgradeneeded = function () {
                const db = request.result;
                if (!db.objectStoreNames.contains(DOCUMENT_STORE_NAME)) {
                    db.createObjectStore(DOCUMENT_STORE_NAME);
                }
                if (!db.objectStoreNames.contains(DOCUMENT_TEXT_STORE_NAME)) {
                    db.createObjectStore(DOCUMENT_TEXT_STORE_NAME);
                }
                if (!db.objectStoreNames.contains(DOCUMENT_STRUCTURE_STORE_NAME)) {
                    db.createObjectStore(DOCUMENT_STRUCTURE_STORE_NAME);
                }
            };
            request.onsuccess = function () { resolve(request.result); };
            request.onerror = function () { reject(request.error || new Error("فشل فتح قاعدة المستندات.")); };
        });
    }

    // =====================================================
    // WORKING FILE OPERATIONS
    // =====================================================

    async function saveWorkingWordFile(fileId, file) {
        const db = await openDocumentDatabase();
        return new Promise(function (resolve, reject) {
            const tx = db.transaction(DOCUMENT_STORE_NAME, "readwrite");
            const store = tx.objectStore(DOCUMENT_STORE_NAME);
            const request = store.put(file, String(fileId));
            request.onerror = function () { reject(request.error); };
            tx.oncomplete = function () { db.close(); resolve(String(fileId)); };
            tx.onerror = function () { reject(tx.error); };
        });
    }

    async function getWorkingWordFile(fileId) {
        const db = await openDocumentDatabase();
        return new Promise(function (resolve, reject) {
            const tx = db.transaction(DOCUMENT_STORE_NAME, "readonly");
            const request = tx.objectStore(DOCUMENT_STORE_NAME).get(String(fileId));
            request.onsuccess = function () { resolve(request.result || null); };
            request.onerror = function () { reject(request.error); };
            tx.oncomplete = function () { db.close(); };
        });
    }

    async function deleteWorkingWordFile(fileId) {
        const db = await openDocumentDatabase();
        return new Promise(function (resolve, reject) {
            const tx = db.transaction(DOCUMENT_STORE_NAME, "readwrite");
            const request = tx.objectStore(DOCUMENT_STORE_NAME).delete(String(fileId));
            request.onerror = function () { reject(request.error); };
            tx.oncomplete = function () { db.close(); resolve(); };
            tx.onerror = function () { reject(tx.error); };
        });
    }

    // =====================================================
    // DOCUMENT TEXT OPERATIONS
    // =====================================================

    async function saveDocumentText(documentId, text) {
        const db = await openDocumentDatabase();
        return new Promise(function (resolve, reject) {
            const tx = db.transaction(DOCUMENT_TEXT_STORE_NAME, "readwrite");
            const record = {
                documentId: String(documentId),
                text: String(text || ""),
                updatedAt: new Date().toISOString()
            };
            const request = tx.objectStore(DOCUMENT_TEXT_STORE_NAME).put(record, String(documentId));
            request.onerror = function () { reject(request.error); };
            tx.oncomplete = function () { db.close(); resolve(record); };
            tx.onerror = function () { reject(tx.error); };
        });
    }

    async function getDocumentText(documentId) {
        const db = await openDocumentDatabase();
        return new Promise(function (resolve, reject) {
            const tx = db.transaction(DOCUMENT_TEXT_STORE_NAME, "readonly");
            const request = tx.objectStore(DOCUMENT_TEXT_STORE_NAME).get(String(documentId));
            request.onsuccess = function () { resolve(request.result || null); };
            request.onerror = function () { reject(request.error); };
            tx.oncomplete = function () { db.close(); };
        });
    }

    // =====================================================
    // DOCUMENT STRUCTURE OPERATIONS
    // =====================================================

    async function saveDocumentStructure(documentId, structureData) {
        const db = await openDocumentDatabase();
        return new Promise(function (resolve, reject) {
            const tx = db.transaction(DOCUMENT_STRUCTURE_STORE_NAME, "readwrite");
            const request = tx.objectStore(DOCUMENT_STRUCTURE_STORE_NAME).put(structureData, String(documentId));
            request.onerror = function () { reject(request.error); };
            tx.oncomplete = function () { db.close(); resolve(structureData); };
            tx.onerror = function () { reject(tx.error); };
        });
    }

    async function getDocumentStructure(documentId) {
        const db = await openDocumentDatabase();
        return new Promise(function (resolve, reject) {
            const tx = db.transaction(DOCUMENT_STRUCTURE_STORE_NAME, "readonly");
            const request = tx.objectStore(DOCUMENT_STRUCTURE_STORE_NAME).get(String(documentId));
            request.onsuccess = function () { resolve(request.result || null); };
            request.onerror = function () { reject(request.error); };
            tx.oncomplete = function () { db.close(); };
        });
    }

    // =====================================================
    // FILE HELPERS
    // =====================================================

    function fileToBase64(file) {
        return new Promise(function (resolve, reject) {
            const reader = new FileReader();
            reader.onload = function () {
                const result = String(reader.result || "");
                const commaIndex = result.indexOf(",");
                if (commaIndex === -1) {
                    reject(new Error("تعذر تحويل ملف Word إلى Base64."));
                    return;
                }
                resolve(result.substring(commaIndex + 1));
            };
            reader.onerror = function () { reject(reader.error || new Error("فشل قراءة ملف Word.")); };
            reader.readAsDataURL(file);
        });
    }

    // =====================================================
    // TEXT NORMALIZATION
    // =====================================================

    function normalizeSearchText(text) {
        return String(text || "")
            .replace(/[\u064B-\u065F\u0670]/g, "")
            .replace(/\u0640/g, "")
            .replace(/[أإآٱ]/g, "ا")
            .replace(/ى/g, "ي")
            .replace(/ؤ/g, "و")
            .replace(/ئ/g, "ي")
            .replace(/\s+/g, " ")
            .trim()
            .toLowerCase();
    }

    function tokenizeDocumentText(text) {
        return normalizeSearchText(text).match(/[\p{L}\p{N}]+/gu) || [];
    }

    function getSearchQueryTokens(query) {
        const stopWords = new Set([
            "ما", "ماذا", "من", "هو", "هي", "هم", "في", "على", "عن", "الى", "إلى",
            "منه", "بها", "به", "لها", "له", "هذا", "هذه", "ذلك", "تلك", "الذي",
            "التي", "الذين", "بين", "مع", "ثم", "او", "أو", "و", "ف", "ب", "ك", "ل",
            "أن", "إن", "هل", "كيف", "لماذا", "أي", "اي", "كان", "كانت", "يكون", "تكون",
            "فيها", "فيه"
        ]);
        const tokens = tokenizeDocumentText(query);
        const filtered = tokens.filter(function (token) { return token.length > 2 && !stopWords.has(token); });
        return filtered.length ? filtered : tokens;
    }

    // =====================================================
    // CONSERVATIVE FAMILY KEY
    // =====================================================

    function getConservativeFamilyKey(word, surfaceSet) {
        let w = normalizeSearchText(word);
        if (!w) return "";
        w = w.replace(/[^\p{L}\p{N}]+/gu, "");
        if (!w) return "";
        if (w.length <= 3) return w;

        const protectedWords = new Set(["الله", "القران", "اسلام", "اسلامي", "اسلامية"]);
        if (protectedWords.has(w)) return w;

        const prefixes = ["وال", "بال", "كال", "فال", "لل", "ول", "بل", "فل", "ال", "وا", "با", "كا", "فا"];
        const suffixes = ["يات", "ات", "هما", "هم", "هن", "ها", "ية", "يا", "ون", "ين", "ان", "كم", "كن", "ه", "ك", "ي", "ة", "ا"];
        const MIN_ROOT_LENGTH = 3;

        let prefixChanged = true;
        while (prefixChanged) {
            prefixChanged = false;
            for (let i = 0; i < prefixes.length; i++) {
                const prefix = prefixes[i];
                if (w.startsWith(prefix) && (w.length - prefix.length) >= MIN_ROOT_LENGTH) {
                    w = w.substring(prefix.length);
                    prefixChanged = true;
                    break;
                }
            }
        }

        let suffixChanged = true;
        while (suffixChanged) {
            suffixChanged = false;
            for (let i = 0; i < suffixes.length; i++) {
                const suffix = suffixes[i];
                if (w.endsWith(suffix) && (w.length - suffix.length) >= MIN_ROOT_LENGTH) {
                    w = w.substring(0, w.length - suffix.length);
                    suffixChanged = true;
                    break;
                }
            }
        }

        return w.length < MIN_ROOT_LENGTH ? normalizeSearchText(word) : w;
    }

    // =====================================================
    // BUILD DOCUMENT INDEX
    // =====================================================

    function buildDocumentIndex(documentId, text) {
        if (!documentId) throw new Error("معرّف المستند غير موجود.");
        if (typeof text !== "string") throw new Error("نص المستند غير صالح للفهرسة.");

        const tokens = tokenizeDocumentText(text);
        const terms = {};
        const families = {};
        const surfaceSet = new Set(tokens);
        const paragraphTexts = String(text || "").split(/\r\n|\r|\n/);

        let globalTokenPosition = 0;

        paragraphTexts.forEach(function (paragraphText, paragraphIndex) {
            const normalizedParagraphText = normalizeSearchText(paragraphText);
            const tokenMatches = Array.from(normalizedParagraphText.matchAll(/[\p{L}\p{N}]+/gu));

            tokenMatches.forEach(function (match, tokenIndex) {
                const surface = match[0];
                const charStart = typeof match.index === "number" ? match.index : -1;
                const charEnd = charStart === -1 ? -1 : charStart + surface.length;

                if (!surface) return;

                const familyKey = getConservativeFamilyKey(surface, surfaceSet);

                if (!terms[surface]) {
                    terms[surface] = { count: 0, positions: [], occurrences: [], family: familyKey || "" };
                }
                terms[surface].count += 1;
                terms[surface].positions.push(globalTokenPosition);
                terms[surface].occurrences.push({
                    paragraphIndex: paragraphIndex,
                    tokenIndex: tokenIndex,
                    globalIndex: globalTokenPosition,
                    charStart: charStart,
                    charEnd: charEnd
                });

                if (!familyKey) {
                    globalTokenPosition += 1;
                    return;
                }

                if (!families[familyKey]) {
                    families[familyKey] = { count: 0, positions: [], occurrences: [], words: {}, uniqueWords: 0 };
                }

                families[familyKey].count += 1;
                families[familyKey].positions.push(globalTokenPosition);
                families[familyKey].occurrences.push({
                    paragraphIndex: paragraphIndex,
                    tokenIndex: tokenIndex,
                    globalIndex: globalTokenPosition,
                    charStart: charStart,
                    charEnd: charEnd,
                    word: surface
                });

                if (!families[familyKey].words[surface]) {
                    families[familyKey].words[surface] = 0;
                    families[familyKey].uniqueWords += 1;
                }
                families[familyKey].words[surface] += 1;

                globalTokenPosition += 1;
            });
        });

        return {
            documentId: String(documentId),
            indexVersion: INDEX_SCHEMA_VERSION,
            tokenCount: tokens.length,
            uniqueTerms: Object.keys(terms).length,
            uniqueFamilies: Object.keys(families).length,
            terms: terms,
            families: families,
            updatedAt: new Date().toISOString()
        };
    }

    // =====================================================
    // DOCUMENT INDEX STORAGE
    // =====================================================

    async function saveDocumentIndex(documentId, indexData) {
        const db = await openDocumentDatabase();
        return new Promise(function (resolve, reject) {
            const tx = db.transaction("indexes", "readwrite");
            const store = tx.objectStore("indexes");
            const request = store.put(indexData, String(documentId));
            request.onsuccess = function () { resolve(indexData); };
            request.onerror = function () { reject(request.error); };
            tx.oncomplete = function () { db.close(); };
        });
    }

    async function getDocumentIndex(documentId) {
        const db = await openDocumentDatabase();
        return new Promise(function (resolve, reject) {
            const tx = db.transaction("indexes", "readonly");
            const store = tx.objectStore("indexes");
            const request = store.get(String(documentId));
            request.onsuccess = function () { resolve(request.result || null); };
            request.onerror = function () { reject(request.error); };
            tx.oncomplete = function () { db.close(); };
        });
    }

    async function rebuildDocumentIndex(documentId, text) {
        if (!documentId) throw new Error("معرّف المستند غير موجود.");
        if (typeof text !== "string") throw new Error("نص المستند غير صالح للفهرسة.");
        const indexData = buildDocumentIndex(documentId, text);
        await saveDocumentIndex(documentId, indexData);
        return indexData;
    }

    // =====================================================
    // ENSURE DOCUMENT INDEX
    // =====================================================

    async function ensureDocumentIndex(documentItem) {
        if (!documentItem) throw new Error("لم يتم تحديد المستند.");

        let index = await getDocumentIndex(documentItem.id);
        let validIndex = false;

        if (index && index.indexVersion === INDEX_SCHEMA_VERSION && 
            index.terms && typeof index.terms === "object" &&
            index.families && typeof index.families === "object") {
            validIndex = true;
            const termKeys = Object.keys(index.terms);
            for (let i = 0; i < termKeys.length; i++) {
                const term = index.terms[termKeys[i]];
                if (!term || !Array.isArray(term.positions) || !Array.isArray(term.occurrences)) {
                    validIndex = false;
                    break;
                }
            }
            if (validIndex) {
                const familyKeys = Object.keys(index.families);
                for (let i = 0; i < familyKeys.length; i++) {
                    const family = index.families[familyKeys[i]];
                    if (!family || !Array.isArray(family.positions) || 
                        !Array.isArray(family.occurrences) || !family.words || typeof family.words !== "object") {
                        validIndex = false;
                        break;
                    }
                }
            }
        }

        if (validIndex) {
            documentItem.indexStatus = "indexed";
            documentItem.indexTokenCount = index.tokenCount || 0;
            documentItem.indexUniqueTerms = index.uniqueTerms || 0;
            documentItem.indexUniqueFamilies = index.uniqueFamilies || 0;
            documentItem.indexSchemaVersion = index.indexVersion;
            documentItem.indexUpdatedAt = index.updatedAt || "";
            saveDocuments();
            return index;
        }

        const textData = await getDocumentText(documentItem.id);
        if (!textData || typeof textData.text !== "string") return null;

        updateDocumentIndexStatus(documentItem, "indexing");
        try {
            index = await rebuildDocumentIndex(documentItem.id, textData.text);
            await ensureDocumentStructure(documentItem);
            documentItem.indexStatus = "indexed";
            documentItem.indexTokenCount = index.tokenCount || 0;
            documentItem.indexUniqueTerms = index.uniqueTerms || 0;
            documentItem.indexUniqueFamilies = index.uniqueFamilies || 0;
            documentItem.indexSchemaVersion = INDEX_SCHEMA_VERSION;
            documentItem.indexUpdatedAt = index.updatedAt || "";
            saveDocuments();
            return index;
        } catch (error) {
            updateDocumentIndexStatus(documentItem, "error");
            throw error;
        }
    }

    // =====================================================
    // DOCUMENT STRUCTURE BUILDING
    // =====================================================

    async function buildDocumentStructure(documentItem) {
        if (!documentItem) throw new Error("لم يتم تحديد المستند.");
        const file = await getWorkingWordFile(documentItem.storageId);
        if (!file) throw new Error("لم يتم العثور على نسخة العمل.");
        const base64 = await fileToBase64(file);

        if (!Office.context.requirements.isSetSupported("WordApiHiddenDocument", "1.3")) {
            throw new Error("إصدار Word الحالي لا يدعم تحليل بنية المستند.");
        }

        return await Word.run(async function (context) {
            const workingDocument = context.application.createDocument(base64);
            const paragraphs = workingDocument.body.paragraphs;
            const tables = workingDocument.body.tables;

            paragraphs.load(["items/text", "items/styleBuiltIn", "items/tableNestingLevel"]);
            tables.load(["items/rowCount", "items/columnCount", "items/styleBuiltIn"]);

            await context.sync();

            const paragraphItems = paragraphs.items.map(function (paragraph, index) {
                return {
                    index: index,
                    id: String(index),
                    text: String(paragraph.text || "").trim(),
                    style: String(paragraph.styleBuiltIn || ""),
                    tableNestingLevel: Number(paragraph.tableNestingLevel || 0)
                };
            });

            const headings = paragraphItems.filter(function (paragraph) {
                return paragraph.style && /^Heading[1-9]$/i.test(paragraph.style);
            });

            const tableItems = tables.items.map(function (table, index) {
                return {
                    index: index,
                    rows: Number(table.rowCount || 0),
                    columns: Number(table.columnCount || 0),
                    style: String(table.styleBuiltIn || "")
                };
            });

            return {
                documentId: String(documentItem.id),
                paragraphCount: paragraphItems.length,
                headingCount: headings.length,
                tableCount: tableItems.length,
                paragraphs: paragraphItems,
                headings: headings,
                tables: tableItems,
                updatedAt: new Date().toISOString()
            };
        });
    }

    async function ensureDocumentStructure(documentItem) {
        let structure = await getDocumentStructure(documentItem.id);
        if (structure && Array.isArray(structure.paragraphs) && Array.isArray(structure.headings)) {
            return structure;
        }
        structure = await buildDocumentStructure(documentItem);
        await saveDocumentStructure(documentItem.id, structure);
        return structure;
    }

    // =====================================================
    // DOCUMENT STATUS UPDATES
    // =====================================================

    function updateDocumentReadStatus(documentItem, status) {
        if (!documentItem) return;
        documentItem.readStatus = status;
        documentItem.updatedAt = new Date().toISOString();
        if (status === "read") documentItem.readAt = new Date().toISOString();
        saveDocuments();
    }

    function updateDocumentIndexStatus(documentItem, status) {
        if (!documentItem) return;
        documentItem.indexStatus = status;
        documentItem.updatedAt = new Date().toISOString();
        saveDocuments();
    }

    // =====================================================
    // GET PROJECT DOCUMENTS
    // =====================================================

    function getProjectDocuments(projectId) {
        if (!projectId) return [];
        return documents.filter(function (item) { return item && item.projectId === projectId; })
            .sort(function (a, b) { return Number(a.order || 0) - Number(b.order || 0); });
    }

    function attachDocumentToProject(project, documentItem) {
        if (!project || !documentItem) return;
        if (!Array.isArray(project.documents)) project.documents = [];
        if (!project.documents.includes(documentItem.id)) {
            project.documents.push(documentItem.id);
            project.updatedAt = new Date().toISOString();
            saveProjects();
        }
    }

    function setCurrentProject(project) {
        currentProject = project || null;
        renderDocuments();
    }

    function setCurrentDocument(documentItem) {
        if (!documentItem) {
            currentDocument = null;
            if (documentTitle) documentTitle.textContent = "لا يوجد مستند مفتوح";
            renderDocuments();
            return;
        }

        currentDocument = documentItem;
        if (documentTitle) documentTitle.textContent = documentItem.name;

        if (documentItem.readStatus === "read") {
            ensureDocumentIndex(documentItem).then(function () { renderDocuments(); })
                .catch(function (error) { console.error("تعذر تحديث فهرس المستند:", error); renderDocuments(); });
            renderDocuments();
            return;
        }

        readCurrentWordDocument(documentItem).then(function (text) {
            console.log("محتوى نسخة العمل:", text);
            renderDocuments();
        }).catch(function (error) {
            console.error("تعذر قراءة نسخة العمل:", error);
            renderDocuments();
        });
        renderDocuments();
    }

    function createDocument(file, projectId, order) {
        const now = new Date().toISOString();
        const documentId = Date.now();
        const item = {
            id: documentId,
            projectId: projectId,
            name: file.name.replace(/\.docx$/i, ""),
            fileName: file.name,
            fileType: file.type || "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            storageId: String(documentId),
            order: typeof order === "number" ? order : 0,
            type: "word",
            readStatus: "new",
            indexStatus: "new",
            indexTokenCount: 0,
            indexUniqueTerms: 0,
            indexUniqueFamilies: 0,
            indexSchemaVersion: 0,
            createdAt: now,
            updatedAt: now
        };
        documents.push(item);
        saveDocuments();
        return item;
    }

    // =====================================================
    // READ CURRENT WORD DOCUMENT
    // =====================================================

    async function readCurrentWordDocument(documentItem) {
        if (!documentItem) throw new Error("لم يتم تحديد المستند.");
        updateDocumentReadStatus(documentItem, "reading");

        try {
            if (!Office.context.requirements.isSetSupported("WordApiHiddenDocument", "1.3")) {
                throw new Error("إصدار Word الحالي لا يدعم قراءة نسخة العمل.");
            }

            const file = await getWorkingWordFile(documentItem.storageId);
            if (!file) throw new Error("لم يتم العثور على نسخة العمل.");
            const base64 = await fileToBase64(file);

            const text = await Word.run(async function (context) {
                const workingDocument = context.application.createDocument(base64);
                const body = workingDocument.body;
                body.load("text");
                await context.sync();
                return body.text || "";
            });

            await saveDocumentText(documentItem.id, text);
            updateDocumentReadStatus(documentItem, "read");

            updateDocumentIndexStatus(documentItem, "indexing");
            const indexData = await rebuildDocumentIndex(documentItem.id, text);

            const structureData = await buildDocumentStructure(documentItem);
            await saveDocumentStructure(documentItem.id, structureData);

            documentItem.indexTokenCount = indexData.tokenCount || 0;
            documentItem.indexUniqueTerms = indexData.uniqueTerms || 0;
            documentItem.indexUniqueFamilies = indexData.uniqueFamilies || 0;
            documentItem.indexSchemaVersion = INDEX_SCHEMA_VERSION;
            documentItem.indexUpdatedAt = indexData.updatedAt || new Date().toISOString();

            updateDocumentIndexStatus(documentItem, "indexed");

            console.log("تمت قراءة المستند وفهرسته بنجاح:", {
                documentId: documentItem.id,
                tokenCount: indexData.tokenCount,
                uniqueTerms: indexData.uniqueTerms,
                uniqueFamilies: indexData.uniqueFamilies
            });

            return text;
        } catch (error) {
            updateDocumentReadStatus(documentItem, "error");
            updateDocumentIndexStatus(documentItem, "error");
            console.error("فشل قراءة/فهرسة المستند:", error);
            throw error;
        }
    }

    // =====================================================
    // HEADING LEVEL
    // =====================================================

    function getHeadingLevelNumber(style) {
        const match = String(style || "").match(/Heading\s*([1-9])/i);
        return match ? Number(match[1]) : 9;
    }

    // =====================================================
    // ORAMA LOADER
    // =====================================================

    function loadOrama() {
        try {
            return require("@orama/orama");
        } catch (error) {
            throw new Error("تعذر تحميل Orama من الحزمة @orama/orama: " + error.message);
        }
    }

    // =====================================================
    // BUILD ORAMA RETRIEVAL INDEX
    // =====================================================

    async function buildOramaRetrievalIndex(documentItem, structureData) {
        const { create, insertMultiple } = loadOrama();

        const paragraphs = Array.isArray(structureData.paragraphs) ? structureData.paragraphs : [];
        const headings = Array.isArray(structureData.headings) ? structureData.headings
            .filter(function (heading) {
                return heading && typeof heading.index !== "undefined" && String(heading.text || "").trim();
            })
            .sort(function (a, b) { return Number(a.index) - Number(b.index); })
            : [];

        const headingMap = new Map();
        headings.forEach(function (heading) { headingMap.set(Number(heading.index), heading); });

        function nearestHeading(paragraphIndex) {
            if (headingMap.has(paragraphIndex)) return headingMap.get(paragraphIndex);
            for (let i = headings.length - 1; i >= 0; i--) {
                if (Number(headings[i].index) < paragraphIndex) return headings[i];
            }
            return null;
        }

        const records = [];
        paragraphs.forEach(function (paragraph) {
            if (!paragraph || !paragraph.text) return;
            const text = String(paragraph.text).trim();
            if (!text) return;
            const paragraphIndex = Number(paragraph.index);
            const heading = nearestHeading(paragraphIndex);
            const isHeading = Boolean(heading && Number(heading.index) === paragraphIndex);

            records.push({
                id: "p-" + String(paragraphIndex),
                paragraphIndex: paragraphIndex,
                text: text,
                heading: heading ? String(heading.text || "").trim() : "",
                headingIndex: heading ? Number(heading.index) : -1,
                headingLevel: heading ? String(heading.style || "") : "",
                headingLevelNumber: heading ? getHeadingLevelNumber(heading.style) : 9,
                isHeading: isHeading,
                documentId: String(documentItem.id)
            });
        });

        const cacheKey = [
            String(documentItem.id),
            String(documentItem.indexUpdatedAt || ""),
            String(ORAMA_SCHEMA_VERSION),
            String(records.length),
            String(headings.length)
        ].join("|");

        if (oramaRetrievalDb && oramaRetrievalCacheKey === cacheKey && 
            oramaRetrievalDocumentId === String(documentItem.id)) {
            return oramaRetrievalDb;
        }

        const db = create({
            schema: {
                id: "string",
                paragraphIndex: "number",
                text: "string",
                heading: "string",
                headingIndex: "number",
                headingLevel: "string",
                headingLevelNumber: "number",
                isHeading: "boolean",
                documentId: "string"
            },
            language: "arabic"
        });

        if (records.length) {
            await insertMultiple(db, records, 500);
        }

        oramaRetrievalDb = db;
        oramaRetrievalCacheKey = cacheKey;
        oramaRetrievalDocumentId = String(documentItem.id);

        documentItem.indexStatus = "indexed";
        documentItem.indexTokenCount = records.reduce(function (total, record) {
            return total + tokenizeDocumentText(record.text).length;
        }, 0);
        documentItem.indexUniqueTerms = new Set(records.flatMap(function (record) {
            return tokenizeDocumentText(record.text);
        })).size;
        documentItem.indexUniqueFamilies = 0;
        documentItem.indexSchemaVersion = ORAMA_SCHEMA_VERSION;
        documentItem.indexUpdatedAt = new Date().toISOString();
        saveDocuments();

        console.log("تم بناء فهرس Orama الرئيسي:", {
            documentId: documentItem.id,
            records: records.length,
            headings: headings.length
        });

        return db;
    }

    // =====================================================
    // SEARCH INDEXED DOCUMENT
    // =====================================================

    async function searchIndexedDocument(documentId, query, options) {
        const searchTerm = normalizeSearchText(query);
        if (!searchTerm) {
            return { query: "", count: 0, results: [] };
        }

        const documentItem = documents.find(function (doc) {
            return doc && String(doc.id) === String(documentId);
        }) || currentDocument;

        if (!documentItem) throw new Error("لم يتم العثور على المستند.");

        const indexData = await ensureDocumentIndex(documentItem);
        if (!indexData) throw new Error("لا يوجد فهرس صالح لهذا المستند.");

        const indexedTerms = indexData.terms || {};
        const indexedFamilies = indexData.families || {};
        const queryTokens = getSearchQueryTokens(searchTerm);
        if (queryTokens.length === 0) {
            return { query: searchTerm, count: 0, results: [] };
        }

        const indexedSurfaceSet = new Set(Object.keys(indexedTerms));
        const queryFamilyKeys = [];
        queryTokens.forEach(function (token) {
            const familyKey = getConservativeFamilyKey(token, indexedSurfaceSet);
            if (familyKey && !queryFamilyKeys.includes(familyKey)) {
                queryFamilyKeys.push(familyKey);
            }
        });

        const matchedFamilies = queryFamilyKeys.filter(function (familyKey) {
            return Boolean(indexedFamilies[familyKey]);
        });

        const matchedExactTerms = queryTokens.filter(function (token) {
            return Boolean(indexedTerms[token]);
        });

        const structureData = await ensureDocumentStructure(documentItem);
        if (!structureData) throw new Error("لا توجد بنية محفوظة لهذا المستند.");

        const paragraphs = Array.isArray(structureData.paragraphs) ? structureData.paragraphs : [];
        const headings = Array.isArray(structureData.headings) ? structureData.headings : [];
        const paragraphMap = new Map();
        paragraphs.forEach(function (paragraph) {
            if (paragraph && typeof paragraph.index !== "undefined") {
                paragraphMap.set(paragraph.index, paragraph);
            }
        });

        const candidateParagraphs = new Map();

        function addCandidateParagraph(paragraphIndex, key, occurrence) {
            if (typeof paragraphIndex === "undefined") return;
            if (!candidateParagraphs.has(paragraphIndex)) {
                candidateParagraphs.set(paragraphIndex, new Map());
            }
            const entry = candidateParagraphs.get(paragraphIndex);
            if (!entry.has(key)) entry.set(key, []);
            if (occurrence) entry.get(key).push(occurrence);
        }

        matchedFamilies.forEach(function (familyKey) {
            const family = indexedFamilies[familyKey];
            if (!family) return;
            if (Array.isArray(family.occurrences)) {
                family.occurrences.forEach(function (occurrence) {
                    if (occurrence) addCandidateParagraph(occurrence.paragraphIndex, familyKey, occurrence);
                });
            }
        });

        matchedExactTerms.forEach(function (term) {
            const termData = indexedTerms[term];
            if (!termData) return;
            if (Array.isArray(termData.occurrences)) {
                termData.occurrences.forEach(function (occurrence) {
                    if (occurrence) addCandidateParagraph(occurrence.paragraphIndex, term, occurrence);
                });
            }
        });

        // Heading-based search
        const meaningfulQueryTokens = queryTokens.filter(function (token) { return token && token.length >= 3; });
        const headingMatchedParagraphs = new Map();

        headings.forEach(function (heading) {
            if (!heading || typeof heading.index === "undefined" || !heading.text) return;
            const headingText = normalizeSearchText(heading.text);
            if (!headingText) return;

            let matchedHeadingTokens = 0;
            meaningfulQueryTokens.forEach(function (token) {
                if (headingText.includes(token)) matchedHeadingTokens += 1;
            });
            const headingCoverage = meaningfulQueryTokens.length > 0 ? matchedHeadingTokens / meaningfulQueryTokens.length : 0;

            const headingExact = headingText === searchTerm;
            const headingContainsQuery = headingText.includes(searchTerm);
            const isStrongHeadingMatch = headingExact || headingContainsQuery || 
                (matchedHeadingTokens >= 2 && headingCoverage >= 0.40);

            if (!isStrongHeadingMatch) return;

            let headingScore = 0;
            if (headingExact) headingScore += 40;
            if (headingContainsQuery) headingScore += 25;
            headingScore += matchedHeadingTokens * 6;
            headingScore += headingCoverage * 20;

            headingMatchedParagraphs.set(heading.index, {
                heading: heading,
                score: headingScore,
                matchedHeadingTokens: matchedHeadingTokens,
                headingCoverage: headingCoverage
            });

            addCandidateParagraph(heading.index, "__heading__", {
                charStart: 0,
                charEnd: headingText.length,
                tokenIndex: 0,
                headingMatch: true
            });

            let addedFollowing = 0;
            for (let i = heading.index + 1; i < paragraphs.length && addedFollowing < 3; i++) {
                const followingParagraph = paragraphs[i];
                if (!followingParagraph) continue;
                const followingText = String(followingParagraph.text || "").trim();
                if (!followingText) continue;
                if (followingParagraph.style && /^Heading[1-9]$/i.test(followingParagraph.style)) break;
                addCandidateParagraph(followingParagraph.index, "__heading__", {
                    headingMatch: true,
                    sourceHeadingIndex: heading.index
                });
                addedFollowing += 1;
            }
        });

        if (candidateParagraphs.size === 0) {
            return {
                query: searchTerm,
                count: 0,
                results: [],
                matchedTerms: matchedExactTerms,
                matchedFamilies: matchedFamilies,
                totalQueryTerms: queryTokens.length,
                indexTokenCount: indexData.tokenCount,
                indexUniqueTerms: indexData.uniqueTerms,
                indexUniqueFamilies: indexData.uniqueFamilies || 0,
                indexedOccurrences: 0
            };
        }

        const results = [];
        candidateParagraphs.forEach(function (matchedEntries, paragraphIndex) {
            const paragraph = paragraphMap.get(paragraphIndex);
            if (!paragraph || !paragraph.text) return;

            const originalText = String(paragraph.text);
            const normalizedText = normalizeSearchText(originalText);

            let nearestHeading = null;
            for (let i = headings.length - 1; i >= 0; i--) {
                const heading = headings[i];
                if (heading && heading.index < paragraphIndex) {
                    nearestHeading = heading;
                    break;
                }
            }

            let headingMatch = matchedEntries.has("__heading__");
            let headingScore = 0;
            let matchedHeadingTokens = 0;
            let headingCoverage = 0;

            if (nearestHeading) {
                const headingData = headingMatchedParagraphs.get(nearestHeading.index);
                if (headingData) {
                    headingScore = headingData.score || 0;
                    matchedHeadingTokens = headingData.matchedHeadingTokens || 0;
                    headingCoverage = headingData.headingCoverage || 0;
                }
            }

            let matchedFamilyCount = 0;
            matchedFamilies.forEach(function (familyKey) {
                if (matchedEntries.has(familyKey)) matchedFamilyCount += 1;
            });

            let matchedFamilyOccurrences = 0;
            matchedFamilies.forEach(function (familyKey) {
                const occurrences = matchedEntries.get(familyKey);
                if (Array.isArray(occurrences)) matchedFamilyOccurrences += occurrences.length;
            });

            let exactWordMatches = 0;
            matchedExactTerms.forEach(function (term) {
                if (matchedEntries.has(term)) exactWordMatches += 1;
            });

            let queryCoverage = 0;
            if (matchedFamilies.length > 0) {
                queryCoverage = matchedFamilyCount / matchedFamilies.length;
            }

            let familyProximityScore = 0;
            let familySpan = null;
            if (matchedFamilies.length > 1) {
                const events = [];
                matchedFamilies.forEach(function (familyKey) {
                    const occurrences = matchedEntries.get(familyKey);
                    if (!Array.isArray(occurrences)) return;
                    occurrences.forEach(function (occurrence) {
                        if (occurrence && typeof occurrence.tokenIndex === "number") {
                            events.push({ position: occurrence.tokenIndex, family: familyKey });
                        }
                    });
                });
                events.sort(function (a, b) { return a.position - b.position; });

                if (events.length) {
                    let left = 0;
                    const familyCounts = new Map();
                    let familiesInside = 0;
                    for (let right = 0; right < events.length; right++) {
                        const rightFamily = events[right].family;
                        const oldCount = familyCounts.get(rightFamily) || 0;
                        familyCounts.set(rightFamily, oldCount + 1);
                        if (oldCount === 0) familiesInside += 1;

                        while (familiesInside === matchedFamilies.length && left <= right) {
                            const currentSpan = events[right].position - events[left].position + 1;
                            if (familySpan === null || currentSpan < familySpan) {
                                familySpan = currentSpan;
                            }
                            const leftFamily = events[left].family;
                            const leftCount = familyCounts.get(leftFamily);
                            if (leftCount === 1) {
                                familyCounts.delete(leftFamily);
                                familiesInside -= 1;
                            } else {
                                familyCounts.set(leftFamily, leftCount - 1);
                            }
                            left += 1;
                        }
                    }
                }
                if (familySpan !== null) {
                    familyProximityScore = 8 / familySpan;
                }
            }

            let familyOccurrencesInParagraph = 0;
            matchedEntries.forEach(function (occurrences, key) {
                if (matchedFamilies.includes(key)) {
                    familyOccurrencesInParagraph += occurrences.length;
                }
            });

            let score = 0;
            if (normalizedText.includes(searchTerm)) score += 12;
            score += exactWordMatches * 4;
            if (matchedFamilies.length > 0) score += queryCoverage * 8;
            score += familyProximityScore;
            score += Math.min(matchedFamilyOccurrences, 8) * 0.75;
            score += Math.min(familyOccurrencesInParagraph, 6) * 0.75;
            if (normalizedText.startsWith(searchTerm)) score += 2;
            if (headingMatch) score += 20;
            if (headingScore > 0) score += headingScore;

            if (nearestHeading) {
                const normalizedHeading = normalizeSearchText(nearestHeading.text);
                if (normalizedHeading === searchTerm) score += 30;
                else if (normalizedHeading.includes(searchTerm)) score += 18;
                if (matchedHeadingTokens >= 2) score += matchedHeadingTokens * 3;
                score += headingCoverage * 10;
            }

            let searchPosition = -1;
            let matchedOccurrence = null;
            matchedEntries.forEach(function (occurrences, key) {
                if (matchedOccurrence) return;
                if (key === "__heading__") return;
                if (!matchedFamilies.includes(key) && !matchedExactTerms.includes(key)) return;
                if (!Array.isArray(occurrences) || occurrences.length === 0) return;
                const occurrence = occurrences[0];
                if (occurrence && typeof occurrence.charStart === "number" && occurrence.charStart >= 0) {
                    matchedOccurrence = occurrence;
                    searchPosition = occurrence.charStart;
                }
            });

            if (searchPosition === -1) {
                searchPosition = normalizedText.indexOf(searchTerm);
            }

            let context = originalText;
            if (matchedOccurrence && typeof matchedOccurrence.charStart === "number") {
                const charStart = matchedOccurrence.charStart;
                const charEnd = typeof matchedOccurrence.charEnd === "number" ? matchedOccurrence.charEnd : 
                    charStart + (matchedOccurrence.word ? matchedOccurrence.word.length : searchTerm.length);
                const contextStart = Math.max(0, charStart - 120);
                const contextEnd = Math.min(normalizedText.length, charEnd + 300);
                context = normalizedText.substring(contextStart, contextEnd);
            } else if (searchPosition !== -1) {
                const start = Math.max(0, searchPosition - 120);
                const end = Math.min(normalizedText.length, searchPosition + searchTerm.length + 300);
                context = normalizedText.substring(start, end);
            } else {
                context = originalText.substring(0, 420);
            }

            let matchType = "family";
            if (headingMatch) matchType = "heading";
            else if (normalizedText.includes(searchTerm)) matchType = "exact";
            else if (exactWordMatches > 0) matchType = "word";

            results.push({
                paragraphIndex: paragraphIndex,
                paragraphId: paragraph.id,
                text: originalText,
                context: context,
                matchedOccurrence: matchedOccurrence,
                heading: nearestHeading ? nearestHeading.text : "",
                headingLevel: nearestHeading ? nearestHeading.style : "",
                matchedTerms: matchedExactTerms,
                matchedFamilies: matchedFamilies,
                matchedFamilyCount: matchedFamilyCount,
                familyOccurrencesInParagraph: familyOccurrencesInParagraph,
                exactWordMatches: exactWordMatches,
                totalQueryTerms: queryTokens.length,
                headingMatch: headingMatch,
                headingScore: headingScore,
                headingCoverage: headingCoverage,
                score: score,
                matchType: matchType
            });
        });

        results.sort(function (a, b) {
            if (b.score !== a.score) return b.score - a.score;
            if (Boolean(b.headingMatch) !== Boolean(a.headingMatch)) return b.headingMatch ? 1 : -1;
            if (b.matchedFamilyCount !== a.matchedFamilyCount) return b.matchedFamilyCount - a.matchedFamilyCount;
            if (b.exactWordMatches !== a.exactWordMatches) return b.exactWordMatches - a.exactWordMatches;
            return a.paragraphIndex - b.paragraphIndex;
        });

        let indexedOccurrences = 0;
        matchedFamilies.forEach(function (familyKey) {
            const family = indexedFamilies[familyKey];
            if (family) indexedOccurrences += Number(family.count || 0);
        });

        return {
            query: searchTerm,
            count: results.length,
            results: results,
            matchedTerms: matchedExactTerms,
            matchedFamilies: matchedFamilies,
            totalQueryTerms: queryTokens.length,
            indexTokenCount: indexData.tokenCount,
            indexUniqueTerms: indexData.uniqueTerms,
            indexUniqueFamilies: indexData.uniqueFamilies || 0,
            indexedOccurrences: indexedOccurrences
        };
    }

    // =====================================================
    // BUILD AI DOCUMENT CONTEXT
    // =====================================================

    function getRetrievalProfile(query) {
        const text = normalizeSearchText(query);
        const profile = { type: "general", maxResults: 8, maxChars: 8000 };

        if (/ماهو|ماهو|ماهى|ماهي|ما هي|المقصود|معنى|تعريف|يقصد ب|المراد ب/.test(text)) {
            profile.type = "definition";
            profile.maxResults = 5;
            profile.maxChars = 6000;
            return profile;
        }
        if (/اثر|أثر|تاثير|تأثير|نتائج|ينتج عن|يترتب على|انعكاس/.test(text)) {
            profile.type = "effect";
            profile.maxResults = 8;
            profile.maxChars = 9000;
            return profile;
        }
        if (/الفرق|الفروق|مقارنة|يقارن|ما الفرق|التمييز بين|يفترق/.test(text)) {
            profile.type = "comparison";
            profile.maxResults = 10;
            profile.maxChars = 10000;
            return profile;
        }
        if (/لماذا|سبب|اسباب|أسباب|علة|علل|لأن|لان|بسبب/.test(text)) {
            profile.type = "causes";
            profile.maxResults = 8;
            profile.maxChars = 9000;
            return profile;
        }
        if (/اين|أين|موضع|موضعه|الفصل|المبحث|المطلب|الصفحة/.test(text)) {
            profile.type = "location";
            profile.maxResults = 6;
            profile.maxChars = 6000;
            return profile;
        }
        return profile;
    }

    function getRetrievalLimits(providerName, modelName) {
        const providerValue = String(providerName || "").toLowerCase();
        if (providerValue === "groq") return { maxResults: 4, maxChars: 3500 };
        if (providerValue === "openrouter") return { maxResults: 5, maxChars: 5000 };
        if (providerValue === "gemini") return { maxResults: 6, maxChars: 6000 };
        if (providerValue === "openai") return { maxResults: 6, maxChars: 6000 };
        return { maxResults: 4, maxChars: 3500 };
    }

    function getCommonTextLength(textA, textB) {
        const a = String(textA || "");
        const b = String(textB || "");
        if (!a || !b) return 0;
        let best = 0;
        const minLength = Math.min(a.length, b.length);
        const maxWindow = Math.min(minLength, 300);
        for (let length = maxWindow; length >= 20; length -= 10) {
            let found = false;
            for (let i = 0; i + length <= a.length; i += 10) {
                const part = a.substring(i, i + length);
                if (b.includes(part)) {
                    best = length;
                    found = true;
                    break;
                }
            }
            if (found) break;
        }
        return best;
    }

    function buildRetrievalContext(searchResult, options) {
        const settings = options || {};
        const maxResults = typeof settings.maxResults === "number" ? settings.maxResults : 4;
        const maxChars = typeof settings.maxChars === "number" ? settings.maxChars : 3500;
        const includeNeighbors = settings.includeNeighbors !== false;

        if (!searchResult || !Array.isArray(searchResult.results)) {
            return { query: searchResult && searchResult.query ? searchResult.query : "", count: 0, selectedCount: 0, totalOccurrences: 0, contexts: [], text: "" };
        }

        const results = searchResult.results.filter(function (result) {
            return result && typeof result.text === "string";
        }).slice().sort(function (a, b) {
            return Number(b.score || 0) - Number(a.score || 0);
        });

        const selected = [];
        const selectedParagraphIndexes = new Set();
        const MAX_TEXT_OVERLAP = 0.75;

        for (let i = 0; i < results.length; i++) {
            const candidate = results[i];
            if (selectedParagraphIndexes.has(candidate.paragraphIndex)) continue;

            const candidateText = String(candidate.context || candidate.text || "").replace(/\s+/g, " ").trim();
            if (!candidateText) continue;

            let tooSimilar = false;
            for (let j = 0; j < selected.length; j++) {
                const selectedText = String(selected[j].context || selected[j].text || "").replace(/\s+/g, " ").trim();
                if (!selectedText) continue;
                const shorterLength = Math.min(candidateText.length, selectedText.length);
                const commonLength = getCommonTextLength(candidateText, selectedText);
                const overlap = shorterLength > 0 ? commonLength / shorterLength : 0;
                if (overlap >= MAX_TEXT_OVERLAP) {
                    tooSimilar = true;
                    break;
                }
            }
            if (tooSimilar) continue;

            selected.push(candidate);
            selectedParagraphIndexes.add(candidate.paragraphIndex);
            if (selected.length >= maxResults) break;
        }

        const contexts = [];
        let totalChars = 0;

        selected.forEach(function (result, index) {
            let mainContext = String(result.context || result.text || "").replace(/\s+/g, " ").trim();
            if (!mainContext) return;

            let previousContext = includeNeighbors ? String(result.previousParagraphText || "").replace(/\s+/g, " ").trim() : "";
            let nextContext = includeNeighbors ? String(result.nextParagraphText || "").replace(/\s+/g, " ").trim() : "";

            if (selectedParagraphIndexes.has(Number(result.paragraphIndex) - 1)) previousContext = "";
            if (selectedParagraphIndexes.has(Number(result.paragraphIndex) + 1)) nextContext = "";

            const heading = String(result.heading || "").trim();
            const remainingChars = maxChars - totalChars;
            if (remainingChars <= 0) return;

            const reservedForMetadata = 250;
            const availableChars = Math.max(300, remainingChars - reservedForMetadata);

            let context = mainContext;
            let remainingForNeighbors = availableChars - context.length;

            if (includeNeighbors && previousContext && remainingForNeighbors > 150) {
                const separatorLength = 1;
                const allowedPreviousLength = Math.max(0, remainingForNeighbors - separatorLength);
                if (allowedPreviousLength > 100) {
                    const previousPart = previousContext.length > allowedPreviousLength ?
                        previousContext.substring(Math.max(0, previousContext.length - allowedPreviousLength)) + "…" :
                        previousContext;
                    context = previousPart + " " + context;
                }
            }

            remainingForNeighbors = availableChars - context.length;

            if (includeNeighbors && nextContext && remainingForNeighbors > 150) {
                const separatorLength = 1;
                const allowedNextLength = Math.max(0, remainingForNeighbors - separatorLength);
                if (allowedNextLength > 100) {
                    const nextPart = nextContext.length > allowedNextLength ?
                        nextContext.substring(0, allowedNextLength) + "…" :
                        nextContext;
                    context = context + " " + nextPart;
                }
            }

            contexts.push({
                rank: index + 1,
                paragraphIndex: result.paragraphIndex,
                heading: heading,
                score: Number(result.score || 0),
                matchType: result.matchType || "family",
                matchedFamilies: Array.isArray(result.matchedFamilies) ? result.matchedFamilies : [],
                familyOccurrences: Number(result.familyOccurrencesInParagraph || 0),
                exactWordMatches: Number(result.exactWordMatches || 0),
                previousParagraph: previousContext,
                mainParagraph: mainContext,
                nextParagraph: nextContext,
                context: context
            });

            totalChars += context.length;
        });

        const textParts = [];
        contexts.forEach(function (item) {
            let block = "[مقطع " + item.rank + "]\n";
            if (item.heading) block += "العنوان: " + item.heading + "\n";
            if (item.previousParagraph) block += "السياق السابق: " + item.previousParagraph + "\n";
            block += "المقطع المطابق: " + item.mainParagraph;
            if (item.nextParagraph) block += "\nالسياق التالي: " + item.nextParagraph;
            textParts.push(block);
        });

        const finalText = textParts.join("\n\n---\n\n");

        return {
            query: searchResult.query || "",
            count: searchResult.count || results.length,
            selectedCount: contexts.length,
            totalOccurrences: Number(searchResult.indexedOccurrences || 0),
            contexts: contexts,
            text: finalText
        };
    }

    async function buildAIDocumentContext(question) {
        if (!currentDocument) {
            currentCitationSources = [];
            return { found: false, query: String(question || ""), text: "", sources: [], resultCount: 0, selectedCount: 0 };
        }

        const cleanQuestion = String(question || "").trim();
        if (!cleanQuestion) {
            currentCitationSources = [];
            return { found: false, query: "", text: "", sources: [], resultCount: 0, selectedCount: 0 };
        }

        const retrievalProfile = getRetrievalProfile(cleanQuestion);
        const requestedLimit = retrievalProfile && typeof retrievalProfile.maxResults === "number" ? retrievalProfile.maxResults : 8;

        const results = await searchIndexedDocument(currentDocument.id, cleanQuestion, { profile: retrievalProfile.type });
        if (!results || results.length === 0) {
            currentCitationSources = [];
            return { found: false, query: cleanQuestion, profile: retrievalProfile.type, text: "", sources: [], resultCount: 0, selectedCount: 0 };
        }

        const selected = [];
        const selectedParagraphs = new Set();
        const maxResults = Math.max(1, requestedLimit);

        for (let i = 0; i < results.length && selected.length < maxResults; i++) {
            const item = results[i];
            if (!item) continue;
            const paragraphKey = String(item.paragraphIndex);
            if (selectedParagraphs.has(paragraphKey)) continue;
            selectedParagraphs.add(paragraphKey);
            selected.push(item);
        }

        currentCitationSources = selected.map(function (item, index) {
            return {
                rank: index + 1,
                paragraphIndex: item.paragraphIndex,
                paragraphId: item.paragraphId || "",
                heading: item.heading || "",
                mainParagraph: item.text || "",
                text: item.text || "",
                score: item.score || 0
            };
        });

        const textParts = [];
        selected.forEach(function (item, index) {
            const block = [
                "[مقطع " + (index + 1) + "]",
                item.heading ? "العنوان: " + item.heading : "",
                "المقطع: " + String(item.text || "").trim()
            ].filter(function (value) { return Boolean(String(value || "").trim()); }).join("\n");
            if (block) textParts.push(block);
        });

        const contextText = textParts.join("\n\n---\n\n");

        return {
            found: selected.length > 0,
            query: cleanQuestion,
            profile: retrievalProfile.type,
            text: contextText,
            sources: currentCitationSources,
            resultCount: results.length,
            selectedCount: selected.length,
            totalOccurrences: results.length,
            matchedFamilies: [],
            matchedTerms: []
        };
    }

    // =====================================================
    // FORMAT AI MESSAGE
    // =====================================================

    function formatAIMessage(text, citationSources) {
        if (!text) return "";

        const sources = Array.isArray(citationSources) ? citationSources : currentCitationSources;

        try {
            let html = marked.parse(String(text), { breaks: true, gfm: true });

            html = html.replace(/\[مقطع\s*([0-9٠-٩\s،,]+)\]/g, function (match, ranksText) {
                const normalized = String(ranksText)
                    .replace(/٠/g, "0").replace(/١/g, "1").replace(/٢/g, "2")
                    .replace(/٣/g, "3").replace(/٤/g, "4").replace(/٥/g, "5")
                    .replace(/٦/g, "6").replace(/٧/g, "7").replace(/٨/g, "8").replace(/٩/g, "9");

                const ranks = normalized.split(/[،,]+/).map(function (value) {
                    return Number(value.trim());
                }).filter(function (value) { return !Number.isNaN(value); });

                if (ranks.length === 0) return match;

                const citationButtons = ranks.map(function (rank) {
                    const source = sources.find(function (item) { return Number(item.rank) === rank; });
                    if (!source) return "[مقطع " + rank + "]";
                    return `<button type="button" class="document-citation" data-citation-rank="${rank}" title="الانتقال إلى المقطع ${rank}">[مقطع ${rank}]</button>`;
                });

                return citationButtons.join(" ");
            });

            return html;
        } catch (error) {
            console.error("Markdown formatting error:", error);
            return String(text).replace(/\n/g, "<br>");
        }
    }

    // =====================================================
    // OPEN CITATION IN WORD
    // =====================================================

    async function openCitationInWord(rank) {
        const source = currentCitationSources.find(function (item) {
            return Number(item.rank) === Number(rank);
        });

        if (!source) {
            console.warn("لم يتم العثور على مصدر الإحالة:", rank);
            return;
        }

        if (typeof Word === "undefined") {
            console.warn("Word API غير متاحة.");
            return;
        }

        try {
            await Word.run(async function (context) {
                const body = context.document.body;
                let searchText = String(source.mainParagraph || source.text || "").replace(/\s+/g, " ").trim();
                if (!searchText) throw new Error("لا يوجد نص صالح للمقطع.");
                if (searchText.length > 180) searchText = searchText.substring(0, 180).trim();

                const results = body.search(searchText, {
                    matchCase: false,
                    matchWholeWord: false,
                    matchWildcards: false,
                    ignorePunct: true,
                    ignoreSpace: true
                });
                results.load("items");
                await context.sync();

                if (results.items.length === 0) {
                    const fallbackText = searchText.substring(0, 80);
                    const fallbackResults = body.search(fallbackText, {
                        matchCase: false,
                        matchWholeWord: false,
                        matchWildcards: false,
                        ignorePunct: true,
                        ignoreSpace: true
                    });
                    fallbackResults.load("items");
                    await context.sync();
                    if (fallbackResults.items.length === 0) {
                        throw new Error("لم يتم العثور على نص المقطع في المستند.");
                    }
                    fallbackResults.items[0].select("Select");
                    await context.sync();
                    return;
                }

                results.items[0].select("Select");
                await context.sync();
            });
        } catch (error) {
            console.error("تعذر الانتقال إلى المقطع:", error);
        }
    }

    // =====================================================
    // RENDER FUNCTIONS
    // =====================================================

    function renderProjects() {
        if (!projectsList) return;
        projectsList.innerHTML = "";
        if (projects.length === 0) {
            projectsList.innerHTML = `<div class="empty-project">لا توجد مشاريع</div>`;
            return;
        }

        projects.forEach(function (project) {
            const item = document.createElement("div");
            item.className = "project-item";
            item.innerHTML = `
                <span class="project-title">${projectIcon} ${project.name}</span>
                <button class="project-menu" type="button">⋮</button>
                <div class="project-options-menu">
                    <div class="rename-project">✏ إعادة تسمية</div>
                    <div class="delete-project">🗑 حذف</div>
                </div>
            `;

            item.onclick = function (e) {
                if (e.target.closest(".project-menu") || e.target.closest(".project-options-menu")) return;
                e.stopPropagation();
                setCurrentProject(project);
                if (projectsPopup) projectsPopup.classList.remove("open");
            };

            const menu = item.querySelector(".project-menu");
            if (menu) {
                menu.onclick = function (e) {
                    e.preventDefault();
                    e.stopPropagation();
                    document.querySelectorAll(".project-options-menu.open").forEach(function (menuItem) {
                        menuItem.classList.remove("open");
                    });
                    const options = item.querySelector(".project-options-menu");
                    if (!options) return;
                    options.classList.add("open");
                    const rect = menu.getBoundingClientRect();
                    const menuWidth = 140;
                    const menuHeight = options.offsetHeight || 80;
                    const margin = 8;
                    let left = rect.left - menuWidth - margin;
                    let top = rect.bottom + margin;
                    if (left < margin) left = rect.right + margin;
                    if (left + menuWidth > window.innerWidth - margin) left = window.innerWidth - menuWidth - margin;
                    if (top + menuHeight > window.innerHeight - margin) top = rect.top - menuHeight - margin;
                    if (top < margin) top = margin;
                    options.style.position = "fixed";
                    options.style.left = left + "px";
                    options.style.top = top + "px";
                    options.style.right = "auto";
                    options.style.bottom = "auto";
                    options.style.zIndex = "999999";
                };
            }

            // Rename Project
            const renameProject = item.querySelector(".rename-project");
            if (renameProject) {
                renameProject.onclick = function (e) {
                    e.preventDefault();
                    e.stopPropagation();
                    const options = item.querySelector(".project-options-menu");
                    if (options) options.classList.remove("open");
                    const titleElement = item.querySelector(".project-title");
                    if (!titleElement) return;
                    const oldName = project.name;
                    titleElement.innerHTML = `<input class="edit-project-title" value="${oldName}">`;
                    const edit = titleElement.querySelector(".edit-project-title");
                    if (!edit) return;
                    edit.focus();
                    edit.setSelectionRange(edit.value.length, edit.value.length);
                    edit.onkeydown = function (event) {
                        if (event.key === "Enter") {
                            event.preventDefault();
                            const value = edit.value.trim();
                            project.name = value || oldName;
                            project.updatedAt = new Date().toISOString();
                            saveProjects();
                            renderProjects();
                            renderExpandedProjects();
                        }
                        if (event.key === "Escape") {
                            event.preventDefault();
                            project.name = oldName;
                            renderProjects();
                        }
                    };
                };
            }

            // Delete Project
            const deleteProject = item.querySelector(".delete-project");
            if (deleteProject) {
                deleteProject.onclick = function (e) {
                    e.preventDefault();
                    e.stopPropagation();
                    const options = item.querySelector(".project-options-menu");
                    if (options) options.classList.remove("open");
                    const confirmBox = document.createElement("div");
                    confirmBox.className = "project-delete-confirm";
                    confirmBox.innerHTML = `
                        <div class="confirm-dialog">
                            <p>هل تريد حذف المشروع:<br><strong>${project.name}</strong>؟</p>
                            <button class="confirm-project-delete" type="button">حذف</button>
                            <button class="cancel-project-delete" type="button">إلغاء</button>
                        </div>
                    `;
                    document.body.appendChild(confirmBox);

                    const confirmDelete = confirmBox.querySelector(".confirm-project-delete");
                    if (confirmDelete) {
                        confirmDelete.onclick = async function () {
                            const projectDocumentIds = Array.isArray(project.documents) ? project.documents : [];
                            for (let i = 0; i < projectDocumentIds.length; i++) {
                                const documentItem = documents.find(function (d) {
                                    return String(d.id) === String(projectDocumentIds[i]);
                                });
                                if (documentItem) {
                                    try {
                                        await deleteWorkingWordFile(documentItem.storageId);
                                    } catch (error) {
                                        console.warn("تعذر حذف نسخة العمل:", error);
                                    }
                                }
                            }
                            documents = documents.filter(function (doc) {
                                return !projectDocumentIds.some(function (id) { return String(id) === String(doc.id); });
                            });
                            projects = projects.filter(function (p) { return String(p.id) !== String(project.id); });
                            if (currentProject && String(currentProject.id) === String(project.id)) {
                                currentProject = null;
                                currentDocument = null;
                                currentCitationSources = [];
                                if (documentTitle) documentTitle.textContent = "لا يوجد مستند مفتوح";
                            }
                            oramaRetrievalDb = null;
                            oramaRetrievalCacheKey = "";
                            oramaRetrievalDocumentId = null;
                            saveDocuments();
                            saveProjects();
                            renderProjects();
                            renderExpandedProjects();
                            renderDocuments();
                            confirmBox.remove();
                        };
                    }

                    const cancelDelete = confirmBox.querySelector(".cancel-project-delete");
                    if (cancelDelete) {
                        cancelDelete.onclick = function () { confirmBox.remove(); };
                    }
                };
            }

            projectsList.appendChild(item);
        });
    }

    function renderExpandedProjects() {
        const list = document.getElementById("expanded-projects-list");
        if (!list) return;
        list.innerHTML = "";
        projects.forEach(function (project) {
            const item = document.createElement("div");
            item.className = "expanded-project-item";
            item.innerHTML = `<span>${projectIcon} ${project.name}</span>`;
            item.onclick = function (e) {
                e.stopPropagation();
                setCurrentProject(project);
            };
            list.appendChild(item);
        });
    }

    function renderDocuments() {
        if (!documentsList) return;
        documentsList.innerHTML = "";
        if (!currentProject) {
            documentsList.innerHTML = `<div class="empty-document">اختر مشروعًا لعرض مستنداته</div>`;
            return;
        }

        const projectDocuments = getProjectDocuments(currentProject.id);
        if (projectDocuments.length === 0) {
            documentsList.innerHTML = `<div class="empty-document">لا توجد مستندات</div>`;
            return;
        }

        projectDocuments.forEach(function (documentItem, index) {
            const item = document.createElement("div");
            item.className = "document-item";
            if (currentDocument && String(currentDocument.id) === String(documentItem.id)) {
                item.classList.add("active-document");
            }

            const title = document.createElement("span");
            title.className = "document-title";
            title.textContent = documentItem.name;

            const status = document.createElement("span");
            status.className = "document-read-status";
            if (documentItem.indexStatus === "indexed") {
                status.textContent = "✓ مفهرس · " + documentItem.indexTokenCount + " كلمة · " + documentItem.indexUniqueTerms + " فريدة";
                if (documentItem.indexUniqueFamilies) {
                    status.textContent += " · " + documentItem.indexUniqueFamilies + " عائلة";
                }
            } else if (documentItem.indexStatus === "indexing") {
                status.textContent = "جارٍ الفهرسة...";
            } else if (documentItem.indexStatus === "error") {
                status.textContent = "⚠ فشل الفهرسة";
            } else if (documentItem.readStatus === "reading") {
                status.textContent = "جارٍ القراءة...";
            } else if (documentItem.readStatus === "read") {
                status.textContent = "✓ تمت القراءة";
            } else {
                status.textContent = "جديد";
            }

            title.onclick = function (e) {
                e.preventDefault();
                e.stopPropagation();
                setCurrentDocument(documentItem);
            };

            const menuButton = document.createElement("button");
            menuButton.className = "document-menu";
            menuButton.type = "button";
            menuButton.title = "خيارات المستند";
            menuButton.textContent = "⋮";

            const options = document.createElement("div");
            options.className = "document-options-menu";
            options.innerHTML = `
                <div class="rename-document">✏ إعادة تسمية</div>
                <div class="move-document-up">↑ نقل إلى أعلى</div>
                <div class="move-document-down">↓ نقل إلى أسفل</div>
                <div class="delete-document">🗑 حذف</div>
            `;

            if (index === 0) {
                const moveUp = options.querySelector(".move-document-up");
                if (moveUp) moveUp.style.display = "none";
            }
            if (index === projectDocuments.length - 1) {
                const moveDown = options.querySelector(".move-document-down");
                if (moveDown) moveDown.style.display = "none";
            }

            menuButton.onclick = function (e) {
                e.preventDefault();
                e.stopPropagation();
                document.querySelectorAll(".document-options-menu.open").forEach(function (menu) {
                    if (menu !== options) menu.classList.remove("open");
                });
                options.classList.toggle("open");
            };

            // Rename Document
            const renameButton = options.querySelector(".rename-document");
            if (renameButton) {
                renameButton.onclick = function (e) {
                    e.preventDefault();
                    e.stopPropagation();
                    options.classList.remove("open");
                    const oldName = documentItem.name;
                    const inputRename = document.createElement("input");
                    inputRename.className = "edit-document-title";
                    inputRename.value = oldName;
                    title.replaceWith(inputRename);
                    inputRename.focus();
                    inputRename.setSelectionRange(inputRename.value.length, inputRename.value.length);

                    function finishRename(saveChange) {
                        const newName = inputRename.value.trim();
                        if (saveChange && newName) {
                            documentItem.name = newName;
                            documentItem.updatedAt = new Date().toISOString();
                            saveDocuments();
                        } else {
                            documentItem.name = oldName;
                        }
                        renderDocuments();
                    }

                    inputRename.onkeydown = function (event) {
                        if (event.key === "Enter") { event.preventDefault(); finishRename(true); }
                        if (event.key === "Escape") { event.preventDefault(); finishRename(false); }
                    };
                };
            }

            // Move Up
            const moveUpButton = options.querySelector(".move-document-up");
            if (moveUpButton) {
                moveUpButton.onclick = function (e) {
                    e.preventDefault();
                    e.stopPropagation();
                    if (index <= 0) return;
                    const previousDocument = projectDocuments[index - 1];
                    const currentOrder = documentItem.order;
                    documentItem.order = previousDocument.order;
                    previousDocument.order = currentOrder;
                    documentItem.updatedAt = new Date().toISOString();
                    previousDocument.updatedAt = new Date().toISOString();
                    saveDocuments();
                    options.classList.remove("open");
                    renderDocuments();
                };
            }

            // Move Down
            const moveDownButton = options.querySelector(".move-document-down");
            if (moveDownButton) {
                moveDownButton.onclick = function (e) {
                    e.preventDefault();
                    e.stopPropagation();
                    if (index >= projectDocuments.length - 1) return;
                    const nextDocument = projectDocuments[index + 1];
                    const currentOrder = documentItem.order;
                    documentItem.order = nextDocument.order;
                    nextDocument.order = currentOrder;
                    documentItem.updatedAt = new Date().toISOString();
                    nextDocument.updatedAt = new Date().toISOString();
                    saveDocuments();
                    options.classList.remove("open");
                    renderDocuments();
                };
            }

            // Delete Document
            const deleteButton = options.querySelector(".delete-document");
            if (deleteButton) {
                deleteButton.onclick = function (e) {
                    e.preventDefault();
                    e.stopPropagation();
                    options.classList.remove("open");

                    const confirmBox = document.createElement("div");
                    confirmBox.className = "document-delete-confirm";
                    confirmBox.innerHTML = `
                        <div class="document-delete-dialog">
                            <div class="document-delete-message">هل تريد حذف المستند؟</div>
                            <div class="document-delete-name">${documentItem.name}</div>
                            <div class="document-delete-buttons">
                                <button type="button" class="confirm-document-delete">حذف</button>
                                <button type="button" class="cancel-document-delete">إلغاء</button>
                            </div>
                        </div>
                    `;
                    document.body.appendChild(confirmBox);

                    const confirmDelete = confirmBox.querySelector(".confirm-document-delete");
                    if (confirmDelete) {
                        confirmDelete.onclick = async function () {
                            documents = documents.filter(function (doc) {
                                return String(doc.id) !== String(documentItem.id);
                            });

                            if (currentProject && Array.isArray(currentProject.documents)) {
                                currentProject.documents = currentProject.documents.filter(function (id) {
                                    return String(id) !== String(documentItem.id);
                                });
                                currentProject.updatedAt = new Date().toISOString();
                                saveProjects();
                            }

                            try {
                                await deleteWorkingWordFile(documentItem.storageId);
                            } catch (storageError) {
                                console.warn("تعذر حذف نسخة العمل:", storageError);
                            }

                            if (currentDocument && String(currentDocument.id) === String(documentItem.id)) {
                                currentDocument = null;
                                currentCitationSources = [];
                                if (documentTitle) documentTitle.textContent = "لا يوجد مستند مفتوح";
                            }

                            const remaining = getProjectDocuments(currentProject ? currentProject.id : null);
                            remaining.forEach(function (doc, newIndex) { doc.order = newIndex + 1; });

                            oramaRetrievalDb = null;
                            oramaRetrievalCacheKey = "";
                            oramaRetrievalDocumentId = null;

                            saveDocuments();
                            confirmBox.remove();
                            renderDocuments();
                        };
                    }

                    const cancelDelete = confirmBox.querySelector(".cancel-document-delete");
                    if (cancelDelete) {
                        cancelDelete.onclick = function () { confirmBox.remove(); };
                    }

                    confirmBox.onclick = function (event) {
                        if (event.target === confirmBox) confirmBox.remove();
                    };
                };
            }

            item.appendChild(title);
            item.appendChild(status);
            item.appendChild(menuButton);
            item.appendChild(options);
            documentsList.appendChild(item);
        });
    }

    function renderSidebarChats() {
        const list = document.getElementById("new-chat-list");
        if (!list) return;
        list.innerHTML = "";
        if (chats.length === 0) {
            list.innerHTML = `<div class="empty-chat">لا توجد محادثات</div>`;
            return;
        }

        chats.slice(0, 8).forEach(function (chat) {
            const item = document.createElement("div");
            item.className = "recent-chat-item";
            item.innerHTML = `<span class="chat-title">${chatIcon} ${chat.title}</span>`;
            item.onclick = function (e) {
                e.stopPropagation();
                currentChat = chat;
                renderChat();
                if (projectsPopup) projectsPopup.classList.remove("open");
                if (chatPopup) chatPopup.classList.remove("open");
                if (searchPopup) searchPopup.classList.remove("open");
            };
            list.appendChild(item);
        });
    }

    function renderChat() {
        if (!chatArea) return;
        chatArea.innerHTML = "";
        if (!currentChat) {
            chatArea.innerHTML = `
                <div class="welcome">
                    <div class="ai-symbol">✦</div>
                    <h2>مرحبًا بك</h2>
                    <p>ابدأ محادثة جديدة</p>
                </div>
            `;
            return;
        }

        currentChat.messages.forEach(function (msg) {
            const div = document.createElement("div");
            div.className = "message " + (msg.role === "user" ? "user-message" : "ai-message");
            if (msg.role === "user") {
                div.textContent = msg.text || "";
            } else {
                div.innerHTML = formatAIMessage(msg.text || "", Array.isArray(msg.citationSources) ? msg.citationSources : []);
            }
            chatArea.appendChild(div);
        });
        chatArea.scrollTop = chatArea.scrollHeight;
    }

    function renderChatList() {
        const list = document.getElementById("chat-list");
        if (!list) return;
        list.innerHTML = "";
        if (chats.length === 0) {
            list.innerHTML = "<div class='empty-chat'>لا توجد محادثات</div>";
            return;
        }

        chats.forEach(function (chat) {
            const item = document.createElement("div");
            item.className = "chat-history-item";
            item.innerHTML = `
                <span class="chat-title">${chatIcon} ${chat.title}</span>
                <button class="chat-menu" type="button">⋮</button>
                <div class="chat-options-menu">
                    <div class="rename-chat">✏ إعادة تسمية</div>
                    <div class="delete-chat">🗑 حذف</div>
                </div>
            `;

            const title = item.querySelector(".chat-title");
            if (title) {
                title.onclick = function (e) {
                    e.stopPropagation();
                    currentChat = chat;
                    renderChat();
                    if (expandedSidebar) expandedSidebar.classList.remove("open");
                };
            }

            const menu = item.querySelector(".chat-menu");
            if (menu) {
                menu.onclick = function (e) {
                    e.stopPropagation();
                    document.querySelectorAll(".chat-options-menu").forEach(function (m) { m.classList.remove("open"); });
                    const options = item.querySelector(".chat-options-menu");
                    if (!options) return;
                    options.classList.add("open");
                    const rect = menu.getBoundingClientRect();
                    options.style.position = "fixed";
                    options.style.left = Math.max(8, rect.left - 140 - 8) + "px";
                    options.style.top = rect.bottom + 8 + "px";
                    options.style.zIndex = "999999";
                };
            }

            // Rename Chat
            const renameBtn = item.querySelector(".rename-chat");
            if (renameBtn) {
                renameBtn.onclick = function (e) {
                    e.stopPropagation();
                    const options = item.querySelector(".chat-options-menu");
                    if (options) options.classList.remove("open");
                    const titleSpan = item.querySelector(".chat-title");
                    if (!titleSpan) return;
                    const oldName = chat.title;
                    titleSpan.innerHTML = `<input class="edit-chat-title" value="${oldName}">`;
                    const editInput = titleSpan.querySelector(".edit-chat-title");
                    if (!editInput) return;
                    editInput.focus();
                    editInput.setSelectionRange(editInput.value.length, editInput.value.length);
                    editInput.onkeydown = function (event) {
                        if (event.key === "Enter") {
                            const value = editInput.value.trim();
                            chat.title = value !== "" ? value : oldName;
                            saveChats();
                            renderChatList();
                            renderSidebarChats();
                            renderRecentChats();
                        }
                        if (event.key === "Escape") {
                            chat.title = oldName;
                            renderChatList();
                        }
                    };
                };
            }

            // Delete Chat
            const deleteBtn = item.querySelector(".delete-chat");
            if (deleteBtn) {
                deleteBtn.onclick = function (e) {
                    e.stopPropagation();
                    const options = item.querySelector(".chat-options-menu");
                    if (options) options.classList.remove("open");
                    const confirmBox = document.createElement("div");
                    confirmBox.className = "delete-confirm";
                    confirmBox.innerHTML = `
                        <div class="confirm-dialog">
                            <p>هل تريد حذف المحادثة:<br><strong>${chat.title}</strong>؟</p>
                            <button class="confirm-delete" type="button">حذف</button>
                            <button class="cancel-delete" type="button">إلغاء</button>
                        </div>
                    `;
                    document.body.appendChild(confirmBox);

                    const confirmDelete = confirmBox.querySelector(".confirm-delete");
                    if (confirmDelete) {
                        confirmDelete.onclick = function () {
                            chats = chats.filter(function (c) { return c.id !== chat.id; });
                            if (currentChat && currentChat.id === chat.id) {
                                currentChat = null;
                                renderChat();
                            }
                            saveChats();
                            renderChatList();
                            renderSidebarChats();
                            renderRecentChats();
                            confirmBox.remove();
                        };
                    }

                    const cancelDelete = confirmBox.querySelector(".cancel-delete");
                    if (cancelDelete) {
                        cancelDelete.onclick = function () { confirmBox.remove(); };
                    }
                };
            }

            list.appendChild(item);
        });
    }

    function renderRecentChats() {
        if (!recentChatList) return;
        recentChatList.innerHTML = "";
        if (chats.length === 0) {
            recentChatList.innerHTML = "<div class='empty-chat'>لا توجد محادثات</div>";
            return;
        }

        chats.slice(0, 8).forEach(function (chat) {
            const div = document.createElement("div");
            div.className = "recent-chat-item";
            div.innerHTML = `<span class="chat-title">${chatIcon} ${chat.title}</span>`;
            div.onclick = function () {
                currentChat = chat;
                renderChat();
                if (chatPopup) chatPopup.classList.remove("open");
            };
            recentChatList.appendChild(div);
        });
    }

    // =====================================================
    // CREATE NEW CHAT
    // =====================================================

    function createNewChat() {
        currentChat = {
            id: Date.now(),
            title: "محادثة جديدة",
            messages: [],
            isTemporary: true,
            projectId: currentProject ? currentProject.id : null
        };
        currentCitationSources = [];
        if (input) { input.value = ""; input.style.height = "auto"; }
        renderSidebarChats();
        renderRecentChats();
        renderChat();
    }

    // =====================================================
    // CHAT BUTTON HANDLERS
    // =====================================================

    if (newChatBtn) {
        newChatBtn.onclick = function (e) {
            e.preventDefault();
            e.stopPropagation();
            if (projectsPopup) projectsPopup.classList.remove("open");
            if (chatPopup) chatPopup.classList.remove("open");
            if (searchPopup) searchPopup.classList.remove("open");
            if (settingsWindow) settingsWindow.classList.remove("open");
            createNewChat();
        };
    }

    if (chatBtn) {
        chatBtn.onclick = function (e) {
            e.stopPropagation();
            if (!chatPopup) return;
            if (projectsPopup) projectsPopup.classList.remove("open");
            if (searchPopup) searchPopup.classList.remove("open");
            if (settingsWindow) settingsWindow.classList.remove("open");
            if (chatPopup.classList.contains("open")) {
                chatPopup.classList.remove("open");
                return;
            }
            renderChatList();
            chatPopup.classList.add("open");
        };
    }

    // =====================================================
    // PROJECTS BUTTON
    // =====================================================

    if (projectsBtn) {
        projectsBtn.onclick = function (e) {
            e.preventDefault();
            e.stopPropagation();
            if (!projectsPopup) return;
            if (chatPopup) chatPopup.classList.remove("open");
            if (searchPopup) searchPopup.classList.remove("open");
            if (settingsWindow) settingsWindow.classList.remove("open");
            projectsPopup.classList.toggle("open");
            renderProjects();
        };
    }

    // =====================================================
    // NEW PROJECT
    // =====================================================

    if (newProjectBtn) {
        newProjectBtn.onclick = function (e) {
            e.preventDefault();
            e.stopPropagation();
            const oldBox = document.querySelector(".project-create-box");
            if (oldBox) oldBox.remove();

            const box = document.createElement("div");
            box.className = "project-create-box";
            box.innerHTML = `
                <div class="rename-dialog">
                    <input class="new-project-name" placeholder="اسم المشروع">
                    <button class="save-project" type="button">حفظ</button>
                    <button class="cancel-project" type="button">إلغاء</button>
                </div>
            `;
            document.body.appendChild(box);

            const buttonRect = newProjectBtn.getBoundingClientRect();
            const screenMargin = 12;
            let left = buttonRect.left;
            let top = buttonRect.bottom + 8;
            const actualBoxWidth = box.offsetWidth || 240;
            const boxHeight = box.offsetHeight || 120;

            if (left + actualBoxWidth > window.innerWidth - screenMargin) {
                left = window.innerWidth - actualBoxWidth - screenMargin;
            }
            if (left < screenMargin) left = screenMargin;
            if (top + boxHeight > window.innerHeight - screenMargin) {
                top = buttonRect.top - boxHeight - 8;
            }

            box.style.position = "fixed";
            box.style.left = left + "px";
            box.style.top = top + "px";
            box.style.zIndex = "999999";

            const inputProject = box.querySelector(".new-project-name");
            if (inputProject) inputProject.focus();

            const saveProjectButton = box.querySelector(".save-project");
            if (saveProjectButton) {
                saveProjectButton.onclick = function () {
                    const name = inputProject ? inputProject.value.trim() : "";
                    if (name) {
                        const now = new Date().toISOString();
                        const newProject = {
                            id: Date.now(),
                            name: name,
                            createdAt: now,
                            updatedAt: now,
                            documents: [],
                            references: [],
                            chatIds: [],
                            settings: { citationStyle: "", notes: "" }
                        };
                        projects.unshift(newProject);
                        saveProjects();
                        renderProjects();
                        renderExpandedProjects();
                    }
                    box.remove();
                };
            }

            const cancelProject = box.querySelector(".cancel-project");
            if (cancelProject) {
                cancelProject.onclick = function () { box.remove(); };
            }

            box.onclick = function (event) { event.stopPropagation(); };
        };
    }

    // =====================================================
    // ADD DOCUMENT
    // =====================================================

    if (addDocumentBtn && wordDocumentPicker) {
        addDocumentBtn.onclick = function (e) {
            e.preventDefault();
            e.stopPropagation();
            if (!currentProject) {
                if (documentsList) {
                    documentsList.innerHTML = `<div class="empty-document">اختر مشروعًا أولًا لإضافة مستند</div>`;
                }
                return;
            }
            wordDocumentPicker.value = "";
            wordDocumentPicker.click();
        };

        wordDocumentPicker.onchange = async function () {
            try {
                const file = wordDocumentPicker.files && wordDocumentPicker.files[0];
                if (!file) return;
                if (!/\.docx$/i.test(file.name)) {
                    console.warn("الملف المختار ليس DOCX.");
                    return;
                }
                if (!currentProject) return;

                const projectDocuments = getProjectDocuments(currentProject.id);
                const nextOrder = projectDocuments.length + 1;
                const documentItem = createDocument(file, currentProject.id, nextOrder);
                await saveWorkingWordFile(documentItem.storageId, file);
                attachDocumentToProject(currentProject, documentItem);
                setCurrentDocument(documentItem);
                renderDocuments();
                console.log("تم استيراد مستند Word:", {
                    name: documentItem.name,
                    fileName: documentItem.fileName,
                    storageId: documentItem.storageId
                });
            } catch (error) {
                console.error("فشل استيراد مستند Word:", error);
                if (documentsList) {
                    documentsList.innerHTML = `<div class="empty-document">تعذر استيراد المستند</div>`;
                }
            }
        };
    }

    // =====================================================
    // SIDEBAR TOGGLE
    // =====================================================

    if (sidebarToggleBtn && expandedSidebar && expandedSidebarToggleSlot) {
        sidebarToggleBtn.onclick = function (e) {
            e.preventDefault();
            e.stopPropagation();
            const isOpening = !expandedSidebar.classList.contains("open");

            if (isOpening) {
                expandedSidebar.classList.add("open");
                document.body.classList.add("expanded-sidebar-open");
                sidebarToggleBtn.title = "إخفاء القائمة";
                sidebarToggleBtn.classList.add("sidebar-open");
                expandedSidebarToggleSlot.appendChild(sidebarToggleBtn);
            } else {
                expandedSidebar.classList.remove("open");
                document.body.classList.remove("expanded-sidebar-open");
                sidebarToggleBtn.title = "إظهار القائمة";
                sidebarToggleBtn.classList.remove("sidebar-open");
                if (sidebarTogglePlaceholder.parentNode) {
                    sidebarTogglePlaceholder.parentNode.insertBefore(
                        sidebarToggleBtn,
                        sidebarTogglePlaceholder.nextSibling
                    );
                }
            }
        };
    }

    // =====================================================
    // SIDEBAR SECTIONS
    // =====================================================

    function initializeSidebarSections() {
        const headers = document.querySelectorAll(".section-title[data-target], .section-toggle[data-target]");
        headers.forEach(function (header) {
            const targetId = header.getAttribute("data-target");
            if (!targetId) return;
            const target = document.getElementById(targetId);
            if (!target) return;
            target.classList.remove("open");
            header.classList.remove("open");

            header.onclick = function (e) {
                e.preventDefault();
                e.stopPropagation();
                const isOpen = target.classList.contains("open");
                target.classList.toggle("open", !isOpen);
                header.classList.toggle("open", !isOpen);
            };
        });
    }

    // =====================================================
    // SETTINGS
    // =====================================================

    function loadSettings() {
        const data = getSavedSettings();
        if (provider) provider.value = data.provider || "openrouter";
        if (apiKey) apiKey.value = data.key || "";
        if (modelSelect) {
            const savedModel = data.model || "";
            modelSelect.innerHTML = "";
            if (savedModel !== "") {
                const option = document.createElement("option");
                option.value = savedModel;
                option.textContent = savedModel;
                modelSelect.appendChild(option);
                modelSelect.value = savedModel;
            } else {
                const option = document.createElement("option");
                option.value = "";
                option.textContent = "أدخل المفتاح ثم حدّث النماذج";
                modelSelect.appendChild(option);
            }
        }
        updateProviderInfo();
    }

    function updateProviderInfo() {
        if (!providerInfo || !provider) return;
        const value = provider.value;
        if (value === "openrouter") {
            providerInfo.innerHTML = "OpenRouter: سيتم جلب النماذج المجانية المتاحة من حسابك.";
        } else if (value === "gemini") {
            providerInfo.innerHTML = "Gemini: سيتم جلب النماذج التي تدعم generateContent.";
        } else if (value === "groq") {
            providerInfo.innerHTML = "Groq: سيتم جلب النماذج المتاحة من حسابك.";
        } else if (value === "openai") {
            providerInfo.innerHTML = "OpenAI: سيتم جلب النماذج المتاحة من حسابك.";
        } else {
            providerInfo.innerHTML = "سيتم تحديد رابط الاتصال حسب مزود الذكاء الاصطناعي.";
        }
    }

    if (settingsBtn) {
        settingsBtn.onclick = function (e) {
            e.preventDefault();
            e.stopPropagation();
            if (projectsPopup) projectsPopup.classList.remove("open");
            if (chatPopup) chatPopup.classList.remove("open");
            if (searchPopup) searchPopup.classList.remove("open");
            if (settingsWindow) settingsWindow.classList.add("open");
            loadSettings();
        };
    }

    if (closeSettings) {
        closeSettings.onclick = function (e) {
            e.preventDefault();
            e.stopPropagation();
            if (settingsWindow) settingsWindow.classList.remove("open");
        };
    }

    if (showKey && apiKey) {
        showKey.onclick = function (e) {
            e.preventDefault();
            e.stopPropagation();
            if (apiKey.type === "password") {
                apiKey.type = "text";
                showKey.innerHTML = "🙈";
            } else {
                apiKey.type = "password";
                showKey.innerHTML = "👁";
            }
        };
    }

    if (provider) {
        provider.onchange = function () {
            updateProviderInfo();
            if (modelSelect) {
                modelSelect.innerHTML = `<option value="">أدخل المفتاح ثم حدّث النماذج</option>`;
            }
            if (settingsStatus) settingsStatus.innerHTML = "";
        };
    }

    if (saveSettings) {
        saveSettings.onclick = function (e) {
            e.preventDefault();
            e.stopPropagation();
            const settings = {
                provider: provider ? provider.value : "openrouter",
                key: apiKey ? apiKey.value.trim() : "",
                model: modelSelect ? modelSelect.value.trim() : ""
            };

            if (!settings.key) {
                if (settingsStatus) settingsStatus.innerHTML = "⚠ يرجى إدخال مفتاح API.";
                return;
            }
            if (!settings.model) {
                if (settingsStatus) settingsStatus.innerHTML = "⚠ يرجى تحديد نموذج الذكاء الاصطناعي.";
                return;
            }

            saveAISettings(settings);
            if (settingsStatus) settingsStatus.innerHTML = "✓ تم حفظ إعدادات الذكاء الاصطناعي";
        };
    }

    // =====================================================
    // LOAD MODELS
    // =====================================================

    async function loadModels() {
        const selectedProvider = provider ? provider.value : "openrouter";
        const key = apiKey ? apiKey.value.trim() : "";

        if (!key) {
            if (settingsStatus) settingsStatus.innerHTML = "⚠ يرجى إدخال مفتاح API أولاً.";
            return;
        }

        try {
            if (settingsStatus) settingsStatus.innerHTML = "⏳ جاري تحميل النماذج...";

            let models = [];

            if (selectedProvider === "openrouter") {
                models = await loadOpenRouterModels(key);
            } else if (selectedProvider === "gemini") {
                models = await loadGeminiModels(key);
            } else if (selectedProvider === "groq") {
                models = await loadGroqModels(key);
            } else if (selectedProvider === "openai") {
                models = await loadOpenAIModels(key);
            } else {
                throw new Error("مزود الذكاء الاصطناعي غير معروف.");
            }

            populateModels(models);
            if (settingsStatus) settingsStatus.innerHTML = "✓ تم تحديث النماذج: " + models.length;
        } catch (error) {
            if (settingsStatus) settingsStatus.innerHTML = "⚠ " + (error.message || "تعذر تحديث النماذج");
        }
    }

    async function loadOpenRouterModels(key) {
        const response = await fetch("https://openrouter.ai/api/v1/models", {
            method: "GET",
            headers: {
                "Authorization": "Bearer " + key,
                "Content-Type": "application/json",
                "HTTP-Referer": window.location.href,
                "X-Title": "Research Tools"
            }
        });
        const result = await readJSON(response);
        if (!response.ok) throw new Error(getAPIError(result, "فشل الاتصال بـ OpenRouter."));
        if (!result.data || !Array.isArray(result.data)) throw new Error("لم تصل قائمة النماذج من OpenRouter.");

        const freeModels = result.data.filter(function (item) {
            return item && item.id && String(item.id).endsWith(":free");
        });

        return freeModels.map(function (item) {
            return { id: item.id, name: (item.name || item.id) + " (مجاني)" };
        }).sort(function (a, b) { return String(a.name).localeCompare(String(b.name), "ar"); });
    }

    async function loadGeminiModels(key) {
        const url = "https://generativelanguage.googleapis.com/v1beta/models?key=" + encodeURIComponent(key);
        const response = await fetch(url, { method: "GET" });
        const result = await readJSON(response);
        if (!response.ok) throw new Error(getAPIError(result, "فشل الاتصال بـ Gemini."));
        if (!result.models || !Array.isArray(result.models)) throw new Error("لم تصل قائمة نماذج Gemini.");

        return result.models.filter(function (item) {
            return item && item.name && item.supportedGenerationMethods &&
                item.supportedGenerationMethods.includes("generateContent");
        }).map(function (item) {
            const cleanId = String(item.name).replace(/^models\//, "");
            return { id: cleanId, name: item.displayName ? item.displayName + " — " + cleanId : cleanId };
        });
    }

    async function loadGroqModels(key) {
        const response = await fetch("https://api.groq.com/openai/v1/models", {
            method: "GET",
            headers: { "Authorization": "Bearer " + key, "Content-Type": "application/json" }
        });
        const result = await readJSON(response);
        if (!response.ok) throw new Error(getAPIError(result, "فشل الاتصال بـ Groq."));
        if (!result.data || !Array.isArray(result.data)) throw new Error("لم تصل قائمة نماذج Groq.");

        return result.data.filter(function (item) { return item && item.id && item.active !== false; })
            .sort(function (a, b) { return a.id.localeCompare(b.id); })
            .map(function (item) { return { id: item.id, name: item.id }; });
    }

    async function loadOpenAIModels(key) {
        const response = await fetch("https://api.openai.com/v1/models", {
            method: "GET",
            headers: { "Authorization": "Bearer " + key, "Content-Type": "application/json" }
        });
        const result = await readJSON(response);
        if (!response.ok) throw new Error(getAPIError(result, "فشل الاتصال بـ OpenAI."));
        if (!result.data || !Array.isArray(result.data)) throw new Error("لم تصل قائمة نماذج OpenAI.");

        return result.data.filter(function (item) {
            return item && item.id && (item.id.toLowerCase().startsWith("gpt-") ||
                item.id.toLowerCase().startsWith("o1") ||
                item.id.toLowerCase().startsWith("o3") ||
                item.id.toLowerCase().startsWith("o4"));
        }).sort(function (a, b) { return a.id.localeCompare(b.id); })
            .map(function (item) { return { id: item.id, name: item.id }; });
    }

    function populateModels(models) {
        if (!modelSelect) return;
        const saved = getSavedSettings();
        const savedModel = saved.model || "";

        modelSelect.innerHTML = "";
        if (!models || models.length === 0) {
            const option = document.createElement("option");
            option.value = "";
            option.textContent = "لا توجد نماذج متاحة";
            modelSelect.appendChild(option);
            return;
        }

        models.forEach(function (item) {
            const option = document.createElement("option");
            option.value = item.id;
            option.textContent = item.name;
            modelSelect.appendChild(option);
        });

        if (savedModel) {
            const exists = Array.from(modelSelect.options).some(function (option) { return option.value === savedModel; });
            if (exists) modelSelect.value = savedModel;
        }
    }

    if (refreshModels) {
        refreshModels.onclick = async function (e) {
            e.preventDefault();
            e.stopPropagation();
            refreshModels.disabled = true;
            try {
                await loadModels();
            } catch (error) {
                if (settingsStatus) settingsStatus.innerHTML = "⚠ " + (error.message || "تعذر تحديث النماذج");
            } finally {
                refreshModels.disabled = false;
            }
        };
    }

    // =====================================================
    // TEST CONNECTION
    // =====================================================

    async function testAIConnection() {
        const data = {
            provider: provider ? provider.value : "openrouter",
            key: apiKey ? apiKey.value.trim() : "",
            model: modelSelect ? modelSelect.value.trim() : ""
        };

        if (!data.key) throw new Error("يرجى إدخال مفتاح API أولاً.");
        if (!data.model) throw new Error("يرجى تحديد نموذج الذكاء الاصطناعي أولاً.");

        const testMessage = "أجب بكلمة واحدة فقط: متصل";

        if (data.provider === "openrouter") {
            const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
                method: "POST",
                headers: {
                    "Authorization": "Bearer " + data.key,
                    "Content-Type": "application/json",
                    "HTTP-Referer": window.location.href,
                    "X-Title": "Research Tools"
                },
                body: JSON.stringify({
                    model: data.model,
                    messages: [{ role: "user", content: testMessage }],
                    max_tokens: 10
                })
            });
            const result = await readJSON(response);
            if (!response.ok) throw new Error(getAPIError(result, "فشل الاتصال بـ OpenRouter."));
            return "✓ تم الاتصال بـ OpenRouter بنجاح";
        }

        if (data.provider === "groq") {
            const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
                method: "POST",
                headers: { "Authorization": "Bearer " + data.key, "Content-Type": "application/json" },
                body: JSON.stringify({
                    model: data.model,
                    messages: [{ role: "user", content: testMessage }],
                    max_tokens: 10
                })
            });
            const result = await readJSON(response);
            if (!response.ok) throw new Error(getAPIError(result, "فشل الاتصال بـ Groq."));
            return "✓ تم الاتصال بـ Groq بنجاح";
        }

        if (data.provider === "openai") {
            const response = await fetch("https://api.openai.com/v1/chat/completions", {
                method: "POST",
                headers: { "Authorization": "Bearer " + data.key, "Content-Type": "application/json" },
                body: JSON.stringify({
                    model: data.model,
                    messages: [{ role: "user", content: testMessage }],
                    max_tokens: 10
                })
            });
            const result = await readJSON(response);
            if (!response.ok) throw new Error(getAPIError(result, "فشل الاتصال بـ OpenAI."));
            return "✓ تم الاتصال بـ OpenAI بنجاح";
        }

        if (data.provider === "gemini") {
            const model = normalizeGeminiModel(data.model);
            const url = "https://generativelanguage.googleapis.com/v1beta/models/" +
                encodeURIComponent(model) +
                ":generateContent?key=" + encodeURIComponent(data.key);

            const response = await fetch(url, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ contents: [{ parts: [{ text: testMessage }] }] })
            });
            const result = await readJSON(response);
            if (!response.ok) throw new Error(getAPIError(result, "فشل الاتصال بـ Gemini."));
            return "✓ تم الاتصال بـ Gemini بنجاح";
        }

        throw new Error("مزود الذكاء الاصطناعي غير معروف.");
    }

    if (testConnection) {
        testConnection.onclick = async function (e) {
            e.preventDefault();
            e.stopPropagation();
            testConnection.disabled = true;
            if (settingsStatus) settingsStatus.innerHTML = "⏳ جاري اختبار الاتصال...";
            try {
                const message = await testAIConnection();
                if (settingsStatus) settingsStatus.innerHTML = message;
            } catch (error) {
                if (settingsStatus) settingsStatus.innerHTML = "⚠ " + (error.message || "تعذر الاتصال");
            } finally {
                testConnection.disabled = false;
            }
        };
    }

    // =====================================================
    // STREAM AI FUNCTIONS
    // =====================================================

    async function streamGroqAI(text, onChunk) {
        const data = getSavedSettings();
        const key = data.key || "";
        const model = data.model || "";
        if (!key.trim()) throw new Error("لم يتم إدخال مفتاح Groq من الإعدادات.");
        if (!model.trim()) throw new Error("لم يتم تحديد نموذج Groq.");

        const documentContext = await buildAIDocumentContext(text);
        const conversationMessages = [];
        const historyLimit = documentContext && documentContext.found ? 2 : 4;

        if (currentChat && Array.isArray(currentChat.messages)) {
            const previousMessages = currentChat.messages.slice(-historyLimit);
            previousMessages.forEach(function (msg) {
                if (!msg || !msg.text) return;
                let messageText = String(msg.text).trim();
                const maxHistoryChars = documentContext && documentContext.found ? 1000 : 1500;
                if (messageText.length > maxHistoryChars) messageText = messageText.substring(0, maxHistoryChars) + "…";
                conversationMessages.push({ role: msg.role === "ai" ? "assistant" : "user", content: messageText });
            });
        }

        let userContent = text;
        if (documentContext && documentContext.found) {
            userContent = [
                "أنت تجيب عن سؤال مستخدم في أداة بحث أكاديمية.",
                "",
                "=== سؤال المستخدم ===",
                text,
                "",
                "=== بيانات المستند ===",
                "اسم المستند: " + (currentDocument ? currentDocument.name : ""),
                "",
                "=== المادة المستخرجة من المستند ===",
                documentContext.text,
                "",
                "=== قواعد الإجابة ===",
                "أجب عن سؤال المستخدم اعتمادًا على المادة المستخرجة من المستند بوصفها المصدر الأساسي.",
                "استخرج الأفكار المرتبطة بالسؤال فقط.",
                "ادمج الأفكار المتشابهة في فكرة واحدة ولا تكررها بصيغ مختلفة.",
                "رتب الإجابة وفق محاور السؤال.",
                "لا تضف معلومة أو حكمًا أو نسبة قول إلى المستند غير موجودة في المقاطع المستخرجة.",
                "إذا لم تكف المقاطع للإجابة عن جزء من السؤال، صرّح بذلك بوضوح.",
                "لا تستخدم المعرفة العامة لسد النقص في المستند إلا إذا طلب المستخدم ذلك صراحة.",
                "حافظ على العربية والأسلوب الأكاديمي.",
                "لا تبدأ باعتذار أو تمهيد عام غير ضروري.",
                "ضع الإحالات [مقطع X] بعد الأفكار التي يدعمها المستند.",
                "إذا تكررت الفكرة نفسها في أكثر من مقطع، اذكرها مرة واحدة واجمع الإحالات.",
                "لا تكرر الإحالة نفسها دون فائدة.",
                "قدّم إجابة كاملة ومترابطة بالقدر الذي يحتاجه السؤال."
            ].join("\n");
        }

        conversationMessages.push({ role: "user", content: userContent });

        const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
            method: "POST",
            headers: { "Content-Type": "application/json", "Authorization": "Bearer " + key },
            body: JSON.stringify({
                model: model,
                messages: conversationMessages,
                max_tokens: 2500,
                temperature: 0.2,
                stream: true
            })
        });

        if (!response.ok) {
            const result = await readJSON(response);
            throw new Error(getAPIError(result, "فشل الاتصال بـ Groq."));
        }
        if (!response.body) throw new Error("المتصفح لا يدعم استقبال الرد المتدفق من Groq.");

        const reader = response.body.getReader();
        const decoder = new TextDecoder("utf-8");
        let buffer = "";
        let fullAnswer = "";

        while (true) {
            const chunk = await reader.read();
            if (chunk.done) break;
            buffer += decoder.decode(chunk.value, { stream: true });
            const events = buffer.split("\n\n");
            buffer = events.pop() || "";

            for (let i = 0; i < events.length; i++) {
                const lines = events[i].split("\n");
                for (let j = 0; j < lines.length; j++) {
                    const line = lines[j].trim();
                    if (!line.startsWith("data:")) continue;
                    const dataText = line.substring(5).trim();
                    if (dataText === "[DONE]") continue;
                    try {
                        const parsed = JSON.parse(dataText);
                        const delta = parsed && parsed.choices && parsed.choices[0] && parsed.choices[0].delta ?
                            parsed.choices[0].delta.content : "";
                        if (typeof delta === "string" && delta !== "") {
                            fullAnswer += delta;
                            if (typeof onChunk === "function") onChunk(delta, fullAnswer);
                        }
                    } catch (error) { continue; }
                }
            }
        }

        return fullAnswer.trim();
    }

    async function streamGeminiAI(text, onChunk) {
        const data = getSavedSettings();
        const key = data.key || "";
        const model = data.model || "";
        if (!key.trim()) throw new Error("لم يتم إدخال مفتاح Gemini من الإعدادات.");
        if (!model.trim()) throw new Error("لم يتم تحديد نموذج Gemini من الإعدادات.");

        const documentContext = await buildAIDocumentContext(text);

        const systemInstruction = [
            "أنت مساعد بحث أكاديمي يعمل على مستندات Word.",
            "اعتمد على المادة المستخرجة من المستند بوصفها المصدر الأساسي للإجابة.",
            "أجب عن السؤال مباشرة وبأسلوب أكاديمي واضح.",
            "رتب الإجابة وفق محاور السؤال، ولا تخلط بين أجزائه.",
            "إذا كان السؤال يتضمن أكثر من جانب، فافصل بينها بعناوين أو فقرات واضحة.",
            "ادمج الأفكار المتشابهة في صياغة واحدة.",
            "لا تحول كل مقطع مستخرج إلى فقرة مستقلة؛ ابنِ إجابة تركيبية من المقاطع.",
            "استبعد المعلومة الجانبية التي لا تجيب مباشرة عن السؤال.",
            "إذا دعمت عدة مقاطع الفكرة نفسها، اجمع إحالاتها بعد الفكرة.",
            "لا تكرر الفكرة نفسها لمجرد ورودها في أكثر من مقطع.",
            "لا تضف معلومة أو حكمًا أو نسبة قول إلى المستند غير موجودة في المادة المستخرجة.",
            "إذا لم تكف المادة المستخرجة للإجابة عن جزء من السؤال، صرّح بذلك بوضوح.",
            "لا تستخدم المعرفة العامة لسد النقص في المستند إلا إذا طلب المستخدم ذلك صراحة.",
            "لا تذكر مشكلة الدراسة أو أهدافها أو منهجها أو أسئلتها إلا إذا طلب المستخدم ذلك صراحة.",
            "ضع الإحالات بعد الأفكار التي يدعمها المستند بصيغة [مقطع X].",
            "إذا كانت الإحالة تشمل أكثر من مقطع فاستخدم [مقطع X، مقطع Y].",
            "لا تخترع أرقام المقاطع.",
            "حافظ على لغة السؤال ولغة المستند.",
            "استخدم العناوين والقوائم باعتدال عندما تساعد على وضوح الإجابة.",
            "لا تبدأ باعتذار أو تمهيد غير ضروري.",
            "لا تعيد صياغة سؤال المستخدم في بداية الإجابة.",
            "قدّم خلاصة مترابطة ومباشرة، لا تلخيصًا منفصلًا لكل مقطع."
        ].join("\n");

        const conversationMessages = [];
        const historyLimit = documentContext && documentContext.found ? 2 : 4;

        if (currentChat && Array.isArray(currentChat.messages)) {
            const previousMessages = currentChat.messages.slice(-historyLimit);
            previousMessages.forEach(function (msg) {
                if (!msg || !msg.text) return;
                let messageText = String(msg.text).trim();
                const maxHistoryChars = documentContext && documentContext.found ? 1000 : 1500;
                if (messageText.length > maxHistoryChars) messageText = messageText.substring(0, maxHistoryChars) + "…";
                conversationMessages.push({
                    role: msg.role === "ai" ? "model" : "user",
                    parts: [{ text: messageText }]
                });
            });
        }

        let userContent = text;
        if (documentContext && documentContext.found) {
            userContent = [
                "=== سؤال المستخدم ===",
                text,
                "",
                "=== اسم المستند ===",
                (currentDocument ? currentDocument.name : ""),
                "",
                "=== المادة المستخرجة من المستند ===",
                documentContext.text
            ].join("\n");
        }

        conversationMessages.push({ role: "user", parts: [{ text: userContent }] });

        const cleanModel = normalizeGeminiModel(model);
        const url = "https://generativelanguage.googleapis.com/v1beta/models/" +
            encodeURIComponent(cleanModel) +
            ":streamGenerateContent?alt=sse";

        const response = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json", "x-goog-api-key": key },
            body: JSON.stringify({
                systemInstruction: { parts: [{ text: systemInstruction }] },
                contents: conversationMessages,
                generationConfig: { temperature: 0.2 }
            })
        });

        if (!response.ok) {
            const result = await readJSON(response);
            throw new Error(getAPIError(result, "فشل الاتصال بـ Gemini."));
        }
        if (!response.body) throw new Error("المتصفح لا يدعم استقبال الرد المتدفق من Gemini.");

        const reader = response.body.getReader();
        const decoder = new TextDecoder("utf-8");
        let buffer = "";
        let fullAnswer = "";

        function processSSELine(line) {
            const cleanLine = String(line || "").trim();
            if (!cleanLine || !cleanLine.startsWith("data:")) return;
            const dataText = cleanLine.substring(5).trim();
            if (!dataText || dataText === "[DONE]") return;

            try {
                const parsed = JSON.parse(dataText);
                if (!parsed || !Array.isArray(parsed.candidates) || !parsed.candidates[0]) return;
                const candidate = parsed.candidates[0];
                if (!candidate.content || !Array.isArray(candidate.content.parts)) return;

                candidate.content.parts.forEach(function (part) {
                    if (!part || typeof part.text !== "string" || part.thought === true) return;
                    const delta = part.text;
                    if (!delta) return;
                    fullAnswer += delta;
                    if (typeof onChunk === "function") onChunk(delta, fullAnswer);
                });
            } catch (error) { return; }
        }

        while (true) {
            const streamResult = await reader.read();
            if (streamResult.done) break;
            buffer += decoder.decode(streamResult.value, { stream: true });
            buffer = buffer.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

            let newlineIndex = buffer.indexOf("\n");
            while (newlineIndex !== -1) {
                const line = buffer.substring(0, newlineIndex);
                buffer = buffer.substring(newlineIndex + 1);
                processSSELine(line);
                newlineIndex = buffer.indexOf("\n");
            }
        }

        if (buffer.trim()) processSSELine(buffer);
        if (!fullAnswer.trim()) throw new Error("لم يصل نص من Gemini عبر البث المتدفق.");
        return fullAnswer.trim();
    }

    async function streamOpenRouterAI(text, onChunk) {
        const data = getSavedSettings();
        const key = data.key || "";
        const model = data.model || "";
        if (!key.trim()) throw new Error("لم يتم إدخال مفتاح OpenRouter من الإعدادات.");
        if (!model.trim()) throw new Error("لم يتم تحديد نموذج OpenRouter.");

        const documentContext = await buildAIDocumentContext(text);
        const conversationMessages = [];
        const historyLimit = documentContext && documentContext.found ? 2 : 4;

        if (currentChat && Array.isArray(currentChat.messages)) {
            const previousMessages = currentChat.messages.slice(-historyLimit);
            previousMessages.forEach(function (msg) {
                if (!msg || !msg.text) return;
                let messageText = String(msg.text).trim();
                const maxHistoryChars = documentContext && documentContext.found ? 1000 : 1500;
                if (messageText.length > maxHistoryChars) messageText = messageText.substring(0, maxHistoryChars) + "…";
                conversationMessages.push({ role: msg.role === "ai" ? "assistant" : "user", content: messageText });
            });
        }

        let userContent = text;
        if (documentContext && documentContext.found) {
            userContent = [
                "أنت تجيب عن سؤال مستخدم في أداة بحث أكاديمية.",
                "",
                "=== سؤال المستخدم ===",
                text,
                "",
                "=== اسم المستند ===",
                (currentDocument ? currentDocument.name : ""),
                "",
                "=== المادة المستخرجة من المستند ===",
                documentContext.text
            ].join("\n");
        }

        conversationMessages.push({ role: "user", content: userContent });

        const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + key,
                "HTTP-Referer": window.location.href,
                "X-Title": "Research Tools"
            },
            body: JSON.stringify({
                model: model,
                messages: conversationMessages,
                max_tokens: 2500,
                temperature: 0.2,
                stream: true
            })
        });

        if (!response.ok) {
            const result = await readJSON(response);
            throw new Error(getAPIError(result, "فشل الاتصال بـ OpenRouter."));
        }
        if (!response.body) throw new Error("المتصفح لا يدعم استقبال الرد المتدفق من OpenRouter.");

        const reader = response.body.getReader();
        const decoder = new TextDecoder("utf-8");
        let buffer = "";
        let fullAnswer = "";

        function processSSELine(line) {
            const cleanLine = String(line || "").trim();
            if (!cleanLine || !cleanLine.startsWith("data:")) return;
            const dataText = cleanLine.substring(5).trim();
            if (!dataText || dataText === "[DONE]") return;

            try {
                const parsed = JSON.parse(dataText);
                const delta = parsed && parsed.choices && parsed.choices[0] && parsed.choices[0].delta ?
                    parsed.choices[0].delta.content : "";
                if (typeof delta === "string" && delta) {
                    fullAnswer += delta;
                    if (typeof onChunk === "function") onChunk(delta, fullAnswer);
                }
            } catch (error) { return; }
        }

        while (true) {
            const streamResult = await reader.read();
            if (streamResult.done) break;
            buffer += decoder.decode(streamResult.value, { stream: true });
            buffer = buffer.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

            let newlineIndex = buffer.indexOf("\n");
            while (newlineIndex !== -1) {
                const line = buffer.substring(0, newlineIndex);
                buffer = buffer.substring(newlineIndex + 1);
                processSSELine(line);
                newlineIndex = buffer.indexOf("\n");
            }
        }

        if (buffer.trim()) processSSELine(buffer);
        if (!fullAnswer.trim()) throw new Error("لم يصل نص من OpenRouter عبر البث المتدفق.");
        return fullAnswer.trim();
    }

    async function streamOpenAI(text, onChunk) {
        const data = getSavedSettings();
        const key = data.key || "";
        const model = data.model || "";
        if (!key.trim()) throw new Error("لم يتم إدخال مفتاح OpenAI من الإعدادات.");
        if (!model.trim()) throw new Error("لم يتم تحديد نموذج OpenAI.");

        const documentContext = await buildAIDocumentContext(text);
        const conversationMessages = [];
        const historyLimit = documentContext && documentContext.found ? 2 : 4;

        if (currentChat && Array.isArray(currentChat.messages)) {
            const previousMessages = currentChat.messages.slice(-historyLimit);
            previousMessages.forEach(function (msg) {
                if (!msg || !msg.text) return;
                let messageText = String(msg.text).trim();
                const maxHistoryChars = documentContext && documentContext.found ? 1000 : 1500;
                if (messageText.length > maxHistoryChars) messageText = messageText.substring(0, maxHistoryChars) + "…";
                conversationMessages.push({ role: msg.role === "ai" ? "assistant" : "user", content: messageText });
            });
        }

        let userContent = text;
        if (documentContext && documentContext.found) {
            userContent = [
                "أنت تجيب عن سؤال مستخدم في أداة بحث أكاديمية.",
                "",
                "=== سؤال المستخدم ===",
                text,
                "",
                "=== اسم المستند ===",
                (currentDocument ? currentDocument.name : ""),
                "",
                "=== المادة المستخرجة من المستند ===",
                documentContext.text
            ].join("\n");
        }

        conversationMessages.push({ role: "user", content: userContent });

        const response = await fetch("https://api.openai.com/v1/chat/completions", {
            method: "POST",
            headers: { "Content-Type": "application/json", "Authorization": "Bearer " + key },
            body: JSON.stringify({
                model: model,
                messages: conversationMessages,
                max_tokens: 2500,
                temperature: 0.2,
                stream: true
            })
        });

        if (!response.ok) {
            const result = await readJSON(response);
            throw new Error(getAPIError(result, "فشل الاتصال بـ OpenAI."));
        }
        if (!response.body) throw new Error("المتصفح لا يدعم استقبال الرد المتدفق من OpenAI.");

        const reader = response.body.getReader();
        const decoder = new TextDecoder("utf-8");
        let buffer = "";
        let fullAnswer = "";

        function processSSELine(line) {
            const cleanLine = String(line || "").trim();
            if (!cleanLine || !cleanLine.startsWith("data:")) return;
            const dataText = cleanLine.substring(5).trim();
            if (!dataText || dataText === "[DONE]") return;

            try {
                const parsed = JSON.parse(dataText);
                const delta = parsed && parsed.choices && parsed.choices[0] && parsed.choices[0].delta ?
                    parsed.choices[0].delta.content : "";
                if (typeof delta === "string" && delta) {
                    fullAnswer += delta;
                    if (typeof onChunk === "function") onChunk(delta, fullAnswer);
                }
            } catch (error) { return; }
        }

        while (true) {
            const streamResult = await reader.read();
            if (streamResult.done) break;
            buffer += decoder.decode(streamResult.value, { stream: true });
            buffer = buffer.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

            let newlineIndex = buffer.indexOf("\n");
            while (newlineIndex !== -1) {
                const line = buffer.substring(0, newlineIndex);
                buffer = buffer.substring(newlineIndex + 1);
                processSSELine(line);
                newlineIndex = buffer.indexOf("\n");
            }
        }

        if (buffer.trim()) processSSELine(buffer);
        if (!fullAnswer.trim()) throw new Error("لم يصل نص من OpenAI عبر البث المتدفق.");
        return fullAnswer.trim();
    }

    // =====================================================
    // SEND MESSAGE
    // =====================================================

    async function sendMessage() {
        if (!input) return;
        const text = input.value.trim();
        if (!text) return;

        if (!currentChat) {
            currentChat = {
                id: Date.now(),
                title: text.substring(0, 30),
                messages: [],
                isTemporary: true,
                projectId: currentProject ? currentProject.id : null
            };
        }

        if (currentChat.isTemporary) {
            currentChat.isTemporary = false;
            currentChat.projectId = currentProject ? currentProject.id : null;
            currentChat.title = text.substring(0, 30);

            const alreadyExists = chats.some(function (chat) { return chat.id === currentChat.id; });
            if (!alreadyExists) chats.unshift(currentChat);

            if (currentProject && currentChat.projectId === currentProject.id) {
                if (!Array.isArray(currentProject.chatIds)) currentProject.chatIds = [];
                if (!currentProject.chatIds.includes(currentChat.id)) {
                    currentProject.chatIds.push(currentChat.id);
                    currentProject.updatedAt = new Date().toISOString();
                    saveProjects();
                }
            }
            saveChats();
        }

        currentChat.messages.push({ role: "user", text: text });
        saveChats();
        renderChat();
        renderChatList();
        renderSidebarChats();
        renderRecentChats();

        input.value = "";
        input.style.height = "auto";

        const loading = document.createElement("div");
        loading.className = "message ai-message";
        loading.innerHTML = "⏳ جاري التفكير...";
        if (chatArea) { chatArea.appendChild(loading); chatArea.scrollTop = chatArea.scrollHeight; }

        const savedSettings = getSavedSettings();
        const selectedProvider = String(savedSettings.provider || "openrouter").toLowerCase();

        let pendingRenderText = "";
        let renderTimer = null;

        function renderStreamingText() {
            if (!loading) return;
            if (pendingRenderText === "") {
                loading.innerHTML = "⏳ جاري التفكير...";
            } else {
                loading.innerHTML = formatAIMessage(pendingRenderText, currentCitationSources);
            }
            if (chatArea) chatArea.scrollTop = chatArea.scrollHeight;
            renderTimer = null;
        }

        function scheduleRender(fullText) {
            pendingRenderText = String(fullText || "");
            if (renderTimer !== null) return;
            renderTimer = setTimeout(function () { renderStreamingText(); }, 60);
        }

        try {
            let answer = "";

            if (selectedProvider === "groq") {
                answer = await streamGroqAI(text, function (delta, fullText) { scheduleRender(fullText); });
            } else if (selectedProvider === "gemini") {
                answer = await streamGeminiAI(text, function (delta, fullText) { scheduleRender(fullText); });
            } else if (selectedProvider === "openrouter") {
                answer = await streamOpenRouterAI(text, function (delta, fullText) { scheduleRender(fullText); });
            } else if (selectedProvider === "openai") {
                answer = await streamOpenAI(text, function (delta, fullText) { scheduleRender(fullText); });
            } else {
                throw new Error("مزود الذكاء الاصطناعي غير معروف: " + selectedProvider);
            }

            if (renderTimer !== null) { clearTimeout(renderTimer); renderTimer = null; }
            pendingRenderText = String(answer || "");
            renderStreamingText();

            if (loading && loading.parentNode) loading.remove();

            currentChat.messages.push({
                role: "ai",
                text: String(answer || ""),
                citationSources: Array.isArray(currentCitationSources) ?
                    currentCitationSources.map(function (source) {
                        return {
                            rank: source.rank,
                            paragraphIndex: source.paragraphIndex,
                            heading: source.heading || "",
                            text: source.text || ""
                        };
                    }) : []
            });

            saveChats();
            renderChat();
            renderChatList();
            renderSidebarChats();
            renderRecentChats();
        } catch (error) {
            if (renderTimer !== null) { clearTimeout(renderTimer); renderTimer = null; }
            if (loading && loading.parentNode) loading.remove();

            currentChat.messages.push({
                role: "ai",
                text: "خطأ: " + (error && error.message ? error.message : "حدث خطأ غير معروف")
            });

            saveChats();
            renderChat();
            renderChatList();
            renderSidebarChats();
            renderRecentChats();
        }
    }

    // =====================================================
    // SEND BUTTON & KEYBOARD
    // =====================================================

    if (sendBtn) {
        sendBtn.onclick = function (e) { e.preventDefault(); sendMessage(); };
    }

    if (input) {
        input.onkeydown = function (e) {
            if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                sendMessage();
            }
        };
    }

    // =====================================================
    // SEARCH POPUP
    // =====================================================

    if (searchBtn) {
        searchBtn.onclick = function (e) {
            e.stopPropagation();
            if (!searchPopup) return;
            if (projectsPopup) projectsPopup.classList.remove("open");
            if (chatPopup) chatPopup.classList.remove("open");
            if (settingsWindow) settingsWindow.classList.remove("open");
            searchPopup.classList.toggle("open");
            if (searchPopup.classList.contains("open") && searchInput) searchInput.focus();
        };
    }

    if (searchInput) {
        searchInput.oninput = function () {
            const keyword = searchInput.value.trim().toLowerCase();
            if (!searchResults) return;
            searchResults.innerHTML = "";
            if (keyword === "") return;

            chats.forEach(function (chat) {
                if (chat.title.toLowerCase().includes(keyword)) {
                    const item = document.createElement("div");
                    item.className = "search-result-item";
                    item.innerHTML = `<span class="chat-title">${chatIcon} ${chat.title}</span>`;
                    item.onclick = function () {
                        currentChat = chat;
                        renderChat();
                        if (searchPopup) searchPopup.classList.remove("open");
                        searchInput.value = "";
                        searchResults.innerHTML = "";
                    };
                    searchResults.appendChild(item);
                }
            });
        };
    }

    // =====================================================
    // CITATION CLICK HANDLER
    // =====================================================

    if (chatArea) {
        chatArea.addEventListener("click", function (event) {
            const citation = event.target.closest(".document-citation");
            if (!citation) return;
            event.preventDefault();
            event.stopPropagation();
            const rank = Number(citation.getAttribute("data-citation-rank"));
            if (!Number.isNaN(rank)) openCitationInWord(rank);
        });
    }

    // =====================================================
    // SIDEBAR PIN
    // =====================================================

    const sidebar = document.querySelector(".sidebar");
    const pinSidebar = document.getElementById("pin-sidebar");

    if (pinSidebar && sidebar) {
        let sidebarPinned = localStorage.getItem("sidebarPinned") === "true";
        if (sidebarPinned) {
            sidebar.classList.add("pinned");
            pinSidebar.classList.add("pinned");
            document.body.classList.add("sidebar-is-pinned");
        }

        pinSidebar.addEventListener("click", function (e) {
            e.stopPropagation();
            sidebarPinned = !sidebarPinned;
            sidebar.classList.toggle("pinned", sidebarPinned);
            pinSidebar.classList.toggle("pinned", sidebarPinned);
            document.body.classList.toggle("sidebar-is-pinned", sidebarPinned);
            localStorage.setItem("sidebarPinned", sidebarPinned ? "true" : "false");
        });
    }

    // =====================================================
    // INITIALIZATION
    // =====================================================

    initializeSidebarSections();
    renderProjects();
    renderExpandedProjects();
    renderDocuments();
    renderChatList();
    renderSidebarChats();
    renderRecentChats();
    renderChat();
    loadSettings();

    console.log("Word AI Assistant initialized successfully.");
    console.log("Orama search engine ready.");
    console.log("Supported providers: OpenRouter, Gemini, OpenAI, Groq");

    // =====================================================
    // END Office.onReady
    // =====================================================
});