/*
 * Word AI Assistant - Orama Version
 * ======================================
 * تم استبدال محرك البحث اليدوي بـ Orama
 * مع الحفاظ على جميع وظائف الواجهة
 */

Office.onReady(function () {

    // ======================================
    // Elements
    // ======================================

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

    const chatPopup = document.getElementById("chat-popup");
    const recentChatList = document.getElementById("recent-chat-list");

    const searchPopup = document.getElementById("search-popup");
    const searchInput = document.getElementById("search-input");
    const searchResults = document.getElementById("search-results");
    const searchBtn = document.getElementById("search-btn");

    let sidebarTogglePlaceholder = document.createComment("sidebar-toggle-placeholder");
    if (sidebarToggleBtn && sidebarToggleBtn.parentNode) {
        sidebarToggleBtn.parentNode.insertBefore(sidebarTogglePlaceholder, sidebarToggleBtn);
    }

    // ======================================
    // State
    // ======================================

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

    // ======================================
    // Constants
    // ======================================

    const DOCUMENT_DB_NAME = "WORD_AI_DOCUMENT_STORAGE";
    const DOCUMENT_DB_VERSION = 4;
    const DOCUMENT_STORE_NAME = "files";
    const DOCUMENT_TEXT_STORE_NAME = "texts";
    const DOCUMENT_STRUCTURE_STORE_NAME = "structures";
    const ORAMA_SCHEMA_VERSION = 1;

    // ======================================
    // Icons
    // ======================================

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

    // ======================================
    // Storage Helpers
    // ======================================

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

    // ======================================
    // Storage Functions
    // ======================================

    function saveProjects() {
        localStorage.setItem("WORD_AI_PROJECTS", JSON.stringify(projects));
    }

    function saveDocuments() {
        localStorage.setItem("WORD_AI_DOCUMENTS", JSON.stringify(documents));
    }

    function saveChats() {
        localStorage.setItem("WORD_AI_CHATS", JSON.stringify(chats));
    }

    // ======================================
    // Load Data
    // ======================================

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

    documents = readStorageArray("WORD_AI_DOCUMENTS")
        .filter(function (item) { return item && typeof item === "object"; })
        .map(function (item) {
            return {
                ...item,
                indexTokenCount: Number(item.indexTokenCount || 0),
                indexUniqueTerms: Number(item.indexUniqueTerms || 0),
                indexSchemaVersion: Number(item.indexSchemaVersion || 0),
                indexStatus: item.indexStatus || "new",
                readStatus: item.readStatus || "new"
            };
        });

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

    // ======================================
    // IndexedDB Functions
    // ======================================

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

    // ======================================
    // Working File Operations
    // ======================================

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

    // ======================================
    // Document Text Operations
    // ======================================

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

    // ======================================
    // Document Structure Operations
    // ======================================

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

    // ======================================
    // File Helpers
    // ======================================

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

    // ======================================
    // Text Normalization
    // ======================================

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

    // ======================================
    // Orama Loader
    // ======================================

    function loadOrama() {
        try {
            return require("@orama/orama");
        } catch (error) {
            throw new Error("تعذر تحميل Orama: " + error.message);
        }
    }

    // ======================================
    // Heading Level Number
    // ======================================

    function getHeadingLevelNumber(style) {
        const match = String(style || "").match(/Heading\s*([1-9])/i);
        return match ? Number(match[1]) : 9;
    }

    // ======================================
    // Build Orama Retrieval Index
    // ======================================

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

        // تحديث إحصائيات المستند
        documentItem.indexStatus = "indexed";
        documentItem.indexTokenCount = records.reduce(function (total, record) {
            return total + (record.text.match(/[\p{L}\p{N}]+/gu) || []).length;
        }, 0);
        documentItem.indexUniqueTerms = new Set(records.flatMap(function (record) {
            return record.text.match(/[\p{L}\p{N}]+/gu) || [];
        })).size;
        documentItem.indexSchemaVersion = ORAMA_SCHEMA_VERSION;
        documentItem.indexUpdatedAt = new Date().toISOString();
        saveDocuments();

        console.log("تم بناء فهرس Orama:", {
            documentId: documentItem.id,
            records: records.length,
            headings: headings.length
        });

        return db;
    }

    // ======================================
    // Search with Orama
    // ======================================

    async function searchOramaDocument(documentItem, query, limit) {
        if (!documentItem || !query) return [];

        const cleanQuery = String(query || "").trim();
        if (!cleanQuery) return [];

        if (!oramaRetrievalDb || oramaRetrievalDocumentId !== String(documentItem.id)) {
            const structureData = await ensureDocumentStructure(documentItem);
            await buildOramaRetrievalIndex(documentItem, structureData);
        }

        const db = oramaRetrievalDb;
        if (!db) return [];

        const { search } = loadOrama();

        const result = await search(db, {
            term: cleanQuery,
            properties: ["text", "heading"],
            limit: Math.max(1, Number(limit || 10)),
            tolerance: 1,
            boost: { heading: 5 }
        });

        if (!result || !Array.isArray(result.hits)) return [];

        return result.hits.map(function (hit, index) {
            const doc = hit.document || {};
            return {
                rank: index + 1,
                paragraphIndex: Number(doc.paragraphIndex || 0),
                paragraphId: doc.id || "",
                heading: String(doc.heading || ""),
                headingIndex: Number(doc.headingIndex || -1),
                headingLevel: String(doc.headingLevel || ""),
                isHeading: Boolean(doc.isHeading),
                text: String(doc.text || ""),
                score: Number(hit.score || 0)
            };
        });
    }

    // ======================================
    // Document Structure Building
    // ======================================

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

    // ======================================
    // Document Status Updates
    // ======================================

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

    // ======================================
    // Project Documents Helpers
    // ======================================

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
            ensureDocumentStructure(documentItem)
                .then(function (structureData) {
                    return buildOramaRetrievalIndex(documentItem, structureData);
                })
                .then(function () { renderDocuments(); })
                .catch(function (error) {
                    console.error("تعذر بناء فهرس Orama:", error);
                    renderDocuments();
                });
            renderDocuments();
            return;
        }

        readCurrentWordDocument(documentItem)
            .then(function (text) {
                console.log("تم قراءة المستند:", text ? text.length : 0, "حرف");
                renderDocuments();
            })
            .catch(function (error) {
                console.error("تعذر قراءة المستند:", error);
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
            indexSchemaVersion: 0,
            createdAt: now,
            updatedAt: now
        };
        documents.push(item);
        saveDocuments();
        return item;
    }

    // ======================================
    // Read Current Word Document
    // ======================================

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
            const structureData = await buildDocumentStructure(documentItem);
            await saveDocumentStructure(documentItem.id, structureData);
            await buildOramaRetrievalIndex(documentItem, structureData);
            updateDocumentIndexStatus(documentItem, "indexed");

            console.log("تمت قراءة المستند وفهرسته باستخدام Orama بنجاح:", {
                documentId: documentItem.id,
                paragraphs: structureData.paragraphCount,
                headings: structureData.headingCount
            });

            return text;
        } catch (error) {
            updateDocumentReadStatus(documentItem, "error");
            updateDocumentIndexStatus(documentItem, "error");
            console.error("فشل قراءة/فهرسة المستند:", error);
            throw error;
        }
    }

    // ======================================
    // Retrieval Profile and Context Building
    // ======================================

    function getRetrievalProfile(query) {
        const text = normalizeSearchText(query);
        const profile = { type: "general", maxResults: 8, maxChars: 8000 };

        if (/ماهو|ماهو|ماهى|ماهي|ما هي|المقصود|معنى|تعريف|يقصد ب|المراد ب/.test(text)) {
            profile.type = "definition";
            profile.maxResults = 5;
            profile.maxChars = 6000;
        } else if (/اثر|أثر|تاثير|تأثير|نتائج|ينتج عن|يترتب على|انعكاس/.test(text)) {
            profile.type = "effect";
            profile.maxResults = 8;
            profile.maxChars = 9000;
        } else if (/الفرق|الفروق|مقارنة|يقارن|ما الفرق|التمييز بين|يفترق/.test(text)) {
            profile.type = "comparison";
            profile.maxResults = 10;
            profile.maxChars = 10000;
        } else if (/لماذا|سبب|اسباب|أسباب|علة|علل|لأن|لان|بسبب/.test(text)) {
            profile.type = "causes";
            profile.maxResults = 8;
            profile.maxChars = 9000;
        } else if (/اين|أين|موضع|موضعه|الفصل|المبحث|المطلب|الصفحة/.test(text)) {
            profile.type = "location";
            profile.maxResults = 6;
            profile.maxChars = 6000;
        }
        return profile;
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

    function buildRetrievalContext(searchResults, options) {
        const settings = options || {};
        const maxResults = typeof settings.maxResults === "number" ? settings.maxResults : 4;
        const maxChars = typeof settings.maxChars === "number" ? settings.maxChars : 3500;

        if (!Array.isArray(searchResults) || searchResults.length === 0) {
            return { contexts: [], text: "", selectedCount: 0 };
        }

        const sorted = searchResults.slice().sort(function (a, b) { return (b.score || 0) - (a.score || 0); });

        const selected = [];
        const selectedParagraphIndexes = new Set();
        const MAX_TEXT_OVERLAP = 0.75;

        for (let i = 0; i < sorted.length; i++) {
            const candidate = sorted[i];
            if (selectedParagraphIndexes.has(candidate.paragraphIndex)) continue;
            const candidateText = String(candidate.text || "").replace(/\s+/g, " ").trim();
            if (!candidateText) continue;

            let tooSimilar = false;
            for (let j = 0; j < selected.length; j++) {
                const selectedText = String(selected[j].text || "").replace(/\s+/g, " ").trim();
                if (!selectedText) continue;
                const shorter = Math.min(candidateText.length, selectedText.length);
                const common = getCommonTextLength(candidateText, selectedText);
                const overlap = shorter > 0 ? common / shorter : 0;
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
            let mainText = String(result.text || "").replace(/\s+/g, " ").trim();
            if (!mainText) return;

            const heading = result.heading || "";
            const remaining = maxChars - totalChars;
            if (remaining <= 0) return;
            const available = Math.max(300, remaining - 250);

            let context = mainText;
            if (context.length > available) {
                context = context.substring(0, available) + "…";
            }

            contexts.push({
                rank: index + 1,
                paragraphIndex: result.paragraphIndex,
                heading: heading,
                score: result.score || 0,
                mainParagraph: mainText,
                context: context
            });

            totalChars += context.length;
        });

        const textParts = contexts.map(function (item) {
            let block = "[مقطع " + item.rank + "]\n";
            if (item.heading) block += "العنوان: " + item.heading + "\n";
            block += "المقطع: " + item.mainParagraph;
            return block;
        });

        return {
            contexts: contexts,
            text: textParts.join("\n\n---\n\n"),
            selectedCount: contexts.length
        };
    }

    // ======================================
    // Build AI Document Context
    // ======================================

    async function buildAIDocumentContext(question) {
        if (!currentDocument) {
            currentCitationSources = [];
            return {
                found: false,
                query: String(question || ""),
                text: "",
                sources: [],
                resultCount: 0,
                selectedCount: 0
            };
        }

        const cleanQuestion = String(question || "").trim();
        if (!cleanQuestion) {
            currentCitationSources = [];
            return {
                found: false,
                query: "",
                text: "",
                sources: [],
                resultCount: 0,
                selectedCount: 0
            };
        }

        const retrievalProfile = getRetrievalProfile(cleanQuestion);
        const maxResults = retrievalProfile.maxResults || 8;

        const results = await searchOramaDocument(currentDocument, cleanQuestion, maxResults * 2);

        if (!results || results.length === 0) {
            currentCitationSources = [];
            return {
                found: false,
                query: cleanQuestion,
                profile: retrievalProfile.type,
                text: "",
                sources: [],
                resultCount: 0,
                selectedCount: 0
            };
        }

        const retrieval = buildRetrievalContext(results, {
            maxResults: maxResults,
            maxChars: 8000
        });

        currentCitationSources = retrieval.contexts.map(function (item) {
            return {
                rank: item.rank,
                paragraphIndex: item.paragraphIndex,
                heading: item.heading || "",
                mainParagraph: item.mainParagraph || "",
                text: item.context || item.mainParagraph || ""
            };
        });

        return {
            found: retrieval.selectedCount > 0,
            query: cleanQuestion,
            profile: retrievalProfile.type,
            text: retrieval.text,
            sources: currentCitationSources,
            resultCount: results.length,
            selectedCount: retrieval.selectedCount
        };
    }

    // ======================================
    // Format AI Message
    // ======================================

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

    // ======================================
    // Open Citation In Word
    // ======================================

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

    // ======================================
    // Render Functions (UI)
    // ======================================

    // ---- Render Documents ----
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
            if (currentDocument && currentDocument.id === documentItem.id) {
                item.classList.add("active-document");
            }

            const title = document.createElement("span");
            title.className = "document-title";
            title.textContent = documentItem.name;

            const status = document.createElement("span");
            status.className = "document-read-status";
            if (documentItem.indexStatus === "indexed") {
                status.textContent = "✓ مفهرس · " + documentItem.indexTokenCount + " كلمة · " + documentItem.indexUniqueTerms + " فريدة";
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

            // Rename
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
                    const prev = projectDocuments[index - 1];
                    const tmp = documentItem.order;
                    documentItem.order = prev.order;
                    prev.order = tmp;
                    documentItem.updatedAt = new Date().toISOString();
                    prev.updatedAt = new Date().toISOString();
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
                    const next = projectDocuments[index + 1];
                    const tmp = documentItem.order;
                    documentItem.order = next.order;
                    next.order = tmp;
                    documentItem.updatedAt = new Date().toISOString();
                    next.updatedAt = new Date().toISOString();
                    saveDocuments();
                    options.classList.remove("open");
                    renderDocuments();
                };
            }

            // Delete
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
                                return doc.id !== documentItem.id;
                            });

                            if (currentProject && Array.isArray(currentProject.documents)) {
                                currentProject.documents = currentProject.documents.filter(function (id) {
                                    return id !== documentItem.id;
                                });
                                currentProject.updatedAt = new Date().toISOString();
                                saveProjects();
                            }

                            try {
                                await deleteWorkingWordFile(documentItem.storageId);
                            } catch (storageError) {
                                console.warn("تعذر حذف نسخة العمل:", storageError);
                            }

                            if (currentDocument && currentDocument.id === documentItem.id) {
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

    // ---- Render Projects ----
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
                    document.querySelectorAll(".project-options-menu.open").forEach(function (m) {
                        m.classList.remove("open");
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
                                const doc = documents.find(function (d) { return d.id === projectDocumentIds[i]; });
                                if (doc) {
                                    try {
                                        await deleteWorkingWordFile(doc.storageId);
                                    } catch (error) {
                                        console.warn("تعذر حذف نسخة العمل:", error);
                                    }
                                }
                            }
                            documents = documents.filter(function (doc) {
                                return !projectDocumentIds.includes(doc.id);
                            });
                            projects = projects.filter(function (p) { return p.id !== project.id; });
                            if (currentProject && currentProject.id === project.id) {
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

    // ---- Render Expanded Projects ----
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

    // ---- Render Sidebar Chats ----
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

    // ---- Render Chat ----
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

    // ---- Render Chat List ----
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

    // ---- Render Recent Chats ----
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

    // ======================================
    // Sidebar Sections Initialization
    // ======================================

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

    // ======================================
    // Chat Management
    // ======================================

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

    // ======================================
    // AI Stream Functions
    // ======================================

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

    // ======================================
    // Send Message
    // ======================================

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

    // ======================================
    // Buttons & Event Handlers
    // ======================================

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

    // ======================================
    // Settings Handlers
    // ======================================

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

    // ======================================
    // Load Models (Settings)
    // ======================================

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

    // ======================================
    // Test Connection
    // ======================================

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

    // ======================================
    // Search Popup (البحث في المحادثات فقط)
    // ======================================

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

    // ======================================
    // Citation Click Handler
    // ======================================

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

    // ======================================
    // Sidebar Pin
    // ======================================

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

    // ======================================
    // Initialization
    // ======================================

    initializeSidebarSections();
    renderProjects();
    renderExpandedProjects();
    renderDocuments();
    renderChatList();
    renderSidebarChats();
    renderRecentChats();
    renderChat();
    loadSettings();

    console.log("Word AI Assistant - Orama Version initialized successfully.");
    console.log("Orama search engine ready.");
    console.log("Supported providers: OpenRouter, Gemini, OpenAI, Groq");

    // ======================================
    // End of Office.onReady
    // ======================================
});