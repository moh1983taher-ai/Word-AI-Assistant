/* =========================================================
 * Research Memory v1
 * Persistent research-result store and AI knowledge-map store.
 * Kept separate from the existing chat database.
 * ======================================================= */
(function (global) {
    "use strict";

    if (global.ResearchMemory) return;

    const DB_NAME = "ResearchMemoryDB";
    const DB_VERSION = 1;
    const STORES = {
        results: "results",
        nodes: "nodes",
        edges: "edges",
        meta: "meta"
    };

    let dbPromise = null;

    function openDB() {
        if (dbPromise) return dbPromise;

        dbPromise = new Promise(function (resolve, reject) {
            if (!global.indexedDB) {
                reject(new Error("IndexedDB غير متاح في هذه البيئة."));
                return;
            }

            const request = global.indexedDB.open(DB_NAME, DB_VERSION);

            request.onupgradeneeded = function () {
                const db = request.result;

                if (!db.objectStoreNames.contains(STORES.results)) {
                    const store = db.createObjectStore(STORES.results, { keyPath: "id" });
                    store.createIndex("scopeId", "scopeId", { unique: false });
                    store.createIndex("fingerprint", "fingerprint", { unique: false });
                }

                if (!db.objectStoreNames.contains(STORES.nodes)) {
                    const store = db.createObjectStore(STORES.nodes, { keyPath: "id" });
                    store.createIndex("scopeId", "scopeId", { unique: false });
                    store.createIndex("nodeKey", "nodeKey", { unique: false });
                }

                if (!db.objectStoreNames.contains(STORES.edges)) {
                    const store = db.createObjectStore(STORES.edges, { keyPath: "id" });
                    store.createIndex("scopeId", "scopeId", { unique: false });
                    store.createIndex("edgeKey", "edgeKey", { unique: false });
                }

                if (!db.objectStoreNames.contains(STORES.meta)) {
                    db.createObjectStore(STORES.meta, { keyPath: "scopeId" });
                }
            };

            request.onsuccess = function () {
                const db = request.result;
                db.onversionchange = function () {
                    db.close();
                    dbPromise = null;
                };
                resolve(db);
            };

            request.onerror = function () {
                dbPromise = null;
                reject(request.error || new Error("تعذر فتح قاعدة الذاكرة البحثية."));
            };

            request.onblocked = function () {
                console.warn("[ResearchMemory] فتح قاعدة البيانات محجوب بسبب اتصال قديم.");
            };
        });

        return dbPromise;
    }

    function requestResult(request) {
        return new Promise(function (resolve, reject) {
            request.onsuccess = function () { resolve(request.result); };
            request.onerror = function () { reject(request.error); };
        });
    }

    function transactionDone(transaction) {
        return new Promise(function (resolve, reject) {
            transaction.oncomplete = function () { resolve(); };
            transaction.onerror = function () { reject(transaction.error); };
            transaction.onabort = function () {
                reject(transaction.error || new Error("أُلغيت معاملة الذاكرة البحثية."));
            };
        });
    }

    function safeString(value) {
        return value === null || value === undefined ? "" : String(value);
    }

    function normalizeArabic(value) {
        return safeString(value)
            .toLowerCase()
            .replace(/[\u064B-\u065F\u0670]/g, "")
            .replace(/[إأآٱ]/g, "ا")
            .replace(/ى/g, "ي")
            .replace(/[ؤئ]/g, function (letter) { return letter === "ؤ" ? "و" : "ي"; })
            .replace(/ة/g, "ه")
            .replace(/ـ/g, "")
            .replace(/[^\p{L}\p{N}\s]/gu, " ")
            .replace(/\s+/g, " ")
            .trim();
    }

    function hashText(value) {
        const text = safeString(value);
        let h1 = 2166136261;
        let h2 = 2246822519;

        for (let i = 0; i < text.length; i += 1) {
            const code = text.charCodeAt(i);
            h1 = Math.imul(h1 ^ code, 16777619);
            h2 = Math.imul(h2 ^ code, 3266489917);
        }

        return (h1 >>> 0).toString(36) + (h2 >>> 0).toString(36) + "_" + text.length.toString(36);
    }

    function cloneJSON(value) {
        try {
            return JSON.parse(JSON.stringify(value));
        } catch (error) {
            return null;
        }
    }

    function makeScopeId(scope) {
        if (typeof scope === "string" && scope.trim()) return scope.trim();

        if (!scope || typeof scope !== "object") {
            throw new Error("يجب تحديد نطاق واضح للذاكرة البحثية.");
        }

        const type = safeString(scope.type || "chat");
        const id = safeString(scope.id || scope.chatId || scope.projectId);

        if (!id) throw new Error("معرف نطاق الذاكرة البحثية غير موجود.");
        return type + ":" + id;
    }

    function normalizeResult(result) {
        const raw = cloneJSON(result);
        if (!raw || typeof raw !== "object") return null;

        const normalized = {
            source: safeString(raw.source),
            bookId: safeString(raw.bookId),
            title: safeString(raw.title),
            author: safeString(raw.author),
            sectionTitle: safeString(raw.sectionTitle),
            part: safeString(raw.part),
            page: safeString(raw.page),
            pageId: safeString(raw.pageId),
            text: safeString(raw.text)
        };

        const fingerprint = [
            normalized.source,
            normalized.bookId,
            normalized.title,
            normalized.author,
            normalized.sectionTitle,
            normalized.part,
            normalized.page,
            normalized.pageId,
            normalized.text
        ].join("\u001f");

        return {
            normalized: normalized,
            raw: raw,
            fingerprint: hashText(fingerprint)
        };
    }

    async function saveResults(scope, results, metadata) {
        const scopeId = makeScopeId(scope);
        const input = Array.isArray(results) ? results : [];
        if (!input.length) return { saved: 0, skipped: 0, total: 0 };

        const db = await openDB();
        const transaction = db.transaction(STORES.results, "readwrite");
        const store = transaction.objectStore(STORES.results);
        const now = new Date().toISOString();
        let saved = 0;
        let skipped = 0;

        input.forEach(function (item) {
            const prepared = normalizeResult(item);
            if (!prepared || !prepared.normalized.text.trim()) {
                skipped += 1;
                return;
            }

            const id = scopeId + "::" + prepared.fingerprint;
            store.put({
                id: id,
                scopeId: scopeId,
                fingerprint: prepared.fingerprint,
                source: prepared.normalized.source,
                bookId: prepared.normalized.bookId,
                title: prepared.normalized.title,
                author: prepared.normalized.author,
                sectionTitle: prepared.normalized.sectionTitle,
                part: prepared.normalized.part,
                page: prepared.normalized.page,
                pageId: prepared.normalized.pageId,
                text: prepared.normalized.text,
                raw: prepared.raw,
                metadata: cloneJSON(metadata || {}),
                createdAt: now,
                updatedAt: now
            });
            saved += 1;
        });

        await transactionDone(transaction);
        return { saved: saved, skipped: skipped, total: input.length };
    }

    async function getResults(scope, options) {
        const scopeId = makeScopeId(scope);
        const db = await openDB();
        const transaction = db.transaction(STORES.results, "readonly");
        const request = transaction.objectStore(STORES.results).index("scopeId").getAll(scopeId);
        const rows = await requestResult(request);
        const opts = options || {};
        const ids = Array.isArray(opts.ids) ? new Set(opts.ids) : null;
        let output = rows.filter(function (row) { return !ids || ids.has(row.id); });

        if (Number.isFinite(opts.limit) && opts.limit > 0) {
            output = output.slice(0, opts.limit);
        }

        return output;
    }

    async function search(scope, query, limit) {
        const scopeId = makeScopeId(scope);
        const rows = await getResults(scopeId);
        const terms = normalizeArabic(query).split(" ").filter(function (term) {
            return term.length > 1;
        });

        if (!terms.length) return [];

        // توسيع الاسترجاع بالمفاهيم والمرادفات والعلاقات التي اعتمدتها الخريطة.
        let knowledgeMap = { nodes: [], edges: [] };

        try {
            knowledgeMap = await getKnowledgeMap(scopeId);
        } catch (error) {
            console.warn("[ResearchMemory] Knowledge map lookup failed:", error);
        }

        const semanticEvidence = new Map();

        (knowledgeMap.nodes || []).forEach(function (node) {
            const conceptText = normalizeArabic([
                node.label,
                node.description,
                ...(Array.isArray(node.aliases) ? node.aliases : [])
            ].join(" "));

            const matched = terms.filter(function (term) {
                return conceptText.includes(term);
            });

            if (!matched.length) return;

            (node.resultIds || []).forEach(function (id) {
                const key = safeString(id);
                const previous = semanticEvidence.get(key) || {
                    score: 0,
                    concepts: []
                };

                previous.score += 5 + matched.length * 3;
                previous.concepts.push(node.label);
                semanticEvidence.set(key, previous);
            });
        });

        (knowledgeMap.edges || []).forEach(function (edge) {
            const relationText = normalizeArabic(edge.relation);
            const matched = terms.filter(function (term) {
                return relationText.includes(term);
            });

            if (!matched.length) return;

            (edge.evidenceResultIds || []).forEach(function (id) {
                const key = safeString(id);
                const previous = semanticEvidence.get(key) || {
                    score: 0,
                    concepts: []
                };

                previous.score += 4 + matched.length * 2;
                semanticEvidence.set(key, previous);
            });
        });

        const ranked = rows.map(function (row) {
            const title = normalizeArabic(row.title + " " + row.sectionTitle);
            const author = normalizeArabic(row.author);
            const source = normalizeArabic(row.source);
            const text = normalizeArabic(row.text);
            let lexicalScore = 0;
            let matched = 0;

            terms.forEach(function (term) {
                let termScore = 0;
                if (title.includes(term)) termScore += 7;
                if (author.includes(term)) termScore += 3;
                if (source.includes(term)) termScore += 1;
                if (text.includes(term)) termScore += 2;
                if (termScore) matched += 1;
                lexicalScore += termScore;
            });

            const semantic = semanticEvidence.get(row.id) || {
                score: 0,
                concepts: []
            };

            return {
                row: row,
                score: lexicalScore + semantic.score,
                lexicalScore: lexicalScore,
                semanticScore: semantic.score,
                matchedTerms: matched,
                matchedConcepts: Array.from(new Set(semantic.concepts))
            };
        }).filter(function (item) {
            return item.score > 0;
        });

        ranked.sort(function (a, b) {
            return b.score - a.score ||
                b.matchedTerms - a.matchedTerms ||
                b.semanticScore - a.semanticScore;
        });

        const max = Number.isFinite(limit) && limit > 0 ? Math.floor(limit) : 30;
        return ranked.slice(0, max).map(function (item) {
            return Object.assign({}, item.row, {
                retrievalScore: item.score,
                lexicalScore: item.lexicalScore,
                semanticScore: item.semanticScore,
                matchedTerms: item.matchedTerms,
                matchedConcepts: item.matchedConcepts
            });
        });
    }

    function normalizeMapNode(scopeId, node) {
        if (!node || typeof node !== "object") return null;
        const label = safeString(node.label || node.name || node.title).trim();
        if (!label) return null;
        const nodeKey = safeString(node.key || node.id || normalizeArabic(label)).trim();
        const id = scopeId + "::node::" + hashText(nodeKey);

        return {
            id: id,
            scopeId: scopeId,
            nodeKey: nodeKey,
            label: label,
            description: safeString(node.description),
            type: safeString(node.type || "concept"),
            aliases: Array.isArray(node.aliases) ? node.aliases.map(safeString) : [],
            resultIds: Array.isArray(node.resultIds) ? node.resultIds.map(safeString) : [],
            updatedAt: new Date().toISOString()
        };
    }

    function normalizeMapEdge(scopeId, edge) {
        if (!edge || typeof edge !== "object") return null;
        const from = safeString(edge.from || edge.source).trim();
        const to = safeString(edge.to || edge.target).trim();
        const relation = safeString(edge.relation || edge.label).trim();
        if (!from || !to || !relation) return null;

        const edgeKey = [from, relation, to].join("\u001f");
        return {
            id: scopeId + "::edge::" + hashText(edgeKey),
            scopeId: scopeId,
            edgeKey: edgeKey,
            from: from,
            to: to,
            relation: relation,
            evidenceResultIds: Array.isArray(edge.evidenceResultIds)
                ? edge.evidenceResultIds.map(safeString)
                : [],
            updatedAt: new Date().toISOString()
        };
    }

    async function saveKnowledgeMap(scope, map, status) {
        const scopeId = makeScopeId(scope);
        const input = map && typeof map === "object" ? map : {};
        const db = await openDB();
        const transaction = db.transaction([STORES.nodes, STORES.edges, STORES.meta], "readwrite");
        const nodesStore = transaction.objectStore(STORES.nodes);
        const edgesStore = transaction.objectStore(STORES.edges);
        const metaStore = transaction.objectStore(STORES.meta);

        const existingNodes = await requestResult(nodesStore.index("scopeId").getAll(scopeId));
        const existingEdges = await requestResult(edgesStore.index("scopeId").getAll(scopeId));
        const nodeMap = new Map(existingNodes.map(function (item) { return [item.nodeKey, item]; }));
        const edgeMap = new Map(existingEdges.map(function (item) { return [item.edgeKey, item]; }));

        (Array.isArray(input.nodes) ? input.nodes : []).forEach(function (rawNode) {
            const node = normalizeMapNode(scopeId, rawNode);
            if (!node) return;
            const old = nodeMap.get(node.nodeKey);
            if (old) {
                node.aliases = Array.from(new Set((old.aliases || []).concat(node.aliases)));
                node.resultIds = Array.from(new Set((old.resultIds || []).concat(node.resultIds)));
                if (!node.description) node.description = old.description || "";
            }
            nodesStore.put(node);
        });

        (Array.isArray(input.edges) ? input.edges : []).forEach(function (rawEdge) {
            const edge = normalizeMapEdge(scopeId, rawEdge);
            if (!edge) return;
            const old = edgeMap.get(edge.edgeKey);
            if (old) {
                edge.evidenceResultIds = Array.from(new Set(
                    (old.evidenceResultIds || []).concat(edge.evidenceResultIds)
                ));
            }
            edgesStore.put(edge);
        });

        metaStore.put({
            scopeId: scopeId,
            schemaVersion: 1,
            mapVersion: Number(input.version) || 1,
            status: safeString(status || "ready"),
            updatedAt: new Date().toISOString(),
            summary: safeString(input.summary)
        });

        await transactionDone(transaction);
        return { nodes: (input.nodes || []).length, edges: (input.edges || []).length };
    }

    async function getKnowledgeMap(scope) {
        const scopeId = makeScopeId(scope);
        const db = await openDB();
        const transaction = db.transaction([STORES.nodes, STORES.edges, STORES.meta], "readonly");
        const nodesRequest = transaction.objectStore(STORES.nodes).index("scopeId").getAll(scopeId);
        const edgesRequest = transaction.objectStore(STORES.edges).index("scopeId").getAll(scopeId);
        const metaRequest = transaction.objectStore(STORES.meta).get(scopeId);

        const values = await Promise.all([
            requestResult(nodesRequest),
            requestResult(edgesRequest),
            requestResult(metaRequest)
        ]);

        return {
            scopeId: scopeId,
            meta: values[2] || null,
            nodes: values[0] || [],
            edges: values[1] || []
        };
    }

    async function getStats(scope) {
        const scopeId = makeScopeId(scope);
        const values = await Promise.all([
            getResults(scopeId),
            getKnowledgeMap(scopeId)
        ]);

        return {
            scopeId: scopeId,
            results: values[0].length,
            nodes: values[1].nodes.length,
            edges: values[1].edges.length,
            mapUpdatedAt: values[1].meta ? values[1].meta.updatedAt : null,
            mapStatus: values[1].meta ? values[1].meta.status : "not-built"
        };
    }

    async function deleteScope(scope) {
        const scopeId = makeScopeId(scope);
        const db = await openDB();
        const transaction = db.transaction(
            [STORES.results, STORES.nodes, STORES.edges, STORES.meta],
            "readwrite"
        );

        [STORES.results, STORES.nodes, STORES.edges].forEach(function (storeName) {
            const store = transaction.objectStore(storeName);
            const request = store.index("scopeId").openKeyCursor(IDBKeyRange.only(scopeId));
            request.onsuccess = function () {
                const cursor = request.result;
                if (cursor) {
                    store.delete(cursor.primaryKey);
                    cursor.continue();
                }
            };
        });

        transaction.objectStore(STORES.meta).delete(scopeId);
        await transactionDone(transaction);
        return true;
    }

    global.ResearchMemory = Object.freeze({
        version: 1,
        open: openDB,
        makeScopeId: makeScopeId,
        saveResults: saveResults,
        getResults: getResults,
        search: search,
        saveKnowledgeMap: saveKnowledgeMap,
        getKnowledgeMap: getKnowledgeMap,
        getStats: getStats,
        deleteScope: deleteScope
    });

})(window);
