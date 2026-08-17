/*
 * Word AI Assistant - consolidated controller
 * Search engine: Orama (primary and only retrieval engine)
 * Keeps the existing UI IDs and the existing chat/provider behavior.
 *
 * Orama notes:
 * - Arabic is a supported built-in language.
 * - The main retrieval index stores paragraphs and headings in one DB.
 * - Heading text receives a higher search boost.
 * - searchIndexedDocument() is retained as the public/internal compatibility
 *   name so the AI layer and older console tests continue to work.
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

    const settingsBtn = document.getElementById("settings-btn");
    const settingsWindow = document.getElementById("settings-window");
    const closeSettings = document.getElementById("close-settings");
    const showKey = document.getElementById("show-key");
    const apiKey = document.getElementById("api-key");
    const provider = document.getElementById("provider-select");
    const modelSelect = document.getElementById("model-select");
    const refreshModels = document.getElementById("refresh-models");
    const saveSettingsBtn = document.getElementById("save-settings");
    const testConnection = document.getElementById("test-connection");
    const settingsStatus = document.getElementById("settings-status");
    const providerInfo = document.getElementById("provider-info");

    const chatPopup = document.getElementById("chat-popup");
    const recentChatList = document.getElementById("recent-chat-list");

    const searchPopup = document.getElementById("search-popup");
    const searchInput = document.getElementById("search-input");
    const searchResults = document.getElementById("search-results");
    const searchBtn = document.getElementById("search-btn");

    let sidebarTogglePlaceholder = document.createComment(
        "sidebar-toggle-placeholder"
    );

    if (
        sidebarToggleBtn &&
        sidebarToggleBtn.parentNode
    ) {

        sidebarToggleBtn.parentNode.insertBefore(
            sidebarTogglePlaceholder,
            sidebarToggleBtn
        );

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


    // =====================================================
    // ORAMA CACHE
    // محرك البحث الوحيد المستخدم في الاسترجاع
    // =====================================================

    let oramaRetrievalDb = null;
    let oramaRetrievalCacheKey = "";
    let oramaRetrievalDocumentId = null;


    // =====================================================
    // DOCUMENT DATABASE
    // =====================================================

    const DOCUMENT_DB_NAME =
        "WORD_AI_DOCUMENT_STORAGE";

    const DOCUMENT_DB_VERSION =
        4;

    const DOCUMENT_STORE_NAME =
        "files";

    const DOCUMENT_TEXT_STORE_NAME =
        "texts";

    const DOCUMENT_STRUCTURE_STORE_NAME =
        "structures";


    // =====================================================
    // ORAMA SCHEMA VERSION
    // تغيير الرقم يجبر المحرك على إعادة بناء ذاكرة Orama
    // =====================================================

    const ORAMA_SCHEMA_VERSION =
        1;


    // =====================================================
    // SAFE STORAGE HELPERS
    // =====================================================

    function readStorageArray(
        key
    ) {

        try {

            const value =
                JSON.parse(
                    localStorage.getItem(
                        key
                    ) ||
                    "[]"
                );


            return (
                Array.isArray(
                    value
                )
                    ? value
                    : []
            );

        }
        catch (
            error
        ) {

            console.warn(
                "تعذر قراءة التخزين:",
                key,
                error
            );


            return [];

        }

    }


    // =====================================================
    // LOAD PROJECTS
    // =====================================================

    projects =
        readStorageArray(
            "WORD_AI_PROJECTS"
        )
        .filter(
            function (
                item
            ) {

                return (
                    item &&
                    typeof item ===
                        "object"
                );

            }
        )
        .map(
            function (
                item
            ) {

                const now =
                    new Date()
                        .toISOString();


                return {

                    id:
                        item.id ||
                        (
                            Date.now() +
                            Math.random()
                        ),

                    name:
                        item.name ||
                        "مشروع جديد",

                    createdAt:
                        item.createdAt ||
                        now,

                    updatedAt:
                        item.updatedAt ||
                        now,

                    documents:
                        Array.isArray(
                            item.documents
                        )
                            ? item.documents
                            : [],

                    references:
                        Array.isArray(
                            item.references
                        )
                            ? item.references
                            : [],

                    chatIds:
                        Array.isArray(
                            item.chatIds
                        )
                            ? item.chatIds
                            : [],

                    settings:
                        item.settings &&
                        typeof item.settings ===
                            "object"
                            ? item.settings
                            : {
                                citationStyle:
                                    "",
                                notes:
                                    ""
                            }

                };

            }
        );


    // =====================================================
    // LOAD DOCUMENTS
    // =====================================================

    documents =
        readStorageArray(
            "WORD_AI_DOCUMENTS"
        )
        .filter(
            function (
                item
            ) {

                return (
                    item &&
                    typeof item ===
                        "object"
                );

            }
        )
        .map(
            function (
                item
            ) {

                return {

                    ...item,

                    indexTokenCount:
                        Number(
                            item.indexTokenCount ||
                            0
                        ),

                    indexUniqueTerms:
                        Number(
                            item.indexUniqueTerms ||
                            0
                        ),

                    indexUniqueFamilies:
                        Number(
                            item.indexUniqueFamilies ||
                            0
                        ),

                    indexSchemaVersion:
                        Number(
                            item.indexSchemaVersion ||
                            0
                        ),

                    indexStatus:
                        item.indexStatus ||
                        "new",

                    readStatus:
                        item.readStatus ||
                        "new"

                };

            }
        );


    // =====================================================
    // LOAD CHATS
    // =====================================================

    chats =
        readStorageArray(
            "WORD_AI_CHATS"
        )
        .filter(
            function (
                item
            ) {

                return (
                    item &&
                    typeof item ===
                        "object"
                );

            }
        )
        .map(
            function (
                item
            ) {

                return {

                    ...item,

                    messages:
                        Array.isArray(
                            item.messages
                        )
                            ? item.messages
                            : []

                };

            }
        );


    // =====================================================
    // SAVE PROJECTS
    // =====================================================

    function saveProjects() {

        localStorage.setItem(
            "WORD_AI_PROJECTS",
            JSON.stringify(
                projects
            )
        );

    }


    // =====================================================
    // SAVE DOCUMENTS
    // =====================================================

    function saveDocuments() {

        localStorage.setItem(
            "WORD_AI_DOCUMENTS",
            JSON.stringify(
                documents
            )
        );

    }


    // =====================================================
    // SAVE CHATS
    // =====================================================

    function saveChats() {

        localStorage.setItem(
            "WORD_AI_CHATS",
            JSON.stringify(
                chats
            )
        );

    }


    saveProjects();
    saveDocuments();
    saveChats();


    // =====================================================
    // INDEXED DB
    // =====================================================

    function openDocumentDatabase() {

        return new Promise(
            function (
                resolve,
                reject
            ) {

                const request =
                    indexedDB.open(
                        DOCUMENT_DB_NAME,
                        DOCUMENT_DB_VERSION
                    );


                request.onupgradeneeded =
                    function () {

                        const db =
                            request.result;


                        if (
                            !db.objectStoreNames
                                .contains(
                                    DOCUMENT_STORE_NAME
                                )
                        ) {

                            db.createObjectStore(
                                DOCUMENT_STORE_NAME
                            );

                        }


                        if (
                            !db.objectStoreNames
                                .contains(
                                    DOCUMENT_TEXT_STORE_NAME
                                )
                        ) {

                            db.createObjectStore(
                                DOCUMENT_TEXT_STORE_NAME
                            );

                        }


                        if (
                            !db.objectStoreNames
                                .contains(
                                    DOCUMENT_STRUCTURE_STORE_NAME
                                )
                        ) {

                            db.createObjectStore(
                                DOCUMENT_STRUCTURE_STORE_NAME
                            );

                        }

                    };


                request.onsuccess =
                    function () {

                        resolve(
                            request.result
                        );

                    };


                request.onerror =
                    function () {

                        reject(
                            request.error ||
                            new Error(
                                "فشل فتح قاعدة المستندات."
                            )
                        );

                    };

            }
        );

    }


    // =====================================================
    // SAVE WORKING WORD FILE
    // =====================================================

    async function saveWorkingWordFile(
        fileId,
        file
    ) {

        const db =
            await openDocumentDatabase();


        return new Promise(
            function (
                resolve,
                reject
            ) {

                const tx =
                    db.transaction(
                        DOCUMENT_STORE_NAME,
                        "readwrite"
                    );


                const store =
                    tx.objectStore(
                        DOCUMENT_STORE_NAME
                    );


                const request =
                    store.put(
                        file,
                        String(
                            fileId
                        )
                    );


                request.onerror =
                    function () {

                        reject(
                            request.error
                        );

                    };


                tx.oncomplete =
                    function () {

                        db.close();

                        resolve(
                            String(
                                fileId
                            )
                        );

                    };


                tx.onerror =
                    function () {

                        reject(
                            tx.error
                        );

                    };

            }
        );

    }


    // =====================================================
    // GET WORKING WORD FILE
    // =====================================================

    async function getWorkingWordFile(
        fileId
    ) {

        const db =
            await openDocumentDatabase();


        return new Promise(
            function (
                resolve,
                reject
            ) {

                const tx =
                    db.transaction(
                        DOCUMENT_STORE_NAME,
                        "readonly"
                    );


                const request =
                    tx.objectStore(
                        DOCUMENT_STORE_NAME
                    )
                    .get(
                        String(
                            fileId
                        )
                    );


                request.onsuccess =
                    function () {

                        resolve(
                            request.result ||
                            null
                        );

                    };


                request.onerror =
                    function () {

                        reject(
                            request.error
                        );

                    };


                tx.oncomplete =
                    function () {

                        db.close();

                    };

            }
        );

    }


    // =====================================================
    // DELETE WORKING WORD FILE
    // =====================================================

    async function deleteWorkingWordFile(
        fileId
    ) {

        const db =
            await openDocumentDatabase();


        return new Promise(
            function (
                resolve,
                reject
            ) {

                const tx =
                    db.transaction(
                        DOCUMENT_STORE_NAME,
                        "readwrite"
                    );


                const request =
                    tx.objectStore(
                        DOCUMENT_STORE_NAME
                    )
                    .delete(
                        String(
                            fileId
                        )
                    );


                request.onerror =
                    function () {

                        reject(
                            request.error
                        );

                    };


                tx.oncomplete =
                    function () {

                        db.close();

                        resolve();

                    };


                tx.onerror =
                    function () {

                        reject(
                            tx.error
                        );

                    };

            }
        );

    }


    // =====================================================
    // SAVE DOCUMENT TEXT
    // =====================================================

    async function saveDocumentText(
        documentId,
        text
    ) {

        const db =
            await openDocumentDatabase();


        return new Promise(
            function (
                resolve,
                reject
            ) {

                const tx =
                    db.transaction(
                        DOCUMENT_TEXT_STORE_NAME,
                        "readwrite"
                    );


                const record = {

                    documentId:
                        String(
                            documentId
                        ),

                    text:
                        String(
                            text ||
                            ""
                        ),

                    updatedAt:
                        new Date()
                            .toISOString()

                };


                const request =
                    tx.objectStore(
                        DOCUMENT_TEXT_STORE_NAME
                    )
                    .put(
                        record,
                        String(
                            documentId
                        )
                    );


                request.onerror =
                    function () {

                        reject(
                            request.error
                        );

                    };


                tx.oncomplete =
                    function () {

                        db.close();

                        resolve(
                            record
                        );

                    };


                tx.onerror =
                    function () {

                        reject(
                            tx.error
                        );

                    };

            }
        );

    }


    // =====================================================
    // GET DOCUMENT TEXT
    // =====================================================

    async function getDocumentText(
        documentId
    ) {

        const db =
            await openDocumentDatabase();


        return new Promise(
            function (
                resolve,
                reject
            ) {

                const tx =
                    db.transaction(
                        DOCUMENT_TEXT_STORE_NAME,
                        "readonly"
                    );


                const request =
                    tx.objectStore(
                        DOCUMENT_TEXT_STORE_NAME
                    )
                    .get(
                        String(
                            documentId
                        )
                    );


                request.onsuccess =
                    function () {

                        resolve(
                            request.result ||
                            null
                        );

                    };


                request.onerror =
                    function () {

                        reject(
                            request.error
                        );

                    };


                tx.oncomplete =
                    function () {

                        db.close();

                    };

            }
        );

    }


    // =====================================================
    // SAVE DOCUMENT STRUCTURE
    // =====================================================

    async function saveDocumentStructure(
        documentId,
        structureData
    ) {

        const db =
            await openDocumentDatabase();


        return new Promise(
            function (
                resolve,
                reject
            ) {

                const tx =
                    db.transaction(
                        DOCUMENT_STRUCTURE_STORE_NAME,
                        "readwrite"
                    );


                const request =
                    tx.objectStore(
                        DOCUMENT_STRUCTURE_STORE_NAME
                    )
                    .put(
                        structureData,
                        String(
                            documentId
                        )
                    );


                request.onerror =
                    function () {

                        reject(
                            request.error
                        );

                    };


                tx.oncomplete =
                    function () {

                        db.close();

                        resolve(
                            structureData
                        );

                    };


                tx.onerror =
                    function () {

                        reject(
                            tx.error
                        );

                    };

            }
        );

    }


    // =====================================================
    // GET DOCUMENT STRUCTURE
    // =====================================================

    async function getDocumentStructure(
        documentId
    ) {

        const db =
            await openDocumentDatabase();


        return new Promise(
            function (
                resolve,
                reject
            ) {

                const tx =
                    db.transaction(
                        DOCUMENT_STRUCTURE_STORE_NAME,
                        "readonly"
                    );


                const request =
                    tx.objectStore(
                        DOCUMENT_STRUCTURE_STORE_NAME
                    )
                    .get(
                        String(
                            documentId
                        )
                    );


                request.onsuccess =
                    function () {

                        resolve(
                            request.result ||
                            null
                        );

                    };


                request.onerror =
                    function () {

                        reject(
                            request.error
                        );

                    };


                tx.oncomplete =
                    function () {

                        db.close();

                    };

            }
        );

    }


    // =====================================================
    // FILE -> BASE64
    // =====================================================

    function fileToBase64(
        file
    ) {

        return new Promise(
            function (
                resolve,
                reject
            ) {

                const reader =
                    new FileReader();


                reader.onload =
                    function () {

                        const result =
                            String(
                                reader.result ||
                                ""
                            );


                        const commaIndex =
                            result.indexOf(
                                ","
                            );


                        if (
                            commaIndex ===
                            -1
                        ) {

                            reject(
                                new Error(
                                    "تعذر تحويل ملف Word إلى Base64."
                                )
                            );

                            return;

                        }


                        resolve(
                            result.substring(
                                commaIndex +
                                1
                            )
                        );

                    };


                reader.onerror =
                    function () {

                        reject(
                            reader.error ||
                            new Error(
                                "فشل قراءة ملف Word."
                            )
                        );

                    };


                reader.readAsDataURL(
                    file
                );

            }
        );

    }


    // =====================================================
    // CREATE DOCUMENT
    // =====================================================

    function createDocument(
        file,
        projectId,
        order
    ) {

        const now =
            new Date()
                .toISOString();


        const documentId =
            Date.now();


        const item = {

            id:
                documentId,

            projectId:
                projectId,

            name:
                file.name.replace(
                    /\.docx$/i,
                    ""
                ),

            fileName:
                file.name,

            fileType:
                file.type ||
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",

            storageId:
                String(
                    documentId
                ),

            order:
                typeof order ===
                    "number"
                        ? order
                        : 0,

            type:
                "word",

            readStatus:
                "new",

            indexStatus:
                "new",

            indexTokenCount:
                0,

            indexUniqueTerms:
                0,

            indexUniqueFamilies:
                0,

            indexSchemaVersion:
                0,

            createdAt:
                now,

            updatedAt:
                now

        };


        documents.push(
            item
        );


        saveDocuments();


        return item;

    }


    // =====================================================
    // DOCUMENT STATUS
    // =====================================================

    function updateDocumentReadStatus(
        item,
        status
    ) {

        if (!item) {

            return;

        }


        item.readStatus =
            status;


        item.updatedAt =
            new Date()
                .toISOString();


        if (
            status ===
            "read"
        ) {

            item.readAt =
                new Date()
                    .toISOString();

        }


        saveDocuments();

    }


    function updateDocumentIndexStatus(
        item,
        status
    ) {

        if (!item) {

            return;

        }


        item.indexStatus =
            status;


        item.updatedAt =
            new Date()
                .toISOString();


        saveDocuments();

    }


    // =====================================================
    // PROJECT DOCUMENTS
    // =====================================================

    function getProjectDocuments(
        projectId
    ) {

        if (!projectId) {

            return [];

        }


        return documents
            .filter(
                function (
                    item
                ) {

                    return (
                        item &&
                        item.projectId ===
                            projectId
                    );

                }
            )
            .sort(
                function (
                    a,
                    b
                ) {

                    return (
                        Number(
                            a.order ||
                            0
                        ) -
                        Number(
                            b.order ||
                            0
                        )
                    );

                }
            );

    }


    function attachDocumentToProject(
        project,
        item
    ) {

        if (
            !project ||
            !item
        ) {

            return;

        }


        if (
            !Array.isArray(
                project.documents
            )
        ) {

            project.documents =
                [];

        }


        if (
            !project.documents.includes(
                item.id
            )
        ) {

            project.documents.push(
                item.id
            );


            project.updatedAt =
                new Date()
                    .toISOString();


            saveProjects();

        }

    }


    function setCurrentProject(
        project
    ) {

        currentProject =
            project ||
            null;


        renderDocuments();

    }


    // =====================================================
    // TEXT NORMALIZATION
    // =====================================================

    function normalizeSearchText(
        text
    ) {

        return String(
            text ||
            ""
        )

        // إزالة التشكيل
        .replace(
            /[\u064B-\u065F\u0670]/g,
            ""
        )

        // إزالة التطويل
        .replace(
            /\u0640/g,
            ""
        )

        // توحيد الألف
        .replace(
            /[أإآٱ]/g,
            "ا"
        )

        // توحيد الياء
        .replace(
            /ى/g,
            "ي"
        )

        // توحيد بعض أشكال الهمزة
        .replace(
            /ؤ/g,
            "و"
        )
        .replace(
            /ئ/g,
            "ي"
        )

        // تنظيف المسافات
        .replace(
            /\s+/g,
            " "
        )

        .trim()
        .toLowerCase();

    }


    function tokenizeDocumentText(
        text
    ) {

        return (
            normalizeSearchText(
                text
            )
            .match(
                /[\p{L}\p{N}]+/gu
            ) ||
            []
        );

    }


    function getSearchQueryTokens(
        query
    ) {

        const stopWords =
            new Set([

                "ما",
                "ماذا",
                "من",
                "هو",
                "هي",
                "هم",
                "في",
                "على",
                "عن",
                "الى",
                "إلى",
                "منه",
                "بها",
                "به",
                "لها",
                "له",
                "هذا",
                "هذه",
                "ذلك",
                "تلك",
                "الذي",
                "التي",
                "الذين",
                "بين",
                "مع",
                "ثم",
                "او",
                "أو",
                "و",
                "ف",
                "ب",
                "ك",
                "ل",
                "أن",
                "إن",
                "هل",
                "كيف",
                "لماذا",
                "أي",
                "اي",
                "كان",
                "كانت",
                "يكون",
                "تكون",
                "فيها",
                "فيه"

            ]);


        const tokens =
            tokenizeDocumentText(
                query
            );


        const filtered =
            tokens.filter(
                function (
                    token
                ) {

                    return (
                        token.length >
                            2 &&
                        !stopWords.has(
                            token
                        )
                    );

                }
            );


        return (
            filtered.length
                ? filtered
                : tokens
        );

    }


    // =====================================================
    // BUILD DOCUMENT STRUCTURE
    // =====================================================

    async function buildDocumentStructure(
        documentItem
    ) {

        if (!documentItem) {

            throw new Error(
                "لم يتم تحديد المستند."
            );

        }


        const file =
            await getWorkingWordFile(
                documentItem.storageId
            );


        if (!file) {

            throw new Error(
                "لم يتم العثور على نسخة العمل."
            );

        }


        const base64 =
            await fileToBase64(
                file
            );


        return await Word.run(
            async function (
                context
            ) {

                if (
                    !Office.context.requirements
                        .isSetSupported(
                            "WordApiHiddenDocument",
                            "1.3"
                        )
                ) {

                    throw new Error(
                        "إصدار Word الحالي لا يدعم تحليل بنية المستند."
                    );

                }


                const workingDocument =
                    context.application
                        .createDocument(
                            base64
                        );


                const paragraphs =
                    workingDocument.body
                        .paragraphs;


                const tables =
                    workingDocument.body
                        .tables;


                paragraphs.load(
                    [
                        "items/text",
                        "items/styleBuiltIn",
                        "items/tableNestingLevel"
                    ]
                );


                tables.load(
                    [
                        "items/rowCount",
                        "items/columnCount",
                        "items/styleBuiltIn"
                    ]
                );


                await context.sync();


                const paragraphItems =
                    paragraphs.items.map(
                        function (
                            paragraph,
                            index
                        ) {

                            return {

                                index:
                                    index,

                                id:
                                    String(
                                        index
                                    ),

                                text:
                                    String(
                                        paragraph.text ||
                                        ""
                                    ).trim(),

                                style:
                                    String(
                                        paragraph.styleBuiltIn ||
                                        ""
                                    ),

                                tableNestingLevel:
                                    Number(
                                        paragraph.tableNestingLevel ||
                                        0
                                    )

                            };

                        }
                    );


                const headings =
                    paragraphItems.filter(
                        function (
                            paragraph
                        ) {

                            return (
                                paragraph.style &&
                                /^Heading[1-9]$/i.test(
                                    paragraph.style
                                )
                            );

                        }
                    );


                const tableItems =
                    tables.items.map(
                        function (
                            table,
                            index
                        ) {

                            return {

                                index:
                                    index,

                                rows:
                                    Number(
                                        table.rowCount ||
                                        0
                                    ),

                                columns:
                                    Number(
                                        table.columnCount ||
                                        0
                                    ),

                                style:
                                    String(
                                        table.styleBuiltIn ||
                                        ""
                                    )

                            };

                        }
                    );


                return {

                    documentId:
                        String(
                            documentItem.id
                        ),

                    paragraphCount:
                        paragraphItems.length,

                    headingCount:
                        headings.length,

                    tableCount:
                        tableItems.length,

                    paragraphs:
                        paragraphItems,

                    headings:
                        headings,

                    tables:
                        tableItems,

                    updatedAt:
                        new Date()
                            .toISOString()

                };

            }
        );

    }


    async function ensureDocumentStructure(
        documentItem
    ) {

        const existing =
            await getDocumentStructure(
                documentItem.id
            );


        if (
            existing &&
            Array.isArray(
                existing.paragraphs
            ) &&
            Array.isArray(
                existing.headings
            )
        ) {

            return existing;

        }


        const structure =
            await buildDocumentStructure(
                documentItem
            );


        await saveDocumentStructure(
            documentItem.id,
            structure
        );


        return structure;

    }


    // =====================================================
    // ORAMA LOADER
    // =====================================================

    function loadOrama() {

        try {

            return require(
                "@orama/orama"
            );

        }
        catch (
            error
        ) {

            throw new Error(
                "تعذر تحميل Orama من الحزمة @orama/orama: " +
                error.message
            );

        }

    }


    // =====================================================
    // HEADING LEVEL
    // =====================================================

    function getHeadingLevelNumber(
        style
    ) {

        const match =
            String(
                style ||
                ""
            )
            .match(
                /Heading\s*([1-9])/i
            );


        return (
            match
                ? Number(
                    match[1]
                )
                : 9
        );

    }


    // =====================================================
    // BUILD ORAMA PRIMARY RETRIEVAL INDEX
    // الفهرس الرئيسي الوحيد
    // =====================================================

    async function buildOramaRetrievalIndex(
        documentItem,
        structureData
    ) {

        const {
            create,
            insertMultiple
        } =
            loadOrama();


        const paragraphs =
            Array.isArray(
                structureData.paragraphs
            )
                ? structureData.paragraphs
                : [];


        const headings =
            Array.isArray(
                structureData.headings
            )
                ? structureData.headings
                    .filter(
                        function (
                            heading
                        ) {

                            return (
                                heading &&
                                typeof heading.index !==
                                    "undefined" &&
                                String(
                                    heading.text ||
                                    ""
                                ).trim()
                            );

                        }
                    )
                    .sort(
                        function (
                            a,
                            b
                        ) {

                            return (
                                Number(
                                    a.index
                                ) -
                                Number(
                                    b.index
                                )
                            );

                        }
                    )
                : [];


        const headingMap =
            new Map();


        headings.forEach(
            function (
                heading
            ) {

                headingMap.set(
                    Number(
                        heading.index
                    ),
                    heading
                );

            }
        );


        function nearestHeading(
            paragraphIndex
        ) {

            if (
                headingMap.has(
                    paragraphIndex
                )
            ) {

                return headingMap.get(
                    paragraphIndex
                );

            }


            for (
                let i =
                    headings.length - 1;
                i >= 0;
                i--
            ) {

                if (
                    Number(
                        headings[i].index
                    ) <
                    paragraphIndex
                ) {

                    return headings[i];

                }

            }


            return null;

        }


        const records =
            [];


        paragraphs.forEach(
            function (
                paragraph
            ) {

                if (
                    !paragraph ||
                    !paragraph.text
                ) {

                    return;

                }


                const text =
                    String(
                        paragraph.text
                    ).trim();


                if (!text) {

                    return;

                }


                const paragraphIndex =
                    Number(
                        paragraph.index
                    );


                const heading =
                    nearestHeading(
                        paragraphIndex
                    );


                const isHeading =
                    Boolean(
                        heading &&
                        Number(
                            heading.index
                        ) ===
                        paragraphIndex
                    );


                records.push({

                    id:
                        "p-" +
                        String(
                            paragraphIndex
                        ),

                    paragraphIndex:
                        paragraphIndex,

                    text:
                        text,

                    heading:
                        heading
                            ? String(
                                heading.text ||
                                ""
                            ).trim()
                            : "",

                    headingIndex:
                        heading
                            ? Number(
                                heading.index
                            )
                            : -1,

                    headingLevel:
                        heading
                            ? String(
                                heading.style ||
                                ""
                            )
                            : "",

                    headingLevelNumber:
                        heading
                            ? getHeadingLevelNumber(
                                heading.style
                            )
                            : 9,

                    isHeading:
                        isHeading,

                    documentId:
                        String(
                            documentItem.id
                        )

                });

            }
        );


        const cacheKey =
            [

                String(
                    documentItem.id
                ),

                String(
                    documentItem.indexUpdatedAt ||
                    ""
                ),

                String(
                    ORAMA_SCHEMA_VERSION
                ),

                String(
                    records.length
                ),

                String(
                    headings.length
                )

            ].join(
                "|"
            );


        if (
            oramaRetrievalDb &&
            oramaRetrievalCacheKey ===
                cacheKey &&
            oramaRetrievalDocumentId ===
                String(
                    documentItem.id
                )
        ) {

            return oramaRetrievalDb;

        }


        // =================================================
        // Orama يدعم العربية
        // =================================================

        const db =
            create({

                schema: {

                    id:
                        "string",

                    paragraphIndex:
                        "number",

                    text:
                        "string",

                    heading:
                        "string",

                    headingIndex:
                        "number",

                    headingLevel:
                        "string",

                    headingLevelNumber:
                        "number",

                    isHeading:
                        "boolean",

                    documentId:
                        "string"

                },

                language:
                    "arabic"

            });


        if (
            records.length
        ) {

            await insertMultiple(
                db,
                records,
                500
            );

        }


        oramaRetrievalDb =
            db;


        oramaRetrievalCacheKey =
            cacheKey;


        oramaRetrievalDocumentId =
            String(
                documentItem.id
            );


        // =================================================
        // تحديث إحصاءات المستند
        // =================================================

        documentItem.indexStatus =
            "indexed";


        documentItem.indexTokenCount =
            records.reduce(
                function (
                    total,
                    record
                ) {

                    return (
                        total +
                        tokenizeDocumentText(
                            record.text
                        ).length
                    );

                },
                0
            );


        documentItem.indexUniqueTerms =
            new Set(
                records.flatMap(
                    function (
                        record
                    ) {

                        return tokenizeDocumentText(
                            record.text
                        );

                    }
                )
            )
            .size;


        // لم نعد نحتاج للفهرس العائلي القديم.
        documentItem.indexUniqueFamilies =
            0;


        documentItem.indexSchemaVersion =
            ORAMA_SCHEMA_VERSION;


        documentItem.indexUpdatedAt =
            new Date()
                .toISOString();


        saveDocuments();


        console.log(
            "تم بناء فهرس Orama الرئيسي:",
            {
                documentId:
                    documentItem.id,

                records:
                    records.length,

                headings:
                    headings.length
            }
        );


        return db;

    }


    // =====================================================
    // READ CURRENT WORKING DOCUMENT
    // =====================================================

    async function readCurrentWordDocument(
        documentItem
    ) {

        if (!documentItem) {

            throw new Error(
                "لم يتم تحديد المستند."
            );

        }


        updateDocumentReadStatus(
            documentItem,
            "reading"
        );


        updateDocumentIndexStatus(
            documentItem,
            "indexing"
        );


        try {

            if (
                !Office.context.requirements
                    .isSetSupported(
                        "WordApiHiddenDocument",
                        "1.3"
                    )
            ) {

                throw new Error(
                    "إصدار Word الحالي لا يدعم قراءة نسخة العمل."
                );

            }


            const file =
                await getWorkingWordFile(
                    documentItem.storageId
                );


            if (!file) {

                throw new Error(
                    "لم يتم العثور على نسخة العمل."
                );

            }


            const base64 =
                await fileToBase64(
                    file
                );


            const text =
                await Word.run(
                    async function (
                        context
                    ) {

                        const workingDocument =
                            context.application
                                .createDocument(
                                    base64
                                );


                        const body =
                            workingDocument.body;


                        body.load(
                            "text"
                        );


                        await context.sync();


                        return (
                            body.text ||
                            ""
                        );

                    }
                );


            await saveDocumentText(
                documentItem.id,
                text
            );


            updateDocumentReadStatus(
                documentItem,
                "read"
            );


            // =================================================
            // بناء البنية
            // =================================================

            const structureData =
                await buildDocumentStructure(
                    documentItem
                );


            await saveDocumentStructure(
                documentItem.id,
                structureData
            );


            // =================================================
            // بناء Orama الرئيسي
            // =================================================

            await buildOramaRetrievalIndex(
                documentItem,
                structureData
            );


            updateDocumentIndexStatus(
                documentItem,
                "indexed"
            );


            saveDocuments();


            console.log(
                "تمت قراءة المستند وفهرسته بواسطة Orama:",
                {
                    documentId:
                        documentItem.id,

                    paragraphCount:
                        structureData.paragraphCount,

                    headingCount:
                        structureData.headingCount,

                    indexTokenCount:
                        documentItem.indexTokenCount
                }
            );


            return text;

        }
        catch (
            error
        ) {

            updateDocumentReadStatus(
                documentItem,
                "error"
            );


            updateDocumentIndexStatus(
                documentItem,
                "error"
            );


            console.error(
                "فشل قراءة/فهرسة المستند:",
                error
            );


            throw error;

        }

    }


    // =====================================================
    // SET CURRENT DOCUMENT
    // =====================================================

    function setCurrentDocument(
        documentItem
    ) {

        if (!documentItem) {

            currentDocument =
                null;


            if (documentTitle) {

                documentTitle.textContent =
                    "لا يوجد مستند مفتوح";

            }


            renderDocuments();


            return;

        }


        currentDocument =
            documentItem;


        if (documentTitle) {

            documentTitle.textContent =
                documentItem.name;

        }


        if (
            documentItem.readStatus ===
            "read"
        ) {

            ensureDocumentStructure(
                documentItem
            )
            .then(
                function (
                    structureData
                ) {

                    return buildOramaRetrievalIndex(
                        documentItem,
                        structureData
                    );

                }
            )
            .then(
                function () {

                    renderDocuments();

                }
            )
            .catch(
                function (
                    error
                ) {

                    console.error(
                        "تعذر تجهيز فهرس Orama:",
                        error
                    );


                    renderDocuments();

                }
            );


            renderDocuments();


            return;

        }


        readCurrentWordDocument(
            documentItem
        )
        .then(
            function (
                text
            ) {

                console.log(
                    "محتوى نسخة العمل:",
                    text
                );


                renderDocuments();

            }
        )
        .catch(
            function (
                error
            ) {

                console.error(
                    "تعذر قراءة نسخة العمل:",
                    error
                );


                renderDocuments();

            }
        );


        renderDocuments();

    }


    // =====================================================
    // UPDATE DOCUMENT TIMESTAMP
    // =====================================================

    function touchDocument(
        documentItem
    ) {

        if (!documentItem) {

            return;

        }


        documentItem.updatedAt =
            new Date()
                .toISOString();

    }


    // =====================================================
    // ADD DOCUMENT
    // =====================================================

    if (
        addDocumentBtn &&
        wordDocumentPicker
    ) {

        addDocumentBtn.onclick =
            function (
                e
            ) {

                e.preventDefault();
                e.stopPropagation();


                if (!currentProject) {

                    if (documentsList) {

                        documentsList.innerHTML = `
                            <div class="empty-document">
                                اختر مشروعًا أولًا لإضافة مستند
                            </div>
                        `;

                    }


                    return;

                }


                wordDocumentPicker.value =
                    "";


                wordDocumentPicker.click();

            };


        wordDocumentPicker.onchange =
            async function () {

                try {

                    const file =
                        wordDocumentPicker.files &&
                        wordDocumentPicker.files[0];


                    if (!file) {

                        return;

                    }


                    if (
                        !/\.docx$/i.test(
                            file.name
                        )
                    ) {

                        console.warn(
                            "الملف المختار ليس DOCX."
                        );


                        return;

                    }


                    if (!currentProject) {

                        return;

                    }


                    const projectDocuments =
                        getProjectDocuments(
                            currentProject.id
                        );


                    const nextOrder =
                        projectDocuments.length +
                        1;


                    const documentItem =
                        createDocument(
                            file,
                            currentProject.id,
                            nextOrder
                        );


                    await saveWorkingWordFile(
                        documentItem.storageId,
                        file
                    );


                    attachDocumentToProject(
                        currentProject,
                        documentItem
                    );


                    setCurrentDocument(
                        documentItem
                    );


                    renderDocuments();


                    console.log(
                        "تم استيراد مستند Word:",
                        {
                            name:
                                documentItem.name,

                            fileName:
                                documentItem.fileName,

                            storageId:
                                documentItem.storageId
                        }
                    );

                }
                catch (
                    error
                ) {

                    console.error(
                        "فشل استيراد مستند Word:",
                        error
                    );


                    if (documentsList) {

                        documentsList.innerHTML = `
                            <div class="empty-document">
                                تعذر استيراد المستند
                            </div>
                        `;

                    }

                }

            };

    }

    // =====================================================
    // نهاية الجزء الأول
    // الجزء الثاني يتابع:
    // المشاريع + المستندات + المحادثات + الواجهة الجانبية
    // =====================================================
    // ======================================
// Word AI Assistant
// PART 2 / 4
// المشاريع + المستندات + المحادثات + الواجهة الجانبية
// ======================================


// =====================================================
// PROJECT / CHAT ICONS
// =====================================================

const projectIcon = `
    <svg
        width="16"
        height="16"
        viewBox="0 0 24 24"
        fill="none"
        stroke="#000000"
        stroke-width="1"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true">

        <path
            d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z">
        </path>

    </svg>
`;


const chatIcon = `
    <svg
        width="16"
        height="16"
        viewBox="0 0 24 24"
        fill="none"
        stroke="#000000"
        stroke-width="1"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true">

        <path
            d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z">
        </path>

    </svg>
`;
            if (
                documentItem.indexStatus ===
                "indexed"
            ) {

                status.textContent =
                    "✓ مفهرس · " +
                    documentItem.indexTokenCount +
                    " كلمة · " +
                    documentItem.indexUniqueTerms +
                    " فريدة";


                if (
                    documentItem.indexUniqueFamilies
                ) {

                    status.textContent +=
                        " · " +
                        documentItem.indexUniqueFamilies +
                        " عائلة";

                }

            }
            else if (
                documentItem.indexStatus ===
                "indexing"
            ) {

                status.textContent =
                    "جارٍ الفهرسة...";

            }
            else if (
                documentItem.indexStatus ===
                "error"
            ) {

                status.textContent =
                    "⚠ فشل الفهرسة";

            }
            else if (
                documentItem.readStatus ===
                "reading"
            ) {

                status.textContent =
                    "جارٍ القراءة...";

            }
            else if (
                documentItem.readStatus ===
                "read"
            ) {

                status.textContent =
                    "✓ تمت القراءة";

            }
            else {

                status.textContent =
                    "جديد";

            }


            title.onclick =
                function (e) {

                    e.preventDefault();
                    e.stopPropagation();


                    setCurrentDocument(
                        documentItem
                    );

                };


            const menuButton =
                document.createElement(
                    "button"
                );


            menuButton.className =
                "document-menu";


            menuButton.type =
                "button";


            menuButton.title =
                "خيارات المستند";


            menuButton.textContent =
                "⋮";


            const options =
                document.createElement(
                    "div"
                );


            options.className =
                "document-options-menu";


            options.innerHTML = `

                <div class="rename-document">
                    ✏ إعادة تسمية
                </div>

                <div class="move-document-up">
                    ↑ نقل إلى أعلى
                </div>

                <div class="move-document-down">
                    ↓ نقل إلى أسفل
                </div>

                <div class="delete-document">
                    🗑 حذف
                </div>

            `;


            if (index === 0) {

                const moveUp =
                    options.querySelector(
                        ".move-document-up"
                    );

                if (moveUp) {

                    moveUp.style.display =
                        "none";

                }

            }


            if (
                index ===
                projectDocuments.length - 1
            ) {

                const moveDown =
                    options.querySelector(
                        ".move-document-down"
                    );

                if (moveDown) {

                    moveDown.style.display =
                        "none";

                }

            }


            menuButton.onclick =
                function (e) {

                    e.preventDefault();
                    e.stopPropagation();


                    document
                        .querySelectorAll(
                            ".document-options-menu.open"
                        )
                        .forEach(
                            function (menu) {

                                if (
                                    menu !==
                                    options
                                ) {

                                    menu.classList.remove(
                                        "open"
                                    );

                                }

                            }
                        );


                    options.classList.toggle(
                        "open"
                    );

                };


            // ==================================
            // Rename
            // ==================================

            const renameButton =
                options.querySelector(
                    ".rename-document"
                );


            if (renameButton) {

                renameButton.onclick =
                    function (e) {

                        e.preventDefault();
                        e.stopPropagation();


                        options.classList.remove(
                            "open"
                        );


                        const oldName =
                            documentItem.name;


                        const inputRename =
                            document.createElement(
                                "input"
                            );


                        inputRename.className =
                            "edit-document-title";


                        inputRename.value =
                            oldName;


                        title.replaceWith(
                            inputRename
                        );


                        inputRename.focus();


                        inputRename.setSelectionRange(
                            inputRename.value.length,
                            inputRename.value.length
                        );


                        function finishRename(
                            saveChange
                        ) {

                            const newName =
                                inputRename.value.trim();


                            if (
                                saveChange &&
                                newName !== ""
                            ) {

                                documentItem.name =
                                    newName;


                                documentItem.updatedAt =
                                    new Date()
                                        .toISOString();


                                saveDocuments();

                            }
                            else {

                                documentItem.name =
                                    oldName;

                            }


                            renderDocuments();

                        }


                        inputRename.onkeydown =
                            function (
                                event
                            ) {

                                if (
                                    event.key ===
                                    "Enter"
                                ) {

                                    event.preventDefault();

                                    finishRename(
                                        true
                                    );

                                }


                                if (
                                    event.key ===
                                    "Escape"
                                ) {

                                    event.preventDefault();

                                    finishRename(
                                        false
                                    );

                                }

                            };

                    };

            }


            // ==================================
            // Move Up
            // ==================================

            const moveUpButton =
                options.querySelector(
                    ".move-document-up"
                );


            if (moveUpButton) {

                moveUpButton.onclick =
                    function (e) {

                        e.preventDefault();
                        e.stopPropagation();


                        if (index <= 0)
                            return;


                        const previousDocument =
                            projectDocuments[
                                index - 1
                            ];


                        const currentOrder =
                            documentItem.order;


                        documentItem.order =
                            previousDocument.order;


                        previousDocument.order =
                            currentOrder;


                        documentItem.updatedAt =
                            new Date()
                                .toISOString();


                        previousDocument.updatedAt =
                            new Date()
                                .toISOString();


                        saveDocuments();


                        options.classList.remove(
                            "open"
                        );


                        renderDocuments();

                    };

            }


            // ==================================
            // Move Down
            // ==================================

            const moveDownButton =
                options.querySelector(
                    ".move-document-down"
                );


            if (moveDownButton) {

                moveDownButton.onclick =
                    function (e) {

                        e.preventDefault();
                        e.stopPropagation();


                        if (
                            index >=
                            projectDocuments.length - 1
                        ) {

                            return;

                        }


                        const nextDocument =
                            projectDocuments[
                                index + 1
                            ];


                        const currentOrder =
                            documentItem.order;


                        documentItem.order =
                            nextDocument.order;


                        nextDocument.order =
                            currentOrder;


                        documentItem.updatedAt =
                            new Date()
                                .toISOString();


                        nextDocument.updatedAt =
                            new Date()
                                .toISOString();


                        saveDocuments();


                        options.classList.remove(
                            "open"
                        );


                        renderDocuments();

                    };

            }


            // ==================================
            // Delete
            // ==================================

            const deleteButton =
                options.querySelector(
                    ".delete-document"
                );


            if (deleteButton) {

                deleteButton.onclick =
                    function (e) {

                        e.preventDefault();
                        e.stopPropagation();


                        options.classList.remove(
                            "open"
                        );


                        const oldConfirm =
                            document.querySelector(
                                ".document-delete-confirm"
                            );


                        if (oldConfirm) {

                            oldConfirm.remove();

                        }


                        const confirmBox =
                            document.createElement(
                                "div"
                            );


                        confirmBox.className =
                            "document-delete-confirm";


                        confirmBox.innerHTML = `
                            <div class="document-delete-dialog">

                                <div class="document-delete-message">
                                    هل تريد حذف المستند؟
                                </div>

                                <div class="document-delete-name">
                                    ${documentItem.name}
                                </div>

                                <div class="document-delete-buttons">

                                    <button
                                        type="button"
                                        class="confirm-document-delete">
                                        حذف
                                    </button>

                                    <button
                                        type="button"
                                        class="cancel-document-delete">
                                        إلغاء
                                    </button>

                                </div>

                            </div>
                        `;


                        document.body.appendChild(
                            confirmBox
                        );


                        const confirmDelete =
                            confirmBox.querySelector(
                                ".confirm-document-delete"
                            );


                        if (confirmDelete) {

                            confirmDelete.onclick =
                                async function () {

                                    documents =
                                        documents.filter(
                                            function (
                                                doc
                                            ) {

                                                return (
                                                    doc.id !==
                                                    documentItem.id
                                                );

                                            }
                                        );


                                    if (
                                        currentProject &&
                                        Array.isArray(
                                            currentProject.documents
                                        )
                                    ) {

                                        currentProject.documents =
                                            currentProject.documents.filter(
                                                function (
                                                    id
                                                ) {

                                                    return (
                                                        id !==
                                                        documentItem.id
                                                    );

                                                }
                                            );


                                        currentProject.updatedAt =
                                            new Date()
                                                .toISOString();


                                        saveProjects();

                                    }


                                    try {

                                        await deleteWorkingWordFile(
                                            documentItem.storageId
                                        );

                                    }
                                    catch (
                                        storageError
                                    ) {

                                        console.warn(
                                            "تعذر حذف نسخة العمل:",
                                            storageError
                                        );

                                    }


                                    if (
                                        currentDocument &&
                                        currentDocument.id ===
                                            documentItem.id
                                    ) {

                                        currentDocument =
                                            null;


                                        if (
                                            documentTitle
                                        ) {

                                            documentTitle.textContent =
                                                "لا يوجد مستند مفتوح";

                                        }

                                    }


                                    const remaining =
                                        getProjectDocuments(
                                            currentProject.id
                                        );


                                    remaining.forEach(
                                        function (
                                            doc,
                                            newIndex
                                        ) {

                                            doc.order =
                                                newIndex + 1;


                                            doc.updatedAt =
                                                new Date()
                                                    .toISOString();

                                        }
                                    );


                                    saveDocuments();


                                    confirmBox.remove();


                                    renderDocuments();

                                };

                        }


                        const cancelDelete =
                            confirmBox.querySelector(
                                ".cancel-document-delete"
                            );


                        if (cancelDelete) {

                            cancelDelete.onclick =
                                function () {

                                    confirmBox.remove();

                                };

                        }


                        confirmBox.onclick =
                            function (
                                event
                            ) {

                                if (
                                    event.target ===
                                    confirmBox
                                ) {

                                    confirmBox.remove();

                                }

                            };

                    };

            }


            item.appendChild(
                title
            );


            item.appendChild(
                status
            );


            item.appendChild(
                menuButton
            );


            item.appendChild(
                options
            );


            documentsList.appendChild(
                item
            );

        }
    );

}

// =====================================================
// RENDER DOCUMENTS
// =====================================================

function renderDocuments() {

    if (!documentsList) {

        return;

    }


    documentsList.innerHTML = "";


    if (!currentProject) {

        documentsList.innerHTML = `
            <div class="empty-document">
                اختر مشروعًا لعرض مستنداته
            </div>
        `;

        return;

    }


    const projectDocuments =
        getProjectDocuments(
            currentProject.id
        );


    if (
        projectDocuments.length ===
        0
    ) {

        documentsList.innerHTML = `
            <div class="empty-document">
                لا توجد مستندات
            </div>
        `;

        return;

    }


    projectDocuments.forEach(
        function (
            documentItem,
            index
        ) {

            const item =
                document.createElement(
                    "div"
                );


            item.className =
                "document-item";


            if (
                currentDocument &&
                String(
                    currentDocument.id
                ) ===
                String(
                    documentItem.id
                )
            ) {

                item.classList.add(
                    "active-document"
                );

            }


            const title =
                document.createElement(
                    "span"
                );


            title.className =
                "document-title";


            title.textContent =
                documentItem.name;


            const status =
                document.createElement(
                    "span"
                );


            status.className =
                "document-read-status";


            if (
                documentItem.indexStatus ===
                "indexed"
            ) {

                status.textContent =
                    "✓ مفهرس";


                if (
                    documentItem.indexTokenCount
                ) {

                    status.textContent +=
                        " · " +
                        documentItem.indexTokenCount +
                        " كلمة";

                }

            }
            else if (
                documentItem.indexStatus ===
                "indexing"
            ) {

                status.textContent =
                    "جارٍ الفهرسة...";

            }
            else if (
                documentItem.indexStatus ===
                "error"
            ) {

                status.textContent =
                    "⚠ فشل الفهرسة";

            }
            else if (
                documentItem.readStatus ===
                "reading"
            ) {

                status.textContent =
                    "جارٍ القراءة...";

            }
            else if (
                documentItem.readStatus ===
                "read"
            ) {

                status.textContent =
                    "✓ تمت القراءة";

            }
            else {

                status.textContent =
                    "جديد";

            }


            title.onclick =
                function (
                    e
                ) {

                    e.preventDefault();
                    e.stopPropagation();


                    setCurrentDocument(
                        documentItem
                    );

                };


            // =================================================
            // DOCUMENT MENU
            // =================================================

            const menuButton =
                document.createElement(
                    "button"
                );


            menuButton.className =
                "document-menu";


            menuButton.type =
                "button";


            menuButton.title =
                "خيارات المستند";


            menuButton.textContent =
                "⋮";


            const options =
                document.createElement(
                    "div"
                );


            options.className =
                "document-options-menu";


            options.innerHTML = `
                <div class="rename-document">
                    ✏ إعادة تسمية
                </div>

                <div class="move-document-up">
                    ↑ نقل إلى أعلى
                </div>

                <div class="move-document-down">
                    ↓ نقل إلى أسفل
                </div>

                <div class="delete-document">
                    🗑 حذف
                </div>
            `;


            if (
                index ===
                0
            ) {

                const moveUp =
                    options.querySelector(
                        ".move-document-up"
                    );


                if (moveUp) {

                    moveUp.style.display =
                        "none";

                }

            }


            if (
                index ===
                projectDocuments.length -
                1
            ) {

                const moveDown =
                    options.querySelector(
                        ".move-document-down"
                    );


                if (moveDown) {

                    moveDown.style.display =
                        "none";

                }

            }


            menuButton.onclick =
                function (
                    e
                ) {

                    e.preventDefault();
                    e.stopPropagation();


                    document
                        .querySelectorAll(
                            ".document-options-menu.open"
                        )
                        .forEach(
                            function (
                                menu
                            ) {

                                if (
                                    menu !==
                                    options
                                ) {

                                    menu.classList.remove(
                                        "open"
                                    );

                                }

                            }
                        );


                    options.classList.toggle(
                        "open"
                    );

                };


            // =================================================
            // RENAME DOCUMENT
            // =================================================

            const renameButton =
                options.querySelector(
                    ".rename-document"
                );


            if (renameButton) {

                renameButton.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();
                        e.stopPropagation();


                        options.classList.remove(
                            "open"
                        );


                        const oldName =
                            documentItem.name;


                        const inputRename =
                            document.createElement(
                                "input"
                            );


                        inputRename.className =
                            "edit-document-title";


                        inputRename.value =
                            oldName;


                        title.replaceWith(
                            inputRename
                        );


                        inputRename.focus();


                        inputRename.setSelectionRange(
                            inputRename.value.length,
                            inputRename.value.length
                        );


                        function finishRename(
                            saveChange
                        ) {

                            const newName =
                                inputRename.value
                                    .trim();


                            if (
                                saveChange &&
                                newName
                            ) {

                                documentItem.name =
                                    newName;


                                documentItem.updatedAt =
                                    new Date()
                                        .toISOString();


                                saveDocuments();

                            }
                            else {

                                documentItem.name =
                                    oldName;

                            }


                            renderDocuments();

                        }


                        inputRename.onkeydown =
                            function (
                                event
                            ) {

                                if (
                                    event.key ===
                                    "Enter"
                                ) {

                                    event.preventDefault();


                                    finishRename(
                                        true
                                    );

                                }


                                if (
                                    event.key ===
                                    "Escape"
                                ) {

                                    event.preventDefault();


                                    finishRename(
                                        false
                                    );

                                }

                            };

                    };

            }


            // =================================================
            // MOVE DOCUMENT UP
            // =================================================

            const moveUpButton =
                options.querySelector(
                    ".move-document-up"
                );


            if (moveUpButton) {

                moveUpButton.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();
                        e.stopPropagation();


                        if (
                            index <=
                            0
                        ) {

                            return;

                        }


                        const previousDocument =
                            projectDocuments[
                                index - 1
                            ];


                        const currentOrder =
                            documentItem.order;


                        documentItem.order =
                            previousDocument.order;


                        previousDocument.order =
                            currentOrder;


                        documentItem.updatedAt =
                            new Date()
                                .toISOString();


                        previousDocument.updatedAt =
                            new Date()
                                .toISOString();


                        saveDocuments();


                        options.classList.remove(
                            "open"
                        );


                        renderDocuments();

                    };

            }


            // =================================================
            // MOVE DOCUMENT DOWN
            // =================================================

            const moveDownButton =
                options.querySelector(
                    ".move-document-down"
                );


            if (moveDownButton) {

                moveDownButton.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();
                        e.stopPropagation();


                        if (
                            index >=
                            projectDocuments.length -
                            1
                        ) {

                            return;

                        }


                        const nextDocument =
                            projectDocuments[
                                index + 1
                            ];


                        const currentOrder =
                            documentItem.order;


                        documentItem.order =
                            nextDocument.order;


                        nextDocument.order =
                            currentOrder;


                        documentItem.updatedAt =
                            new Date()
                                .toISOString();


                        nextDocument.updatedAt =
                            new Date()
                                .toISOString();


                        saveDocuments();


                        options.classList.remove(
                            "open"
                        );


                        renderDocuments();

                    };

            }


            // =================================================
            // DELETE DOCUMENT
            // =================================================

            const deleteButton =
                options.querySelector(
                    ".delete-document"
                );


            if (deleteButton) {

                deleteButton.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();
                        e.stopPropagation();


                        options.classList.remove(
                            "open"
                        );


                        const oldConfirm =
                            document.querySelector(
                                ".document-delete-confirm"
                            );


                        if (oldConfirm) {

                            oldConfirm.remove();

                        }


                        const confirmBox =
                            document.createElement(
                                "div"
                            );


                        confirmBox.className =
                            "document-delete-confirm";


                        confirmBox.innerHTML = `
                            <div class="document-delete-dialog">

                                <div class="document-delete-message">
                                    هل تريد حذف المستند؟
                                </div>

                                <div class="document-delete-name">
                                    ${documentItem.name}
                                </div>

                                <div class="document-delete-buttons">

                                    <button
                                        type="button"
                                        class="confirm-document-delete">
                                        حذف
                                    </button>

                                    <button
                                        type="button"
                                        class="cancel-document-delete">
                                        إلغاء
                                    </button>

                                </div>

                            </div>
                        `;


                        document.body.appendChild(
                            confirmBox
                        );


                        const confirmDelete =
                            confirmBox.querySelector(
                                ".confirm-document-delete"
                            );


                        if (confirmDelete) {

                            confirmDelete.onclick =
                                async function () {

                                    documents =
                                        documents.filter(
                                            function (
                                                doc
                                            ) {

                                                return (
                                                    String(
                                                        doc.id
                                                    ) !==
                                                    String(
                                                        documentItem.id
                                                    )
                                                );

                                            }
                                        );


                                    if (
                                        currentProject &&
                                        Array.isArray(
                                            currentProject.documents
                                        )
                                    ) {

                                        currentProject.documents =
                                            currentProject.documents.filter(
                                                function (
                                                    id
                                                ) {

                                                    return (
                                                        String(
                                                            id
                                                        ) !==
                                                        String(
                                                            documentItem.id
                                                        )
                                                    );

                                                }
                                            );


                                        currentProject.updatedAt =
                                            new Date()
                                                .toISOString();


                                        saveProjects();

                                    }


                                    try {

                                        await deleteWorkingWordFile(
                                            documentItem.storageId
                                        );

                                    }
                                    catch (
                                        storageError
                                    ) {

                                        console.warn(
                                            "تعذر حذف نسخة العمل:",
                                            storageError
                                        );

                                    }


                                    if (
                                        currentDocument &&
                                        String(
                                            currentDocument.id
                                        ) ===
                                        String(
                                            documentItem.id
                                        )
                                    ) {

                                        currentDocument =
                                            null;


                                        currentCitationSources =
                                            [];


                                        if (
                                            documentTitle
                                        ) {

                                            documentTitle.textContent =
                                                "لا يوجد مستند مفتوح";

                                        }

                                    }


                                    const remaining =
                                        getProjectDocuments(
                                            currentProject
                                                ? currentProject.id
                                                : null
                                        );


                                    remaining.forEach(
                                        function (
                                            doc,
                                            newIndex
                                        ) {

                                            doc.order =
                                                newIndex + 1;

                                        }
                                    );


                                    oramaRetrievalDb =
                                        null;


                                    oramaRetrievalCacheKey =
                                        "";


                                    oramaRetrievalDocumentId =
                                        null;


                                    saveDocuments();


                                    confirmBox.remove();


                                    renderDocuments();

                                };

                        }


                        const cancelDelete =
                            confirmBox.querySelector(
                                ".cancel-document-delete"
                            );


                        if (cancelDelete) {

                            cancelDelete.onclick =
                                function () {

                                    confirmBox.remove();

                                };

                        }


                        confirmBox.onclick =
                            function (
                                event
                            ) {

                                if (
                                    event.target ===
                                    confirmBox
                                ) {

                                    confirmBox.remove();

                                }

                            };

                    };

            }


            item.appendChild(
                title
            );


            item.appendChild(
                status
            );


            item.appendChild(
                menuButton
            );


            item.appendChild(
                options
            );


            documentsList.appendChild(
                item
            );

        }
    );

}


// =====================================================
// RENDER PROJECTS
// =====================================================

function renderProjects() {

    if (!projectsList) {

        return;

    }


    projectsList.innerHTML =
        "";


    if (
        projects.length ===
        0
    ) {

        projectsList.innerHTML = `
            <div class="empty-project">
                لا توجد مشاريع
            </div>
        `;

        return;

    }


    projects.forEach(
        function (
            project
        ) {

            const item =
                document.createElement(
                    "div"
                );


            item.className =
                "project-item";


            item.innerHTML = `
                <span class="project-title">
                    ${projectIcon}
                    ${project.name}
                </span>

                <button
                    class="project-menu"
                    type="button">
                    ⋮
                </button>

                <div class="project-options-menu">

                    <div class="rename-project">
                        ✏ إعادة تسمية
                    </div>

                    <div class="delete-project">
                        🗑 حذف
                    </div>

                </div>
            `;


            item.onclick =
                function (
                    e
                ) {

                    if (
                        e.target.closest(
                            ".project-menu"
                        ) ||
                        e.target.closest(
                            ".project-options-menu"
                        )
                    ) {

                        return;

                    }


                    e.stopPropagation();


                    setCurrentProject(
                        project
                    );


                    if (projectsPopup) {

                        projectsPopup.classList.remove(
                            "open"
                        );

                    }

                };


            const menu =
                item.querySelector(
                    ".project-menu"
                );


            if (menu) {

                menu.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();
                        e.stopPropagation();


                        document
                            .querySelectorAll(
                                ".project-options-menu.open"
                            )
                            .forEach(
                                function (
                                    menuItem
                                ) {

                                    menuItem.classList.remove(
                                        "open"
                                    );

                                }
                            );


                        const options =
                            item.querySelector(
                                ".project-options-menu"
                            );


                        if (!options) {

                            return;

                        }


                        options.classList.add(
                            "open"
                        );


                        const rect =
                            menu.getBoundingClientRect();


                        const menuWidth =
                            140;


                        const menuHeight =
                            options.offsetHeight ||
                            80;


                        const margin =
                            8;


                        let left =
                            rect.left -
                            menuWidth -
                            margin;


                        let top =
                            rect.bottom +
                            margin;


                        if (
                            left <
                            margin
                        ) {

                            left =
                                rect.right +
                                margin;

                        }


                        if (
                            left +
                            menuWidth >
                            window.innerWidth -
                            margin
                        ) {

                            left =
                                window.innerWidth -
                                menuWidth -
                                margin;

                        }


                        if (
                            top +
                            menuHeight >
                            window.innerHeight -
                            margin
                        ) {

                            top =
                                rect.top -
                                menuHeight -
                                margin;

                        }


                        if (
                            top <
                            margin
                        ) {

                            top =
                                margin;

                        }


                        options.style.position =
                            "fixed";


                        options.style.left =
                            left +
                            "px";


                        options.style.top =
                            top +
                            "px";


                        options.style.right =
                            "auto";


                        options.style.bottom =
                            "auto";


                        options.style.zIndex =
                            "999999";

                    };

            }


            // =================================================
            // RENAME PROJECT
            // =================================================

            const renameProject =
                item.querySelector(
                    ".rename-project"
                );


            if (renameProject) {

                renameProject.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();
                        e.stopPropagation();


                        const options =
                            item.querySelector(
                                ".project-options-menu"
                            );


                        if (options) {

                            options.classList.remove(
                                "open"
                            );

                        }


                        const titleElement =
                            item.querySelector(
                                ".project-title"
                            );


                        if (!titleElement) {

                            return;

                        }


                        const oldName =
                            project.name;


                        titleElement.innerHTML = `
                            <input
                                class="edit-project-title"
                                value="${oldName}">
                        `;


                        const edit =
                            titleElement.querySelector(
                                ".edit-project-title"
                            );


                        if (!edit) {

                            return;

                        }


                        edit.focus();


                        edit.setSelectionRange(
                            edit.value.length,
                            edit.value.length
                        );


                        edit.onkeydown =
                            function (
                                event
                            ) {

                                if (
                                    event.key ===
                                    "Enter"
                                ) {

                                    event.preventDefault();


                                    const value =
                                        edit.value.trim();


                                    project.name =
                                        value ||
                                        oldName;


                                    project.updatedAt =
                                        new Date()
                                            .toISOString();


                                    saveProjects();


                                    renderProjects();


                                    renderExpandedProjects();

                                }


                                if (
                                    event.key ===
                                    "Escape"
                                ) {

                                    event.preventDefault();


                                    project.name =
                                        oldName;


                                    renderProjects();

                                }

                            };

                    };

            }


            // =================================================
            // DELETE PROJECT
            // =================================================

            const deleteProject =
                item.querySelector(
                    ".delete-project"
                );


            if (deleteProject) {

                deleteProject.onclick =
                    function (
                        e
                    ) {

                        e.preventDefault();
                        e.stopPropagation();


                        const options =
                            item.querySelector(
                                ".project-options-menu"
                            );


                        if (options) {

                            options.classList.remove(
                                "open"
                            );

                        }


                        const confirmBox =
                            document.createElement(
                                "div"
                            );


                        confirmBox.className =
                            "project-delete-confirm";


                        confirmBox.innerHTML = `
                            <div class="confirm-dialog">

                                <p>
                                    هل تريد حذف المشروع:
                                    <br>

                                    <strong>
                                        ${project.name}
                                    </strong>

                                    ؟
                                </p>

                                <button
                                    class="confirm-project-delete"
                                    type="button">
                                    حذف
                                </button>

                                <button
                                    class="cancel-project-delete"
                                    type="button">
                                    إلغاء
                                </button>

                            </div>
                        `;


                        document.body.appendChild(
                            confirmBox
                        );


                        const confirmDelete =
                            confirmBox.querySelector(
                                ".confirm-project-delete"
                            );


                        if (confirmDelete) {

                            confirmDelete.onclick =
                                async function () {

                                    const projectDocumentIds =
                                        Array.isArray(
                                            project.documents
                                        )
                                            ? project.documents
                                            : [];


                                    for (
                                        let i = 0;
                                        i <
                                        projectDocumentIds.length;
                                        i++
                                    ) {

                                        const documentItem =
                                            documents.find(
                                                function (
                                                    d
                                                ) {

                                                    return (
                                                        String(
                                                            d.id
                                                        ) ===
                                                        String(
                                                            projectDocumentIds[i]
                                                        )
                                                    );

                                                }
                                            );


                                        if (
                                            documentItem
                                        ) {

                                            try {

                                                await deleteWorkingWordFile(
                                                    documentItem.storageId
                                                );

                                            }
                                            catch (
                                                error
                                            ) {

                                                console.warn(
                                                    "تعذر حذف نسخة العمل:",
                                                    error
                                                );

                                            }

                                        }

                                    }


                                    documents =
                                        documents.filter(
                                            function (
                                                doc
                                            ) {

                                                return (
                                                    !projectDocumentIds.some(
                                                        function (
                                                            id
                                                        ) {

                                                            return (
                                                                String(
                                                                    id
                                                                ) ===
                                                                String(
                                                                    doc.id
                                                                )
                                                            );

                                                        }
                                                    )
                                                );

                                            }
                                        );


                                    projects =
                                        projects.filter(
                                            function (
                                                p
                                            ) {

                                                return (
                                                    String(
                                                        p.id
                                                    ) !==
                                                    String(
                                                        project.id
                                                    )
                                                );

                                            }
                                        );


                                    if (
                                        currentProject &&
                                        String(
                                            currentProject.id
                                        ) ===
                                        String(
                                            project.id
                                        )
                                    ) {

                                        currentProject =
                                            null;


                                        currentDocument =
                                            null;


                                        currentCitationSources =
                                            [];


                                        if (
                                            documentTitle
                                        ) {

                                            documentTitle.textContent =
                                                "لا يوجد مستند مفتوح";

                                        }

                                    }


                                    oramaRetrievalDb =
                                        null;


                                    oramaRetrievalCacheKey =
                                        "";


                                    oramaRetrievalDocumentId =
                                        null;


                                    saveDocuments();


                                    saveProjects();


                                    renderProjects();


                                    renderExpandedProjects();


                                    renderDocuments();


                                    confirmBox.remove();

                                };

                        }


                        const cancelDelete =
                            confirmBox.querySelector(
                                ".cancel-project-delete"
                            );


                        if (cancelDelete) {

                            cancelDelete.onclick =
                                function () {

                                    confirmBox.remove();

                                };

                        }

                    };

            }


            projectsList.appendChild(
                item
            );

        }
    );

}


// =====================================================
// NEW PROJECT
// =====================================================

if (newProjectBtn) {

    newProjectBtn.onclick =
        function (
            e
        ) {

            e.preventDefault();
            e.stopPropagation();


            const oldBox =
                document.querySelector(
                    ".project-create-box"
                );


            if (oldBox) {

                oldBox.remove();

            }


            const box =
                document.createElement(
                    "div"
                );


            box.className =
                "project-create-box";


            box.innerHTML = `
                <div class="rename-dialog">

                    <input
                        class="new-project-name"
                        placeholder="اسم المشروع">

                    <button
                        class="save-project"
                        type="button">
                        حفظ
                    </button>

                    <button
                        class="cancel-project"
                        type="button">
                        إلغاء
                    </button>

                </div>
            `;


            document.body.appendChild(
                box
            );


            const buttonRect =
                newProjectBtn.getBoundingClientRect();


            const screenMargin =
                12;


            let left =
                buttonRect.left;


            let top =
                buttonRect.bottom +
                8;


            const actualBoxWidth =
                box.offsetWidth ||
                240;


            const boxHeight =
                box.offsetHeight ||
                120;


            if (
                left +
                actualBoxWidth >
                window.innerWidth -
                screenMargin
            ) {

                left =
                    window.innerWidth -
                    actualBoxWidth -
                    screenMargin;

            }


            if (
                left <
                screenMargin
            ) {

                left =
                    screenMargin;

            }


            if (
                top +
                boxHeight >
                window.innerHeight -
                screenMargin
            ) {

                top =
                    buttonRect.top -
                    boxHeight -
                    8;

            }


            box.style.position =
                "fixed";


            box.style.left =
                left +
                "px";


            box.style.top =
                top +
                "px";


            box.style.zIndex =
                "999999";


            const inputProject =
                box.querySelector(
                    ".new-project-name"
                );


            if (inputProject) {

                inputProject.focus();

            }


            const saveProjectButton =
                box.querySelector(
                    ".save-project"
                );


            if (saveProjectButton) {

                saveProjectButton.onclick =
                    function () {

                        const name =
                            inputProject
                                ? inputProject.value.trim()
                                : "";


                        if (
                            name
                        ) {

                            const now =
                                new Date()
                                    .toISOString();


                            const newProject = {

                                id:
                                    Date.now(),

                                name:
                                    name,

                                createdAt:
                                    now,

                                updatedAt:
                                    now,

                                documents:
                                    [],

                                references:
                                    [],

                                chatIds:
                                    [],

                                settings: {

                                    citationStyle:
                                        "",

                                    notes:
                                        ""

                                }

                            };


                            projects.unshift(
                                newProject
                            );


                            saveProjects();


                            renderProjects();


                            renderExpandedProjects();

                        }


                        box.remove();

                    };

            }


            const cancelProject =
                box.querySelector(
                    ".cancel-project"
                );


            if (cancelProject) {

                cancelProject.onclick =
                    function () {

                        box.remove();

                    };

            }


            box.onclick =
                function (
                    event
                ) {

                    event.stopPropagation();

                };

        };

}


// =====================================================
// PROJECTS BUTTON
// =====================================================

if (projectsBtn) {

    projectsBtn.onclick =
        function (
            e
        ) {

            e.preventDefault();
            e.stopPropagation();


            if (!projectsPopup) {

                return;

            }


            if (chatPopup) {

                chatPopup.classList.remove(
                    "open"
                );

            }


            if (searchPopup) {

                searchPopup.classList.remove(
                    "open"
                );

            }


            if (settingsWindow) {

                settingsWindow.classList.remove(
                    "open"
                );

            }


            projectsPopup.classList.toggle(
                "open"
            );


            renderProjects();

        };

}


// =====================================================
// EXPANDED SIDEBAR
// =====================================================

if (
    sidebarToggleBtn &&
    expandedSidebar &&
    expandedSidebarToggleSlot
) {

    sidebarToggleBtn.onclick =
        function (
            e
        ) {

            e.preventDefault();
            e.stopPropagation();


            const isOpening =
                !expandedSidebar.classList.contains(
                    "open"
                );


            if (isOpening) {

                expandedSidebar.classList.add(
                    "open"
                );


                document.body.classList.add(
                    "expanded-sidebar-open"
                );


                sidebarToggleBtn.title =
                    "إخفاء القائمة";


                sidebarToggleBtn.classList.add(
                    "sidebar-open"
                );


                expandedSidebarToggleSlot.appendChild(
                    sidebarToggleBtn
                );

            }
            else {

                expandedSidebar.classList.remove(
                    "open"
                );


                document.body.classList.remove(
                    "expanded-sidebar-open"
                );


                sidebarToggleBtn.title =
                    "إظهار القائمة";


                sidebarToggleBtn.classList.remove(
                    "sidebar-open"
                );


                if (
                    sidebarTogglePlaceholder.parentNode
                ) {

                    sidebarTogglePlaceholder.parentNode.insertBefore(
                        sidebarToggleBtn,
                        sidebarTogglePlaceholder.nextSibling
                    );

                }

            }

        };

}


// =====================================================
// EXPANDED PROJECTS
// =====================================================

function renderExpandedProjects() {

    const list =
        document.getElementById(
            "expanded-projects-list"
        );


    if (!list) {

        return;

    }


    list.innerHTML =
        "";


    projects.forEach(
        function (
            project
        ) {

            const item =
                document.createElement(
                    "div"
                );


            item.className =
                "expanded-project-item";


            item.innerHTML = `
                <span>
                    ${projectIcon}
                    ${project.name}
                </span>
            `;


            item.onclick =
                function (
                    e
                ) {

                    e.stopPropagation();


                    setCurrentProject(
                        project
                    );

                };


            list.appendChild(
                item
            );

        }
    );

}


// =====================================================
// SIDEBAR RECENT CHATS
// =====================================================

function renderSidebarChats() {

    const list =
        document.getElementById(
            "new-chat-list"
        );


    if (!list) {

        return;

    }


    list.innerHTML =
        "";


    if (
        chats.length ===
        0
    ) {

        list.innerHTML = `
            <div class="empty-chat">
                لا توجد محادثات
            </div>
        `;


        return;

    }


    chats
        .slice(
            0,
            8
        )
        .forEach(
            function (
                chat
            ) {

                const item =
                    document.createElement(
                        "div"
                    );


                item.className =
                    "recent-chat-item";


                item.innerHTML = `
                    <span class="chat-title">
                        ${chatIcon}
                        ${chat.title}
                    </span>
                `;


                item.onclick =
                    function (
                        e
                    ) {

                        e.stopPropagation();


                        currentChat =
                            chat;


                        renderChat();


                        if (projectsPopup) {

                            projectsPopup.classList.remove(
                                "open"
                            );

                        }


                        if (chatPopup) {

                            chatPopup.classList.remove(
                                "open"
                            );

                        }


                        if (searchPopup) {

                            searchPopup.classList.remove(
                                "open"
                            );

                        }

                    };


                list.appendChild(
                    item
                );

            }
        );

}


// =====================================================
// CREATE NEW CHAT
// =====================================================

function createNewChat() {

    currentChat = {

        id:
            Date.now(),

        title:
            "محادثة جديدة",

        messages:
            [],

        isTemporary:
            true,

        projectId:
            currentProject
                ? currentProject.id
                : null

    };


    currentCitationSources =
        [];


    if (input) {

        input.value =
            "";


        input.style.height =
            "auto";

    }


    renderSidebarChats();


    renderRecentChats();


    renderChat();

}


// =====================================================
// NEW CHAT BUTTON
// =====================================================

if (newChatBtn) {

    newChatBtn.onclick =
        function (
            e
        ) {

            e.preventDefault();
            e.stopPropagation();


            if (projectsPopup) {

                projectsPopup.classList.remove(
                    "open"
                );

            }


            if (chatPopup) {

                chatPopup.classList.remove(
                    "open"
                );

            }


            if (searchPopup) {

                searchPopup.classList.remove(
                    "open"
                );

            }


            if (settingsWindow) {

                settingsWindow.classList.remove(
                    "open"
                );

            }


            createNewChat();

        };

}


// =====================================================
// INITIALIZE SIDEBAR SECTIONS
// =====================================================

function initializeSidebarSections() {

    const headers =
        document.querySelectorAll(
            ".section-title[data-target], .section-toggle[data-target]"
        );


    headers.forEach(
        function (
            header
        ) {

            const targetId =
                header.getAttribute(
                    "data-target"
                );


            if (!targetId) {

                return;

            }


            const target =
                document.getElementById(
                    targetId
                );


            if (!target) {

                return;

            }


            target.classList.remove(
                "open"
            );


            header.classList.remove(
                "open"
            );


            header.onclick =
                function (
                    e
                ) {

                    e.preventDefault();
                    e.stopPropagation();


                    const isOpen =
                        target.classList.contains(
                            "open"
                        );


                    target.classList.toggle(
                        "open",
                        !isOpen
                    );


                    header.classList.toggle(
                        "open",
                        !isOpen
                    );

                };

        }
    );

}


// =====================================================
// CHAT STORAGE
// =====================================================

function ensureCurrentChatObject() {

    if (
        currentChat
    ) {

        return;

    }


    createNewChat();

}


// =====================================================
// END PART 2
// الجزء الثالث يتابع:
// الإعدادات + مزودو الذكاء الاصطناعي + البحث + الاسترجاع
// =====================================================
// =====================================================
// PART 3 / 4
// الإعدادات + مزودو الذكاء الاصطناعي
// + محرك البحث الجديد + الاسترجاع
// =====================================================


// =====================================================
// SETTINGS
// =====================================================

function loadSettings() {

    const data =
        getSavedSettings();


    if (
        providerSelect &&
        data.provider
    ) {

        providerSelect.value =
            data.provider;

    }


    if (
        modelSelect &&
        data.model
    ) {

        modelSelect.value =
            data.model;

    }


    if (
        apiKeyInput &&
        data.key
    ) {

        apiKeyInput.value =
            data.key;

    }


    if (
        showKeyCheckbox &&
        data.showKey
    ) {

        showKeyCheckbox.checked =
            true;

    }


    updateAPIKeyVisibility();


    updateProviderUI();

}


// =====================================================
// SAVE SETTINGS
// =====================================================

function saveSettings() {

    const provider =
        providerSelect
            ? providerSelect.value
            : "openrouter";


    const model =
        modelSelect
            ? modelSelect.value.trim()
            : "";


    const key =
        apiKeyInput
            ? apiKeyInput.value.trim()
            : "";


    const showKey =
        showKeyCheckbox
            ? showKeyCheckbox.checked
            : false;


    const settings = {

        provider:
            provider,

        model:
            model,

        key:
            key,

        showKey:
            showKey

    };


    localStorage.setItem(
        "AI_SETTINGS",
        JSON.stringify(
            settings
        )
    );


    updateAPIKeyVisibility();


    updateProviderUI();

}


// =====================================================
// API KEY VISIBILITY
// =====================================================

function updateAPIKeyVisibility() {

    if (
        !apiKeyInput ||
        !showKeyCheckbox
    ) {

        return;

    }


    apiKeyInput.type =
        showKeyCheckbox.checked
            ? "text"
            : "password";

}


// =====================================================
// PROVIDER UI
// =====================================================

function updateProviderUI() {

    const provider =
        providerSelect
            ? String(
                providerSelect.value ||
                ""
            ).toLowerCase()
            : "";


    if (
        providerLabel
    ) {

        if (
            provider ===
            "gemini"
        ) {

            providerLabel.textContent =
                "Gemini";

        }
        else if (
            provider ===
            "openai"
        ) {

            providerLabel.textContent =
                "OpenAI";

        }
        else if (
            provider ===
            "groq"
        ) {

            providerLabel.textContent =
                "Groq";

        }
        else {

            providerLabel.textContent =
                "OpenRouter";

        }

    }


    if (
        modelSelect
    ) {

        modelSelect.disabled =
            false;

    }

}


// =====================================================
// SETTINGS BUTTON
// =====================================================

if (
    settingsBtn
) {

    settingsBtn.onclick =
        function (
            e
        ) {

            e.preventDefault();
            e.stopPropagation();


            if (
                settingsWindow
            ) {

                settingsWindow.classList.toggle(
                    "open"
                );

            }


            if (
                projectsPopup
            ) {

                projectsPopup.classList.remove(
                    "open"
                );

            }


            if (
                chatPopup
            ) {

                chatPopup.classList.remove(
                    "open"
                );

            }


            if (
                searchPopup
            ) {

                searchPopup.classList.remove(
                    "open"
                );

            }


            loadSettings();

        };

}


// =====================================================
// CLOSE SETTINGS
// =====================================================

if (
    closeSettingsBtn
) {

    closeSettingsBtn.onclick =
        function (
            e
        ) {

            e.preventDefault();
            e.stopPropagation();


            if (
                settingsWindow
            ) {

                settingsWindow.classList.remove(
                    "open"
                );

            }

        };

}


// =====================================================
// SAVE SETTINGS BUTTON
// =====================================================

if (
    saveSettingsBtn
) {

    saveSettingsBtn.onclick =
        function (
            e
        ) {

            e.preventDefault();
            e.stopPropagation();


            saveSettings();


            if (
                settingsWindow
            ) {

                settingsWindow.classList.remove(
                    "open"
                );

            }

        };

}


// =====================================================
// PROVIDER CHANGE
// =====================================================

if (
    providerSelect
) {

    providerSelect.onchange =
        async function () {

            updateProviderUI();


            const provider =
                String(
                    providerSelect.value ||
                    ""
                ).toLowerCase();


            if (
                modelSelect
            ) {

                modelSelect.innerHTML = `
                    <option value="">
                        جاري تحميل النماذج...
                    </option>
                `;

            }


            try {

                await LoadModels(
                    provider
                );

            }
            catch (
                error
            ) {

                console.warn(
                    "تعذر تحميل النماذج:",
                    error
                );


                if (
                    modelSelect
                ) {

                    modelSelect.innerHTML = `
                        <option value="">
                            تعذر تحميل النماذج
                        </option>
                    `;

                }

            }

        };

}


// =====================================================
// REFRESH MODELS
// =====================================================

if (
    refreshModelsBtn
) {

    refreshModelsBtn.onclick =
        async function (
            e
        ) {

            e.preventDefault();
            e.stopPropagation();


            const provider =
                providerSelect
                    ? String(
                        providerSelect.value ||
                        ""
                    ).toLowerCase()
                    : "openrouter";


            try {

                refreshModelsBtn.disabled =
                    true;


                if (
                    modelStatus
                ) {

                    modelStatus.textContent =
                        "جاري تحميل النماذج...";

                }


                await LoadModels(
                    provider
                );


                if (
                    modelStatus
                ) {

                    modelStatus.textContent =
                        "تم تحديث النماذج";

                }

            }
            catch (
                error
            ) {

                if (
                    modelStatus
                ) {

                    modelStatus.textContent =
                        error &&
                        error.message
                            ? error.message
                            : "فشل تحديث النماذج";

                }

            }
            finally {

                refreshModelsBtn.disabled =
                    false;

            }

        };

}


// =====================================================
// LOAD MODELS
// =====================================================

async function LoadModels(
    provider
) {

    const cleanProvider =
        String(
            provider ||
            ""
        ).toLowerCase();


    if (
        cleanProvider ===
        "gemini"
    ) {

        await loadGeminiModels();

        return;

    }


    if (
        cleanProvider ===
        "openrouter"
    ) {

        await loadOpenRouterModels();

        return;

    }


    if (
        cleanProvider ===
        "groq"
    ) {

        await loadGroqModels();

        return;

    }


    if (
        cleanProvider ===
        "openai"
    ) {

        await loadOpenAIModels();

        return;

    }


    throw new Error(
        "مزود غير مدعوم."
    );

}


// =====================================================
// GEMINI MODELS
// =====================================================

async function loadGeminiModels() {

    const data =
        getSavedSettings();


    const key =
        data.key ||
        "";


    if (
        !key.trim()
    ) {

        throw new Error(
            "أدخل مفتاح Gemini أولًا."
        );

    }


    const response =
        await fetch(
            "https://generativelanguage.googleapis.com/v1beta/models",
            {

                method:
                    "GET",

                headers: {

                    "x-goog-api-key":
                        key

                }

            }
        );


    const result =
        await readJSON(
            response
        );


    if (
        !response.ok
    ) {

        throw new Error(
            getAPIError(
                result,
                "فشل تحميل نماذج Gemini."
            )
        );

    }


    const models =
        Array.isArray(
            result.models
        )
            ? result.models
            : [];


    const usableModels =
        models.filter(
            function (
                item
            ) {

                return (
                    item &&
                    item.name &&
                    (
                        !item.supportedGenerationMethods ||
                        item.supportedGenerationMethods.includes(
                            "generateContent"
                        )
                    )
                );

            }
        );


    fillModelSelect(
        usableModels.map(
            function (
                item
            ) {

                return {

                    id:
                        normalizeGeminiModel(
                            item.name
                        ),

                    name:
                        item.displayName ||
                        normalizeGeminiModel(
                            item.name
                        )

                };

            }
        )
    );

}


// =====================================================
// OPENROUTER MODELS
// =====================================================

async function loadOpenRouterModels() {

    const data =
        getSavedSettings();


    const key =
        data.key ||
        "";


    const response =
        await fetch(
            "https://openrouter.ai/api/v1/models",
            {

                method:
                    "GET",

                headers: {

                    "Content-Type":
                        "application/json",

                    ...(key.trim()
                        ? {
                            "Authorization":
                                "Bearer " +
                                key
                        }
                        : {})

                }

            }
        );


    const result =
        await readJSON(
            response
        );


    if (
        !response.ok
    ) {

        throw new Error(
            getAPIError(
                result,
                "فشل تحميل نماذج OpenRouter."
            )
        );

    }


    const models =
        Array.isArray(
            result.data
        )
            ? result.data
            : [];


    fillModelSelect(
        models.map(
            function (
                item
            ) {

                return {

                    id:
                        item.id,

                    name:
                        item.name ||
                        item.id

                };

            }
        )
    );

}


// =====================================================
// GROQ MODELS
// =====================================================

async function loadGroqModels() {

    const data =
        getSavedSettings();


    const key =
        data.key ||
        "";


    if (
        !key.trim()
    ) {

        throw new Error(
            "أدخل مفتاح Groq أولًا."
        );

    }


    const response =
        await fetch(
            "https://api.groq.com/openai/v1/models",
            {

                method:
                    "GET",

                headers: {

                    "Authorization":
                        "Bearer " +
                        key

                }

            }
        );


    const result =
        await readJSON(
            response
        );


    if (
        !response.ok
    ) {

        throw new Error(
            getAPIError(
                result,
                "فشل تحميل نماذج Groq."
            )
        );

    }


    const models =
        Array.isArray(
            result.data
        )
            ? result.data
            : [];


    fillModelSelect(
        models.map(
            function (
                item
            ) {

                return {

                    id:
                        item.id,

                    name:
                        item.id

                };

            }
        )
    );

}


// =====================================================
// OPENAI MODELS
// =====================================================

async function loadOpenAIModels() {

    const data =
        getSavedSettings();


    const key =
        data.key ||
        "";


    if (
        !key.trim()
    ) {

        throw new Error(
            "أدخل مفتاح OpenAI أولًا."
        );

    }


    const response =
        await fetch(
            "https://api.openai.com/v1/models",
            {

                method:
                    "GET",

                headers: {

                    "Authorization":
                        "Bearer " +
                        key

                }

            }
        );


    const result =
        await readJSON(
            response
        );


    if (
        !response.ok
    ) {

        throw new Error(
            getAPIError(
                result,
                "فشل تحميل نماذج OpenAI."
            )
        );

    }


    const models =
        Array.isArray(
            result.data
        )
            ? result.data
            : [];


    fillModelSelect(
        models
            .filter(
                function (
                    item
                ) {

                    return (
                        item &&
                        item.id &&
                        (
                            item.id.includes(
                                "gpt"
                            )
                        )
                    );

                }
            )
            .map(
                function (
                    item
                ) {

                    return {

                        id:
                            item.id,

                        name:
                            item.id

                    };

                }
            )
    );

}


// =====================================================
// FILL MODEL SELECT
// =====================================================

function fillModelSelect(
    models
) {

    if (
        !modelSelect
    ) {

        return;

    }


    modelSelect.innerHTML =
        "";


    if (
        !Array.isArray(models) ||
        models.length ===
        0
    ) {

        modelSelect.innerHTML = `
            <option value="">
                لا توجد نماذج متاحة
            </option>
        `;

        return;

    }


    const savedModel =
        getSavedSettings().model ||
        "";


    models.forEach(
        function (
            model
        ) {

            if (
                !model ||
                !model.id
            ) {

                return;

            }


            const option =
                document.createElement(
                    "option"
                );


            option.value =
                model.id;


            option.textContent =
                model.name ||
                model.id;


            modelSelect.appendChild(
                option
            );

        }
    );


    if (
        savedModel &&
        Array.from(
            modelSelect.options
        ).some(
            function (
                option
            ) {

                return (
                    option.value ===
                    savedModel
                );

            }
        )
    ) {

        modelSelect.value =
            savedModel;

    }
    else {

        modelSelect.selectedIndex =
            0;

    }

}


// =====================================================
// TEST CONNECTION
// =====================================================

if (
    testConnectionBtn
) {

    testConnectionBtn.onclick =
        async function (
            e
        ) {

            e.preventDefault();
            e.stopPropagation();


            try {

                testConnectionBtn.disabled =
                    true;


                if (
                    modelStatus
                ) {

                    modelStatus.textContent =
                        "جاري اختبار الاتصال...";

                }


                saveSettings();


                await TestAIConnection();


                if (
                    modelStatus
                ) {

                    modelStatus.textContent =
                        "تم الاتصال بنجاح";

                }

            }
            catch (
                error
            ) {

                if (
                    modelStatus
                ) {

                    modelStatus.textContent =
                        error &&
                        error.message
                            ? error.message
                            : "فشل الاتصال";

                }

            }
            finally {

                testConnectionBtn.disabled =
                    false;

            }

        };

}


// =====================================================
// TEST AI CONNECTION
// =====================================================

async function TestAIConnection() {

    const data =
        getSavedSettings();


    const provider =
        String(
            data.provider ||
            ""
        ).toLowerCase();


    if (
        provider ===
        "gemini"
    ) {

        const model =
            normalizeGeminiModel(
                data.model
            );


        const response =
            await fetch(
                "https://generativelanguage.googleapis.com/v1beta/models/" +
                encodeURIComponent(
                    model
                ),
                {

                    headers: {

                        "x-goog-api-key":
                            data.key

                    }

                }
            );


        const result =
            await readJSON(
                response
            );


        if (
            !response.ok
        ) {

            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ Gemini."
                )
            );

        }


        return true;

    }


    if (
        provider ===
        "openrouter"
    ) {

        const response =
            await fetch(
                "https://openrouter.ai/api/v1/models",
                {

                    headers: {

                        "Authorization":
                            "Bearer " +
                            data.key

                    }

                }
            );


        if (
            !response.ok
        ) {

            const result =
                await readJSON(
                    response
                );


            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ OpenRouter."
                )
            );

        }


        return true;

    }


    if (
        provider ===
        "groq"
    ) {

        const response =
            await fetch(
                "https://api.groq.com/openai/v1/models",
                {

                    headers: {

                        "Authorization":
                            "Bearer " +
                            data.key

                    }

                }
            );


        if (
            !response.ok
        ) {

            const result =
                await readJSON(
                    response
                );


            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ Groq."
                )
            );

        }


        return true;

    }


    if (
        provider ===
        "openai"
    ) {

        const response =
            await fetch(
                "https://api.openai.com/v1/models",
                {

                    headers: {

                        "Authorization":
                            "Bearer " +
                            data.key

                    }

                }
            );


        if (
            !response.ok
        ) {

            const result =
                await readJSON(
                    response
                );


            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ OpenAI."
                )
            );

        }


        return true;

    }


    throw new Error(
        "مزود الذكاء الاصطناعي غير معروف."
    );

}


// =====================================================
// NEW SEARCH ENGINE
// محرك البحث الجديد باستخدام Orama
// =====================================================

let oramaRetrievalDb =
    null;


let oramaRetrievalCacheKey =
    "";


let oramaRetrievalDocumentId =
    null;


// =====================================================
// LOAD ORAMA MODULE
// تحميل محرك Orama بطريقة متوافقة مع Webpack
// =====================================================

async function getOramaModule() {

    if (
        window.__researchOramaModule
    ) {

        return window.__researchOramaModule;

    }


    try {

        if (
            typeof require !==
            "function"
        ) {

            throw new Error(
                "Webpack require غير متاح."
            );

        }


        const module =
            require(
                "@orama/orama"
            );


        if (
            !module
        ) {

            throw new Error(
                "تعذر تحميل @orama/orama."
            );

        }


        window.__researchOramaModule =
            module;


        return module;

    }
    catch (
        error
    ) {

        console.error(
            "تعذر تحميل محرك Orama:",
            error
        );


        return null;

    }

}


// =====================================================
// BUILD NEW SEARCH INDEX
// بناء فهرس البحث الجديد
// =====================================================

async function buildNewSearchIndex(
    documentItem
) {

    if (
        !documentItem
    ) {

        return null;

    }


    // ==================================
    // تحميل بنية المستند الفعلية
    // ==================================

    const structureData =
        await ensureDocumentStructure(
            documentItem
        );


    if (
        !structureData ||
        !Array.isArray(
            structureData.paragraphs
        )
    ) {

        console.warn(
            "لا توجد فقرات صالحة لبناء فهرس Orama."
        );


        return null;

    }


    const paragraphs =
        structureData.paragraphs;


    // ==================================
    // تحميل Orama
    // ==================================

    const module =
        await getOramaModule();


    if (
        !module
    ) {

        return null;

    }


    const create =
        module.create ||
        (
            module.default &&
            module.default.create
        );


    const insertMultiple =
        module.insertMultiple ||
        (
            module.default &&
            module.default.insertMultiple
        );


    if (
        typeof create !==
        "function" ||
        typeof insertMultiple !==
        "function"
    ) {

        throw new Error(
            "دوال Orama المطلوبة غير متاحة."
        );

    }


    // ==================================
    // إنشاء قاعدة البيانات
    // ==================================

    const db =
        create({

            schema: {

                id:
                    "string",

                text:
                    "string",

                heading:
                    "string",

                paragraphIndex:
                    "number",

                headingIndex:
                    "number",

                headingLevel:
                    "string",

                isHeading:
                    "boolean"

            },

            language:
                "arabic"

        });


    // ==================================
    // ترتيب العناوين
    // ==================================

    const headings =
        Array.isArray(
            structureData.headings
        )
            ? structureData.headings
                .filter(
                    function (
                        heading
                    ) {

                        return (
                            heading &&
                            typeof heading.index !==
                                "undefined" &&
                            String(
                                heading.text ||
                                ""
                            ).trim()
                        );

                    }
                )
                .sort(
                    function (
                        a,
                        b
                    ) {

                        return (
                            Number(
                                a.index
                            ) -
                            Number(
                                b.index
                            )
                        );

                    }
                )
            : [];


    // ==================================
    // خريطة العناوين
    // ==================================

    const headingMap =
        new Map();


    headings.forEach(
        function (
            heading
        ) {

            headingMap.set(
                Number(
                    heading.index
                ),
                heading
            );

        }
    );


    // ==================================
    // العثور على أقرب عنوان
    // ==================================

    function getParagraphHeading(
        paragraphIndex
    ) {

        // الفقرة نفسها عنوان
        if (
            headingMap.has(
                paragraphIndex
            )
        ) {

            return headingMap.get(
                paragraphIndex
            );

        }


        // أقرب عنوان سابق
        for (
            let i =
                headings.length - 1;

            i >= 0;

            i--
        ) {

            const heading =
                headings[i];


            if (
                Number(
                    heading.index
                ) <
                paragraphIndex
            ) {

                return heading;

            }

        }


        return null;

    }


    // ==================================
    // بناء سجلات Orama
    // ==================================

    const records =
        [];


    paragraphs.forEach(
        function (
            paragraph
        ) {

            if (
                !paragraph
            ) {

                return;

            }


            const text =
                String(
                    paragraph.text ||
                    ""
                ).trim();


            if (
                !text
            ) {

                return;

            }


            const paragraphIndex =
                Number(
                    paragraph.index
                );


            const heading =
                getParagraphHeading(
                    paragraphIndex
                );


            const isHeading =
                Boolean(
                    heading &&
                    Number(
                        heading.index
                    ) ===
                    paragraphIndex
                );


            records.push({

                id:
                    "p-" +
                    String(
                        paragraphIndex
                    ),

                text:
                    text,

                heading:
                    heading
                        ? String(
                            heading.text ||
                            ""
                        ).trim()
                        : "",

                paragraphIndex:
                    paragraphIndex,

                headingIndex:
                    heading
                        ? Number(
                            heading.index
                        )
                        : -1,

                headingLevel:
                    heading
                        ? String(
                            heading.style ||
                            ""
                        )
                        : "",

                isHeading:
                    isHeading

            });

        }
    );


    // ==================================
    // إدخال جميع السجلات دفعة واحدة
    // ==================================

    if (
        records.length >
        0
    ) {

        await insertMultiple(
            db,
            records
        );

    }


    // ==================================
    // بناء مفتاح الذاكرة
    // ==================================

    const cacheKey =
        [

            String(
                documentItem.id
            ),

            String(
                documentItem.indexUpdatedAt ||
                ""
            ),

            String(
                paragraphs.length
            ),

            String(
                headings.length
            )

        ].join(
            "|"
        );


    // ==================================
    // حفظ الفهرس في الذاكرة
    // ==================================

    oramaRetrievalDb =
        db;


    oramaRetrievalCacheKey =
        cacheKey;


    oramaRetrievalDocumentId =
        String(
            documentItem.id
        );


    console.log(
        "تم بناء فهرس Orama الجديد:",
        {

            documentId:
                documentItem.id,

            paragraphs:
                records.length,

            headings:
                headings.length

        }
    );


    return db;

}


// =====================================================
// ENSURE NEW SEARCH INDEX
// التأكد من وجود فهرس Orama حديث
// =====================================================

async function ensureNewSearchIndex(
    documentItem
) {

    if (
        !documentItem
    ) {

        return null;

    }


    const structureData =
        await ensureDocumentStructure(
            documentItem
        );


    if (
        !structureData
    ) {

        return null;

    }


    const paragraphs =
        Array.isArray(
            structureData.paragraphs
        )
            ? structureData.paragraphs
            : [];


    const headings =
        Array.isArray(
            structureData.headings
        )
            ? structureData.headings
            : [];


    const expectedCacheKey =
        [

            String(
                documentItem.id
            ),

            String(
                documentItem.indexUpdatedAt ||
                ""
            ),

            String(
                paragraphs.length
            ),

            String(
                headings.length
            )

        ].join(
            "|"
        );


    // ==================================
    // استخدام الفهرس الموجود
    // ==================================

    if (
        oramaRetrievalDb &&
        oramaRetrievalCacheKey ===
            expectedCacheKey &&
        String(
            oramaRetrievalDocumentId
        ) ===
        String(
            documentItem.id
        )
    ) {

        return oramaRetrievalDb;

    }


    // ==================================
    // إعادة البناء
    // ==================================

    return await buildNewSearchIndex(
        documentItem
    );

}


// =====================================================
// SEARCH DOCUMENT WITH NEW ENGINE
// البحث في المستند باستخدام Orama
// =====================================================

async function searchDocumentWithNewEngine(
    query,
    documentItem,
    limit
) {

    const cleanQuery =
        String(
            query ||
            ""
        ).trim();


    if (
        !cleanQuery ||
        !documentItem
    ) {

        return [];

    }


    // ==================================
    // التأكد من وجود الفهرس
    // ==================================

    const db =
        await ensureNewSearchIndex(
            documentItem
        );


    if (
        !db
    ) {

        return [];

    }


    // ==================================
    // تحميل دالة البحث
    // ==================================

    const module =
        await getOramaModule();


    if (
        !module
    ) {

        return [];

    }


    const search =
        module.search ||
        (
            module.default &&
            module.default.search
        );


    if (
        typeof search !==
        "function"
    ) {

        throw new Error(
            "دالة search في Orama غير متاحة."
        );

    }


    // ==================================
    // تنفيذ البحث
    // ==================================

    const result =
        await search(
            db,
            {

                term:
                    cleanQuery,

                properties: [

                    "text",

                    "heading"

                ],

                limit:
                    Math.max(
                        1,
                        Number(
                            limit ||
                            8
                        )
                    ),

                tolerance:
                    1

            }
        );


    if (
        !result ||
        !Array.isArray(
            result.hits
        )
    ) {

        return [];

    }


    // ==================================
    // تحويل النتائج
    // ==================================

    return result.hits.map(
        function (
            hit,
            index
        ) {

            const document =
                hit.document ||
                {};


            return {

                rank:
                    index + 1,

                paragraphIndex:
                    Number(
                        document.paragraphIndex ||
                        0
                    ),

                paragraphId:
                    document.id ||
                    "",

                heading:
                    String(
                        document.heading ||
                        ""
                    ),

                headingIndex:
                    Number(
                        document.headingIndex ||
                        -1
                    ),

                headingLevel:
                    String(
                        document.headingLevel ||
                        ""
                    ),

                isHeading:
                    Boolean(
                        document.isHeading
                    ),

                text:
                    String(
                        document.text ||
                        ""
                    ),

                score:
                    Number(
                        hit.score ||
                        0
                    )

            };

        }
    );

}


// =====================================================
// BUILD AI DOCUMENT CONTEXT
// بناء سياق المستند من نتائج Orama
// =====================================================

async function buildAIDocumentContext(
    question
) {

    if (
        !currentDocument
    ) {

        currentCitationSources =
            [];


        return {

            found:
                false,

            query:
                String(
                    question ||
                    ""
                ),

            text:
                "",

            sources:
                [],

            resultCount:
                0,

            selectedCount:
                0

        };

    }


    const cleanQuestion =
        String(
            question ||
            ""
        ).trim();


    if (
        !cleanQuestion
    ) {

        currentCitationSources =
            [];


        return {

            found:
                false,

            query:
                "",

            text:
                "",

            sources:
                [],

            resultCount:
                0,

            selectedCount:
                0

        };

    }


    // ==================================
    // تحديد نوع السؤال
    // ==================================

    const retrievalProfile =
        typeof getRetrievalProfile ===
            "function"
                ? getRetrievalProfile(
                    cleanQuestion
                )
                : {

                    type:
                        "general",

                    maxResults:
                        8,

                    maxChars:
                        8000

                };


    // ==================================
    // عدد النتائج
    // ==================================

    const requestedLimit =
        retrievalProfile &&
        typeof retrievalProfile.maxResults ===
            "number"
                ? retrievalProfile.maxResults
                : 8;


    const results =
        await searchDocumentWithNewEngine(
            cleanQuestion,
            currentDocument,
            Math.max(
                10,
                requestedLimit
            )
        );


    if (
        !results ||
        results.length ===
            0
    ) {

        currentCitationSources =
            [];


        return {

            found:
                false,

            query:
                cleanQuestion,

            profile:
                retrievalProfile.type,

            text:
                "",

            sources:
                [],

            resultCount:
                0,

            selectedCount:
                0

        };

    }


    // ==================================
    // منع التكرار
    // ==================================

    const selected =
        [];


    const selectedParagraphs =
        new Set();


    const maxResults =
        Math.max(
            1,
            requestedLimit
        );


    for (
        let i = 0;

        i <
        results.length &&
        selected.length <
            maxResults;

        i++
    ) {

        const item =
            results[i];


        if (
            !item
        ) {

            continue;

        }


        const paragraphKey =
            String(
                item.paragraphIndex
            );


        if (
            selectedParagraphs.has(
                paragraphKey
            )
        ) {

            continue;

        }


        selectedParagraphs.add(
            paragraphKey
        );


        selected.push(
            item
        );

    }


    // ==================================
    // بناء مصادر الإحالات
    // ==================================

    currentCitationSources =
        selected.map(
            function (
                item,
                index
            ) {

                return {

                    rank:
                        index + 1,

                    paragraphIndex:
                        item.paragraphIndex,

                    paragraphId:
                        item.paragraphId ||
                        "",

                    heading:
                        item.heading ||
                        "",

                    mainParagraph:
                        item.text ||
                        "",

                    text:
                        item.text ||
                        "",

                    score:
                        item.score ||
                        0

                };

            }
        );


    // ==================================
    // بناء سياق النموذج
    // ==================================

    const textParts =
        [];


    selected.forEach(
        function (
            item,
            index
        ) {

            const block =
                [

                    "[مقطع " +
                    (
                        index + 1
                    ) +
                    "]",

                    item.heading
                        ? "العنوان: " +
                          item.heading
                        : "",

                    "المقطع: " +
                    String(
                        item.text ||
                        ""
                    ).trim()

                ]
                .filter(
                    function (
                        value
                    ) {

                        return Boolean(
                            String(
                                value ||
                                ""
                            ).trim()
                        );

                    }
                )
                .join(
                    "\n"
                );


            if (
                block
            ) {

                textParts.push(
                    block
                );

            }

        }
    );


    const contextText =
        textParts.join(
            "\n\n---\n\n"
        );


    return {

        found:
            selected.length >
            0,

        query:
            cleanQuestion,

        profile:
            retrievalProfile.type,

        text:
            contextText,

        sources:
            currentCitationSources,

        resultCount:
            results.length,

        selectedCount:
            selected.length,

        totalOccurrences:
            results.length,

        matchedFamilies:
            [],

        matchedTerms:
            []

    };

}


// =====================================================
// SEARCH POPUP
// =====================================================

if (
    searchBtn
) {

    searchBtn.onclick =
        function (
            e
        ) {

            e.preventDefault();
            e.stopPropagation();


            if (
                searchPopup
            ) {

                searchPopup.classList.toggle(
                    "open"
                );

            }


            if (
                projectsPopup
            ) {

                projectsPopup.classList.remove(
                    "open"
                );

            }


            if (
                chatPopup
            ) {

                chatPopup.classList.remove(
                    "open"
                );

            }


            if (
                settingsWindow
            ) {

                settingsWindow.classList.remove(
                    "open"
                );

            }


            if (
                searchInput
            ) {

                searchInput.focus();

            }

        };

}


// =====================================================
// SEARCH INPUT
// =====================================================

async function executeDocumentSearch() {

    if (
        !searchInput ||
        !currentDocument
    ) {

        return;

    }


    const query =
        searchInput.value.trim();


    if (
        !query
    ) {

        return;

    }


    if (
        searchStatus
    ) {

        searchStatus.textContent =
            "جاري البحث...";

    }


    try {

        const results =
            await searchDocumentWithNewEngine(
                query,
                currentDocument,
                20
            );


        if (
            searchResultsList
        ) {

            searchResultsList.innerHTML =
                "";

        }


        if (
            results.length ===
            0
        ) {

            if (
                searchStatus
            ) {

                searchStatus.textContent =
                    "لا توجد نتائج.";

            }


            return;

        }


        results.forEach(
            function (
                result
            ) {

                const item =
                    document.createElement(
                        "div"
                    );


                item.className =
                    "search-result-item";


                item.innerHTML = `

                    <div class="search-result-heading">
                        ${result.heading || ""}
                    </div>

                    <div class="search-result-text">
                        ${result.text || ""}
                    </div>

                    <div class="search-result-meta">
                        مقطع ${result.rank}
                    </div>

                `;


                item.onclick =
                    function () {

                        currentCitationSources =
                            results;


                        if (
                            chatArea
                        ) {

                            chatArea.scrollTop =
                                chatArea.scrollHeight;

                        }

                    };


                if (
                    searchResultsList
                ) {

                    searchResultsList.appendChild(
                        item
                    );

                }

            }
        );


        if (
            searchStatus
        ) {

            searchStatus.textContent =
                "تم العثور على " +
                results.length +
                " نتيجة.";

        }

    }
    catch (
        error
    ) {

        if (
            searchStatus
        ) {

            searchStatus.textContent =
                error &&
                error.message
                    ? error.message
                    : "فشل البحث.";

        }

    }

}


if (
    searchInput
) {

    searchInput.onkeydown =
        function (
            e
        ) {

            if (
                e.key ===
                "Enter"
            ) {

                e.preventDefault();


                executeDocumentSearch();

            }

        };

}


// =====================================================
// END PART 3
// الجزء الرابع:
// محركات البث + sendMessage + التهيئة النهائية
// =====================================================
// =====================================================
// PART 4 / 4
// محركات البث + إرسال الرسالة + التهيئة النهائية
// =====================================================


// =====================================================
// STREAM GEMINI AI
// =====================================================

async function streamGeminiAI(
    text,
    onChunk
) {

    const data =
        getSavedSettings();


    const key =
        data.key ||
        "";


    const model =
        data.model ||
        "";


    if (
        !key.trim()
    ) {

        throw new Error(
            "لم يتم إدخال مفتاح Gemini من الإعدادات."
        );

    }


    if (
        !model.trim()
    ) {

        throw new Error(
            "لم يتم تحديد نموذج Gemini من الإعدادات."
        );

    }


    const documentContext =
        await buildAIDocumentContext(
            text
        );


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

    ].join(
        "\n"
    );


    const conversationMessages =
        [];


    const historyLimit =
        documentContext &&
        documentContext.found
            ? 2
            : 4;


    if (
        currentChat &&
        Array.isArray(
            currentChat.messages
        )
    ) {

        const messagesWithoutCurrent =
            currentChat.messages.slice(
                0,
                -1
            );


        const previousMessages =
            messagesWithoutCurrent.slice(
                -historyLimit
            );


        previousMessages.forEach(
            function (
                msg
            ) {

                if (
                    !msg ||
                    !msg.text
                ) {

                    return;

                }


                let messageText =
                    String(
                        msg.text
                    ).trim();


                const maxHistoryChars =
                    documentContext &&
                    documentContext.found
                        ? 1000
                        : 1500;


                if (
                    messageText.length >
                    maxHistoryChars
                ) {

                    messageText =
                        messageText.substring(
                            0,
                            maxHistoryChars
                        ) +
                        "…";

                }


                conversationMessages.push({

                    role:
                        msg.role ===
                        "ai"
                            ? "model"
                            : "user",

                    parts: [

                        {
                            text:
                                messageText
                        }

                    ]

                });

            }
        );

    }


    let userContent =
        text;


    if (
        documentContext &&
        documentContext.found
    ) {

        userContent =
            [

                "=== سؤال المستخدم ===",

                text,

                "",

                "=== اسم المستند ===",

                (
                    currentDocument
                        ? currentDocument.name
                        : ""
                ),

                "",

                "=== المادة المستخرجة من المستند ===",

                documentContext.text

            ].join(
                "\n"
            );

    }


    conversationMessages.push({

        role:
            "user",

        parts: [

            {
                text:
                    userContent
            }

        ]

    });


    const cleanModel =
        normalizeGeminiModel(
            model
        );


    const url =
        "https://generativelanguage.googleapis.com/v1beta/models/" +
        encodeURIComponent(
            cleanModel
        ) +
        ":streamGenerateContent?alt=sse";


    const response =
        await fetch(
            url,
            {

                method:
                    "POST",

                headers: {

                    "Content-Type":
                        "application/json",

                    "x-goog-api-key":
                        key

                },

                body:
                    JSON.stringify({

                        systemInstruction: {

                            parts: [

                                {
                                    text:
                                        systemInstruction
                                }

                            ]

                        },

                        contents:
                            conversationMessages,

                        generationConfig: {

                            temperature:
                                0.2

                        }

                    })

            }
        );


    if (
        !response.ok
    ) {

        const result =
            await readJSON(
                response
            );


        throw new Error(
            getAPIError(
                result,
                "فشل الاتصال بـ Gemini."
            )
        );

    }


    if (
        !response.body
    ) {

        throw new Error(
            "المتصفح لا يدعم استقبال الرد المتدفق من Gemini."
        );

    }


    const reader =
        response.body.getReader();


    const decoder =
        new TextDecoder(
            "utf-8"
        );


    let buffer =
        "";


    let fullAnswer =
        "";


    function processSSELine(
        line
    ) {

        const cleanLine =
            String(
                line ||
                ""
            ).trim();


        if (
            !cleanLine ||
            !cleanLine.startsWith(
                "data:"
            )
        ) {

            return;

        }


        const dataText =
            cleanLine
                .substring(
                    5
                )
                .trim();


        if (
            !dataText ||
            dataText ===
                "[DONE]"
        ) {

            return;

        }


        let parsed;


        try {

            parsed =
                JSON.parse(
                    dataText
                );

        }
        catch (
            error
        ) {

            return;

        }


        if (
            !parsed ||
            !Array.isArray(
                parsed.candidates
            ) ||
            !parsed.candidates[0]
        ) {

            return;

        }


        const candidate =
            parsed.candidates[0];


        if (
            !candidate.content ||
            !Array.isArray(
                candidate.content.parts
            )
        ) {

            return;

        }


        candidate.content.parts.forEach(
            function (
                part
            ) {

                if (
                    !part ||
                    typeof part.text !==
                        "string"
                ) {

                    return;

                }


                if (
                    part.thought ===
                    true
                ) {

                    return;

                }


                const delta =
                    part.text;


                if (
                    !delta
                ) {

                    return;

                }


                fullAnswer +=
                    delta;


                if (
                    typeof onChunk ===
                    "function"
                ) {

                    onChunk(
                        delta,
                        fullAnswer
                    );

                }

            }
        );

    }


    while (
        true
    ) {

        const streamResult =
            await reader.read();


        if (
            streamResult.done
        ) {

            break;

        }


        buffer +=
            decoder.decode(
                streamResult.value,
                {
                    stream:
                        true
                }
            );


        buffer =
            buffer.replace(
                /\r\n/g,
                "\n"
            );


        buffer =
            buffer.replace(
                /\r/g,
                "\n"
            );


        let newlineIndex =
            buffer.indexOf(
                "\n"
            );


        while (
            newlineIndex !==
            -1
        ) {

            const line =
                buffer.substring(
                    0,
                    newlineIndex
                );


            buffer =
                buffer.substring(
                    newlineIndex + 1
                );


            processSSELine(
                line
            );


            newlineIndex =
                buffer.indexOf(
                    "\n"
                );

        }

    }


    if (
        buffer.trim()
    ) {

        processSSELine(
            buffer
        );

    }


    if (
        !fullAnswer.trim()
    ) {

        throw new Error(
            "لم يصل نص من Gemini عبر البث المتدفق."
        );

    }


    return fullAnswer.trim();

}


// =====================================================
// STREAM OPENROUTER AI
// =====================================================

async function streamOpenRouterAI(
    text,
    onChunk
) {

    const data =
        getSavedSettings();


    const key =
        data.key ||
        "";


    const model =
        data.model ||
        "";


    if (
        !key.trim()
    ) {

        throw new Error(
            "لم يتم إدخال مفتاح OpenRouter من الإعدادات."
        );

    }


    if (
        !model.trim()
    ) {

        throw new Error(
            "لم يتم تحديد نموذج OpenRouter."
        );

    }


    const documentContext =
        await buildAIDocumentContext(
            text
        );


    const conversationMessages =
        [];


    const historyLimit =
        documentContext &&
        documentContext.found
            ? 2
            : 4;


    if (
        currentChat &&
        Array.isArray(
            currentChat.messages
        )
    ) {

        const messagesWithoutCurrent =
            currentChat.messages.slice(
                0,
                -1
            );


        const previousMessages =
            messagesWithoutCurrent.slice(
                -historyLimit
            );


        previousMessages.forEach(
            function (
                msg
            ) {

                if (
                    !msg ||
                    !msg.text
                ) {

                    return;

                }


                let messageText =
                    String(
                        msg.text
                    ).trim();


                const maxHistoryChars =
                    documentContext &&
                    documentContext.found
                        ? 1000
                        : 1500;


                if (
                    messageText.length >
                    maxHistoryChars
                ) {

                    messageText =
                        messageText.substring(
                            0,
                            maxHistoryChars
                        ) +
                        "…";

                }


                conversationMessages.push({

                    role:
                        msg.role ===
                        "ai"
                            ? "assistant"
                            : "user",

                    content:
                        messageText

                });

            }
        );

    }


    let userContent =
        text;


    if (
        documentContext &&
        documentContext.found
    ) {

        userContent =
            [

                "أنت تجيب عن سؤال مستخدم في أداة بحث أكاديمية.",

                "",

                "=== سؤال المستخدم ===",

                text,

                "",

                "=== اسم المستند ===",

                (
                    currentDocument
                        ? currentDocument.name
                        : ""
                ),

                "",

                "=== المادة المستخرجة من المستند ===",

                documentContext.text

            ].join(
                "\n"
            );

    }


    conversationMessages.push({

        role:
            "user",

        content:
            userContent

    });


    const response =
        await fetch(
            "https://openrouter.ai/api/v1/chat/completions",
            {

                method:
                    "POST",

                headers: {

                    "Content-Type":
                        "application/json",

                    "Authorization":
                        "Bearer " +
                        key,

                    "HTTP-Referer":
                        window.location.href,

                    "X-Title":
                        "Research Tools"

                },

                body:
                    JSON.stringify({

                        model:
                            model,

                        messages:
                            conversationMessages,

                        max_tokens:
                            2500,

                        temperature:
                            0.2,

                        stream:
                            true

                    })

            }
        );


    if (
        !response.ok
    ) {

        const result =
            await readJSON(
                response
            );


        throw new Error(
            getAPIError(
                result,
                "فشل الاتصال بـ OpenRouter."
            )
        );

    }


    if (
        !response.body
    ) {

        throw new Error(
            "المتصفح لا يدعم استقبال الرد المتدفق من OpenRouter."
        );

    }


    const reader =
        response.body.getReader();


    const decoder =
        new TextDecoder(
            "utf-8"
        );


    let buffer =
        "";


    let fullAnswer =
        "";


    function processSSELine(
        line
    ) {

        const cleanLine =
            String(
                line ||
                ""
            ).trim();


        if (
            !cleanLine ||
            !cleanLine.startsWith(
                "data:"
            )
        ) {

            return;

        }


        const dataText =
            cleanLine
                .substring(
                    5
                )
                .trim();


        if (
            !dataText ||
            dataText ===
                "[DONE]"
        ) {

            return;

        }


        let parsed;


        try {

            parsed =
                JSON.parse(
                    dataText
                );

        }
        catch (
            error
        ) {

            return;

        }


        const delta =
            parsed &&
            parsed.choices &&
            parsed.choices[0] &&
            parsed.choices[0].delta
                ? parsed.choices[0].delta.content
                : "";


        if (
            typeof delta !==
                "string" ||
            !delta
        ) {

            return;

        }


        fullAnswer +=
            delta;


        if (
            typeof onChunk ===
            "function"
        ) {

            onChunk(
                delta,
                fullAnswer
            );

        }

    }


    while (
        true
    ) {

        const streamResult =
            await reader.read();


        if (
            streamResult.done
        ) {

            break;

        }


        buffer +=
            decoder.decode(
                streamResult.value,
                {
                    stream:
                        true
                }
            );


        buffer =
            buffer.replace(
                /\r\n/g,
                "\n"
            );


        buffer =
            buffer.replace(
                /\r/g,
                "\n"
            );


        let newlineIndex =
            buffer.indexOf(
                "\n"
            );


        while (
            newlineIndex !==
            -1
        ) {

            const line =
                buffer.substring(
                    0,
                    newlineIndex
                );


            buffer =
                buffer.substring(
                    newlineIndex + 1
                );


            processSSELine(
                line
            );


            newlineIndex =
                buffer.indexOf(
                    "\n"
                );

        }

    }


    if (
        buffer.trim()
    ) {

        processSSELine(
            buffer
        );

    }


    if (
        !fullAnswer.trim()
    ) {

        throw new Error(
            "لم يصل نص من OpenRouter عبر البث المتدفق."
        );

    }


    return fullAnswer.trim();

}


// =====================================================
// STREAM OPENAI
// =====================================================

async function streamOpenAI(
    text,
    onChunk
) {

    const data =
        getSavedSettings();


    const key =
        data.key ||
        "";


    const model =
        data.model ||
        "";


    if (
        !key.trim()
    ) {

        throw new Error(
            "لم يتم إدخال مفتاح OpenAI من الإعدادات."
        );

    }


    if (
        !model.trim()
    ) {

        throw new Error(
            "لم يتم تحديد نموذج OpenAI."
        );

    }


    const documentContext =
        await buildAIDocumentContext(
            text
        );


    const conversationMessages =
        [];


    const historyLimit =
        documentContext &&
        documentContext.found
            ? 2
            : 4;


    if (
        currentChat &&
        Array.isArray(
            currentChat.messages
        )
    ) {

        const messagesWithoutCurrent =
            currentChat.messages.slice(
                0,
                -1
            );


        const previousMessages =
            messagesWithoutCurrent.slice(
                -historyLimit
            );


        previousMessages.forEach(
            function (
                msg
            ) {

                if (
                    !msg ||
                    !msg.text
                ) {

                    return;

                }


                let messageText =
                    String(
                        msg.text
                    ).trim();


                const maxHistoryChars =
                    documentContext &&
                    documentContext.found
                        ? 1000
                        : 1500;


                if (
                    messageText.length >
                    maxHistoryChars
                ) {

                    messageText =
                        messageText.substring(
                            0,
                            maxHistoryChars
                        ) +
                        "…";

                }


                conversationMessages.push({

                    role:
                        msg.role ===
                        "ai"
                            ? "assistant"
                            : "user",

                    content:
                        messageText

                });

            }
        );

    }


    let userContent =
        text;


    if (
        documentContext &&
        documentContext.found
    ) {

        userContent =
            [

                "أنت تجيب عن سؤال مستخدم في أداة بحث أكاديمية.",

                "",

                "=== سؤال المستخدم ===",

                text,

                "",

                "=== اسم المستند ===",

                (
                    currentDocument
                        ? currentDocument.name
                        : ""
                ),

                "",

                "=== المادة المستخرجة من المستند ===",

                documentContext.text

            ].join(
                "\n"
            );

    }


    conversationMessages.push({

        role:
            "user",

        content:
            userContent

    });


    const response =
        await fetch(
            "https://api.openai.com/v1/chat/completions",
            {

                method:
                    "POST",

                headers: {

                    "Content-Type":
                        "application/json",

                    "Authorization":
                        "Bearer " +
                        key

                },

                body:
                    JSON.stringify({

                        model:
                            model,

                        messages:
                            conversationMessages,

                        max_tokens:
                            2500,

                        temperature:
                            0.2,

                        stream:
                            true

                    })

            }
        );


    if (
        !response.ok
    ) {

        const result =
            await readJSON(
                response
            );


        throw new Error(
            getAPIError(
                result,
                "فشل الاتصال بـ OpenAI."
            )
        );

    }


    if (
        !response.body
    ) {

        throw new Error(
            "المتصفح لا يدعم استقبال الرد المتدفق من OpenAI."
        );

    }


    const reader =
        response.body.getReader();


    const decoder =
        new TextDecoder(
            "utf-8"
        );


    let buffer =
        "";


    let fullAnswer =
        "";


    function processSSELine(
        line
    ) {

        const cleanLine =
            String(
                line ||
                ""
            ).trim();


        if (
            !cleanLine ||
            !cleanLine.startsWith(
                "data:"
            )
        ) {

            return;

        }


        const dataText =
            cleanLine
                .substring(
                    5
                )
                .trim();


        if (
            !dataText ||
            dataText ===
                "[DONE]"
        ) {

            return;

        }


        let parsed;


        try {

            parsed =
                JSON.parse(
                    dataText
                );

        }
        catch (
            error
        ) {

            return;

        }


        const delta =
            parsed &&
            parsed.choices &&
            parsed.choices[0] &&
            parsed.choices[0].delta
                ? parsed.choices[0].delta.content
                : "";


        if (
            typeof delta !==
                "string" ||
            !delta
        ) {

            return;

        }


        fullAnswer +=
            delta;


        if (
            typeof onChunk ===
            "function"
        ) {

            onChunk(
                delta,
                fullAnswer
            );

        }

    }


    while (
        true
    ) {

        const streamResult =
            await reader.read();


        if (
            streamResult.done
        ) {

            break;

        }


        buffer +=
            decoder.decode(
                streamResult.value,
                {
                    stream:
                        true
                }
            );


        buffer =
            buffer.replace(
                /\r\n/g,
                "\n"
            );


        buffer =
            buffer.replace(
                /\r/g,
                "\n"
            );


        let newlineIndex =
            buffer.indexOf(
                "\n"
            );


        while (
            newlineIndex !==
            -1
        ) {

            const line =
                buffer.substring(
                    0,
                    newlineIndex
                );


            buffer =
                buffer.substring(
                    newlineIndex + 1
                );


            processSSELine(
                line
            );


            newlineIndex =
                buffer.indexOf(
                    "\n"
                );

        }

    }


    if (
        buffer.trim()
    ) {

        processSSELine(
            buffer
        );

    }


    if (
        !fullAnswer.trim()
    ) {

        throw new Error(
            "لم يصل نص من OpenAI عبر البث المتدفق."
        );

    }


    return fullAnswer.trim();

}


// =====================================================
// SEND MESSAGE
// =====================================================

async function sendMessage() {

    if (
        !input
    ) {

        return;

    }


    const text =
        input.value.trim();


    if (
        !text
    ) {

        return;

    }


    if (
        !currentChat
    ) {

        currentChat = {

            id:
                Date.now(),

            title:
                text.substring(
                    0,
                    30
                ),

            messages:
                [],

            isTemporary:
                true,

            projectId:
                currentProject
                    ? currentProject.id
                    : null

        };

    }


    if (
        currentChat.isTemporary
    ) {

        currentChat.isTemporary =
            false;


        currentChat.projectId =
            currentProject
                ? currentProject.id
                : null;


        currentChat.title =
            text.substring(
                0,
                30
            );


        const alreadyExists =
            chats.some(
                function (
                    chat
                ) {

                    return (
                        chat.id ===
                        currentChat.id
                    );

                }
            );


        if (
            !alreadyExists
        ) {

            chats.unshift(
                currentChat
            );

        }


        if (
            currentProject &&
            currentChat.projectId ===
                currentProject.id
        ) {

            if (
                !Array.isArray(
                    currentProject.chatIds
                )
            ) {

                currentProject.chatIds =
                    [];

            }


            if (
                !currentProject.chatIds.includes(
                    currentChat.id
                )
            ) {

                currentProject.chatIds.push(
                    currentChat.id
                );


                currentProject.updatedAt =
                    new Date().toISOString();


                saveProjects();

            }

        }


        saveChats();

    }


    currentChat.messages.push({

        role:
            "user",

        text:
            text

    });


    saveChats();


    renderChat();


    renderChatList();


    renderSidebarChats();


    renderRecentChats();


    input.value =
        "";


    input.style.height =
        "auto";


    const loading =
        document.createElement(
            "div"
        );


    loading.className =
        "message ai-message";


    loading.innerHTML =
        "⏳ جاري التفكير...";


    if (
        chatArea
    ) {

        chatArea.appendChild(
            loading
        );


        chatArea.scrollTop =
            chatArea.scrollHeight;

    }


    const savedSettings =
        getSavedSettings();


    const selectedProvider =
        String(
            savedSettings.provider ||
            "openrouter"
        ).toLowerCase();


    let pendingRenderText =
        "";


    let renderTimer =
        null;


    function renderStreamingText() {

        if (
            !loading
        ) {

            return;

        }


        if (
            !pendingRenderText
        ) {

            loading.innerHTML =
                "⏳ جاري التفكير...";

        }
        else {

            loading.innerHTML =
                formatAIMessage(
                    pendingRenderText,
                    currentCitationSources
                );

        }


        if (
            chatArea
        ) {

            chatArea.scrollTop =
                chatArea.scrollHeight;

        }


        renderTimer =
            null;

    }


    function scheduleRender(
        fullText
    ) {

        pendingRenderText =
            String(
                fullText ||
                ""
            );


        if (
            renderTimer !==
            null
        ) {

            return;

        }


        renderTimer =
            setTimeout(
                function () {

                    renderStreamingText();

                },
                60
            );

    }


    try {

        let answer =
            "";


        if (
            selectedProvider ===
            "groq"
        ) {

            answer =
                await streamGroqAI(
                    text,
                    function (
                        delta,
                        fullText
                    ) {

                        scheduleRender(
                            fullText
                        );

                    }
                );

        }
        else if (
            selectedProvider ===
            "gemini"
        ) {

            answer =
                await streamGeminiAI(
                    text,
                    function (
                        delta,
                        fullText
                    ) {

                        scheduleRender(
                            fullText
                        );

                    }
                );

        }
        else if (
            selectedProvider ===
            "openrouter"
        ) {

            answer =
                await streamOpenRouterAI(
                    text,
                    function (
                        delta,
                        fullText
                    ) {

                        scheduleRender(
                            fullText
                        );

                    }
                );

        }
        else if (
            selectedProvider ===
            "openai"
        ) {

            answer =
                await streamOpenAI(
                    text,
                    function (
                        delta,
                        fullText
                    ) {

                        scheduleRender(
                            fullText
                        );

                    }
                );

        }
        else {

            throw new Error(
                "مزود الذكاء الاصطناعي غير معروف: " +
                selectedProvider
            );

        }


        if (
            renderTimer !==
            null
        ) {

            clearTimeout(
                renderTimer
            );


            renderTimer =
                null;

        }


        pendingRenderText =
            String(
                answer ||
                ""
            );


        renderStreamingText();


        if (
            loading &&
            loading.parentNode
        ) {

            loading.remove();

        }


        currentChat.messages.push({

            role:
                "ai",

            text:
                String(
                    answer ||
                    ""
                ),

            citationSources:
                Array.isArray(
                    currentCitationSources
                )
                    ? currentCitationSources.map(
                        function (
                            source
                        ) {

                            return {

                                rank:
                                    source.rank,

                                paragraphIndex:
                                    source.paragraphIndex,

                                heading:
                                    source.heading ||
                                    "",

                                text:
                                    source.text ||
                                    ""

                            };

                        }
                    )
                    : []

        });


        saveChats();


        renderChat();


        renderChatList();


        renderSidebarChats();


        renderRecentChats();

    }
    catch (
        error
    ) {

        if (
            renderTimer !==
            null
        ) {

            clearTimeout(
                renderTimer
            );


            renderTimer =
                null;

        }


        if (
            loading &&
            loading.parentNode
        ) {

            loading.remove();

        }


        currentChat.messages.push({

            role:
                "ai",

            text:
                "خطأ: " +
                (
                    error &&
                    error.message
                        ? error.message
                        : "حدث خطأ غير معروف"
                )

        });


        saveChats();


        renderChat();


        renderChatList();


        renderSidebarChats();


        renderRecentChats();

    }

}


// =====================================================
// SEND BUTTON
// =====================================================

if (
    sendBtn
) {

    sendBtn.onclick =
        function (
            e
        ) {

            e.preventDefault();
            e.stopPropagation();


            sendMessage();

        };

}


// =====================================================
// KEYBOARD
// =====================================================

if (
    input
) {

    input.onkeydown =
        function (
            e
        ) {

            if (
                e.key ===
                    "Enter" &&
                !e.shiftKey
            ) {

                e.preventDefault();


                sendMessage();

            }

        };

}


// =====================================================
// FINAL INITIALIZATION
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


// =====================================================
// END OFFICE.ONREADY
// =====================================================

