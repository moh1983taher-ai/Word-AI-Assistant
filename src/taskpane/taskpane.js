// ======================================
// Word AI Assistant
// Main Application Controller
// ======================================

Office.onReady(function () {

// ======================================
// Elements
// ======================================

const projectsBtn =
    document.getElementById("projects-btn");

const projectsPopup =
    document.getElementById("projects-popup");

const projectsList =
    document.getElementById("projects-list");

const documentsList =
    document.getElementById("documents-list");

const addDocumentBtn =
    document.getElementById("add-document-btn");

const wordDocumentPicker =
    document.getElementById(
        "word-document-picker"
    );

// ======================================
// Render Project Documents
// ======================================

function renderDocuments() {

    if (!documentsList)
        return;


    documentsList.innerHTML = "";


    // لا يوجد مشروع محدد
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


    // لا توجد مستندات
    if (projectDocuments.length === 0) {

        documentsList.innerHTML = `
            <div class="empty-document">
                لا توجد مستندات
            </div>
        `;

        return;
    }


    // ==================================
    // عرض المستندات
    // ==================================

    projectDocuments.forEach(
        function (documentItem, index) {

            const item =
                document.createElement("div");


            item.className =
                "document-item";

            if (
                currentDocument &&
                currentDocument.id === documentItem.id
            ) {

                item.classList.add(
                    "active-document"
                );

            }
            // --------------------------------
            // عنوان المستند
            // --------------------------------

            const title =
                document.createElement("span");

            title.className =
                "document-title";

            title.textContent =
                documentItem.name;
            const status =
                document.createElement("span");

            status.className =
                "document-read-status";


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

            // --------------------------------
            // زر الخيارات
            // --------------------------------

            const menuButton =
                document.createElement("button");

            menuButton.className =
                "document-menu";

            menuButton.type =
                "button";

            menuButton.title =
                "خيارات المستند";

            menuButton.textContent =
                "⋮";


            // --------------------------------
            // قائمة الخيارات
            // --------------------------------

            const options =
                document.createElement("div");

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


            // ==================================
            // ضبط ترتيب الخيارات
            // ==================================

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


            // ==================================
            // فتح / إغلاق القائمة
            // ==================================

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
            // إعادة التسمية
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
                                    new Date().toISOString();

                                saveDocuments();

                            }
                            else {

                                documentItem.name =
                                    oldName;

                            }


                            renderDocuments();

                        }


                        inputRename.onkeydown =
                            function (event) {

                                if (
                                    event.key ===
                                    "Enter"
                                ) {

                                    event.preventDefault();

                                    finishRename(true);

                                }


                                if (
                                    event.key ===
                                    "Escape"
                                ) {

                                    event.preventDefault();

                                    finishRename(false);

                                }

                            };

                    };

            }


            // ==================================
            // نقل إلى أعلى
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
                            new Date().toISOString();

                        previousDocument.updatedAt =
                            new Date().toISOString();


                        saveDocuments();


                        options.classList.remove(
                            "open"
                        );


                        renderDocuments();

                    };

            }


            // ==================================
            // نقل إلى أسفل
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
                            new Date().toISOString();

                        nextDocument.updatedAt =
                            new Date().toISOString();


                        saveDocuments();


                        options.classList.remove(
                            "open"
                        );


                        renderDocuments();

                    };

            }


// ==================================
// حذف المستند
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


            // منع وجود أكثر من نافذة تأكيد
            const oldConfirm =
                document.querySelector(
                    ".document-delete-confirm"
                );

            if (oldConfirm) {
                oldConfirm.remove();
            }


            // ==================================
            // نافذة التأكيد
            // ==================================

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


            // ==================================
            // تأكيد الحذف
            // ==================================

            const confirmDelete =
                confirmBox.querySelector(
                    ".confirm-document-delete"
                );


            if (confirmDelete) {

                confirmDelete.onclick =
                    function () {

                        // حذف المستند من التخزين
                        documents =
                            documents.filter(
                                function (doc) {

                                    return (
                                        doc.id !==
                                        documentItem.id
                                    );

                                }
                            );


                        // حذف معرف المستند من المشروع
                        if (
                            currentProject &&
                            Array.isArray(
                                currentProject.documents
                            )
                        ) {

                            currentProject.documents =
                                currentProject.documents.filter(
                                    function (id) {

                                        return (
                                            id !==
                                            documentItem.id
                                        );

                                    }
                                );


                            currentProject.updatedAt =
                                new Date().toISOString();


                            saveProjects();

                        }


                        // إعادة ترتيب المستندات المتبقية
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
                                    new Date().toISOString();

                            }
                        );


                        // حفظ المستندات
                        saveDocuments();


                        // إغلاق نافذة التأكيد
                        confirmBox.remove();


                        // إعادة العرض
                        renderDocuments();

                    };

            }


            // ==================================
            // إلغاء
            // ==================================

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


            // ==================================
            // النقر خارج نافذة الحوار = إلغاء
            // ==================================

            confirmBox.onclick =
                function (event) {

                    if (
                        event.target ===
                        confirmBox
                    ) {

                        confirmBox.remove();

                    }

                };

        };

}


            // ==================================
            // إضافة العناصر
            // ==================================

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

const newProjectBtn =
document.getElementById("new-project-btn");

const chatBtn =
document.getElementById("chat-btn");

const newChatBtn =
document.getElementById("new-chat-btn");

const expandedSidebar =
document.getElementById("expanded-sidebar");

const sidebarToggleBtn =
document.getElementById("sidebar-toggle-btn");

const expandedSidebarToggleSlot =
document.getElementById(
    "expanded-sidebar-toggle-slot"
);

const sidebarTogglePlaceholder =
    document.createComment(
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

const input =
document.getElementById("user-input");

const sendBtn =
document.getElementById("send-btn");

const chatArea =
document.getElementById("chat-area");

const documentTitle =
document.getElementById("document-title");

// ======================================
// AI Settings
// ======================================

const settingsBtn =
document.getElementById("settings-btn");

const settingsWindow =
document.getElementById("settings-window");

const closeSettings =
document.getElementById("close-settings");

const showKey =
document.getElementById("show-key");

const apiKey =
document.getElementById("api-key");

const provider =
document.getElementById("provider-select");

const modelSelect =
document.getElementById("model-select");

const refreshModels =
document.getElementById("refresh-models");

const saveSettings =
document.getElementById("save-settings");

const testConnection =
document.getElementById("test-connection");

const settingsStatus =
document.getElementById("settings-status");

const providerInfo =
document.getElementById("provider-info");

// ======================================
// Chat popup
// ======================================

const chatPopup =
document.getElementById("chat-popup");

const recentChatList =
document.getElementById("recent-chat-list");

// ======================================
// Search
// ======================================

const searchPopup =
document.getElementById("search-popup");

const searchInput =
document.getElementById("search-input");

const searchResults =
document.getElementById("search-results");

const searchBtn =
document.getElementById("search-btn");



// ======================================
// Projects System
// ======================================

let projects = [];

try {

    projects =
        JSON.parse(
            localStorage.getItem(
                "WORD_AI_PROJECTS"
            )
        ) || [];

}
catch (e) {

    projects = [];

}


/* ======================================
   ضمان البنية الجديدة للمشاريع
   ====================================== */

projects =
    projects
        .filter(function (project) {

            return (
                project &&
                typeof project === "object"
            );

        })
        .map(function (project) {

            const now =
                new Date().toISOString();

            return {

                id:
                    project.id ||
                    Date.now(),

                name:
                    project.name ||
                    "مشروع جديد",

                createdAt:
                    project.createdAt ||
                    now,

                updatedAt:
                    project.updatedAt ||
                    now,

                documents:
                    Array.isArray(
                        project.documents
                    )
                        ? project.documents
                        : [],

                references:
                    Array.isArray(
                        project.references
                    )
                        ? project.references
                        : [],

                chatIds:
                    Array.isArray(
                        project.chatIds
                    )
                        ? project.chatIds
                        : [],

                settings:
                    project.settings &&
                    typeof project.settings === "object"
                        ? project.settings
                        : {
                            citationStyle: "",
                            notes: ""
                        }

            };

        });


let currentProject = null;
let currentDocument = null;


// ======================================
// Documents System
// ======================================

let documents = [];

try {

    documents =
        JSON.parse(
            localStorage.getItem(
                "WORD_AI_DOCUMENTS"
            )
        ) || [];


    documents =
        documents
            .filter(function (documentItem) {

                return (
                    documentItem &&
                    typeof documentItem ===
                        "object"
                );

            })
            .map(function (documentItem) {

                if (
                    typeof documentItem.indexTokenCount !==
                    "number"
                ) {

                    documentItem.indexTokenCount =
                        0;

                }

                if (
                    typeof documentItem.indexUniqueTerms !==
                    "number"
                ) {

                    documentItem.indexUniqueTerms =
                        0;

                }

                return documentItem;

            });

}
catch (e) {

    documents = [];

}


// ======================================
// Save Documents
// ======================================

function saveDocuments() {

    localStorage.setItem(
        "WORD_AI_DOCUMENTS",
        JSON.stringify(documents)
    );

}

saveDocuments();

// ======================================
// Working Documents Storage
// IndexedDB
// ======================================

const DOCUMENT_DB_NAME =
    "WORD_AI_DOCUMENT_STORAGE";

const DOCUMENT_DB_VERSION =
    4;

const DOCUMENT_STORE_NAME =
    "files";
const DOCUMENT_TEXT_STORE_NAME =
    "texts";
const DOCUMENT_INDEX_STORE_NAME =
    "indexes";
const DOCUMENT_STRUCTURE_STORE_NAME =
    "structures";

// ======================================
// Open Documents Database
// ======================================

function openDocumentDatabase() {

    return new Promise(
        function (resolve, reject) {

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
                                DOCUMENT_INDEX_STORE_NAME
                            )
                    ) {

                        db.createObjectStore(
                            DOCUMENT_INDEX_STORE_NAME
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
                        request.error
                    );

                };

        }
    );

}


// ======================================
// Save Working Word File
// ======================================

async function saveWorkingWordFile(
    fileId,
    file
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (resolve, reject) {

            const transaction =
                db.transaction(
                    DOCUMENT_STORE_NAME,
                    "readwrite"
                );


            const store =
                transaction.objectStore(
                    DOCUMENT_STORE_NAME
                );


            const request =
                store.put(
                    file,
                    fileId
                );


            request.onsuccess =
                function () {

                    resolve(
                        fileId
                    );

                };


            request.onerror =
                function () {

                    reject(
                        request.error
                    );

                };


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}


// ======================================
// Get Working Word File
// ======================================

async function getWorkingWordFile(
    fileId
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (resolve, reject) {

            const transaction =
                db.transaction(
                    DOCUMENT_STORE_NAME,
                    "readonly"
                );


            const store =
                transaction.objectStore(
                    DOCUMENT_STORE_NAME
                );


            const request =
                store.get(
                    fileId
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


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}

// ======================================
// Save Indexed Document Text
// ======================================

async function saveDocumentText(
    documentId,
    text
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (resolve, reject) {

            const transaction =
                db.transaction(
                    DOCUMENT_TEXT_STORE_NAME,
                    "readwrite"
                );


            const store =
                transaction.objectStore(
                    DOCUMENT_TEXT_STORE_NAME
                );


            const record = {

                documentId:
                    String(documentId),

                text:
                    String(text || ""),

                updatedAt:
                    new Date().toISOString()

            };


            const request =
                store.put(
                    record,
                    String(documentId)
                );


            request.onsuccess =
                function () {

                    resolve(record);

                };


            request.onerror =
                function () {

                    reject(
                        request.error
                    );

                };


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}

// ======================================
// Get Indexed Document Text
// ======================================

async function getDocumentText(
    documentId
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (resolve, reject) {

            const transaction =
                db.transaction(
                    DOCUMENT_TEXT_STORE_NAME,
                    "readonly"
                );


            const store =
                transaction.objectStore(
                    DOCUMENT_TEXT_STORE_NAME
                );


            const request =
                store.get(
                    String(documentId)
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


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}

// ======================================
// Normalize Search Text
// ======================================

function normalizeSearchText(text) {

    return String(
        text || ""
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

        // توحيد الهمزات
        .replace(
            /[أإآ]/g,
            "ا"
        )

        // توحيد الياء والألف المقصورة
        .replace(
            /ى/g,
            "ي"
        )

        // إزالة المسافات المتكررة
        .replace(
            /\s+/g,
            " "
        )

        .trim()

        .toLowerCase();

}

// ======================================
// Tokenize Document Text
// ======================================

function tokenizeDocumentText(text) {

    const normalized =
        normalizeSearchText(
            text
        );


    const matches =
        normalized.match(
            /[A-Za-z0-9\u0600-\u06FF\u0750-\u077F]+/g
        );


    return matches || [];

}
// ======================================
// Build Document Index
// ======================================

function buildDocumentIndex(
    documentId,
    text
) {

    const tokens =
        tokenizeDocumentText(
            text
        );


    const terms = {};


    tokens.forEach(
        function (token, index) {

            if (!terms[token]) {

                terms[token] = {

                    count:
                        0,

                    positions:
                        []

                };

            }


            terms[token].count += 1;


            terms[token].positions.push(
                index
            );

        }
    );


    return {

        documentId:
            String(documentId),

        tokenCount:
            tokens.length,

        uniqueTerms:
            Object.keys(
                terms
            ).length,

        terms:
            terms,

        updatedAt:
            new Date().toISOString()

    };

}
// ======================================
// Build Document Structure
// ======================================

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
        async function (context) {

            if (
                !Office.context.requirements.isSetSupported(
                    "WordApiHiddenDocument",
                    "1.3"
                )
            ) {

                throw new Error(
                    "إصدار Word الحالي لا يدعم تحليل بنية المستند."
                );

            }


            const workingDocument =
                context.application.createDocument(
                    base64
                );


            const paragraphs =
                workingDocument.body.paragraphs;


            paragraphs.load(
                [
                    "items/text",
                    "items/styleBuiltIn",
                    "items/tableNestingLevel"
                ]
            );


            const tables =
                workingDocument.body.tables;


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
                    function (paragraph, index) {

                        const text =
                            String(
                                paragraph.text ||
                                ""
                            ).trim();


                        return {

                            index:
                                index,

                            id:
                                String(index),

                            text:
                                text,

                            style:
                                paragraph.styleBuiltIn ||
                                "",

                            tableNestingLevel:
                                paragraph.tableNestingLevel ||
                                0

                        };

                    }
                );


            const headings =
                paragraphItems.filter(
                    function (paragraph) {

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
                    function (table, index) {

                        return {

                            index:
                                index,

                            rows:
                                table.rowCount,

                            columns:
                                table.columnCount,

                            style:
                                table.styleBuiltIn ||
                                ""

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
                    new Date().toISOString()

            };

        }
    );

}
// ======================================
// Save Document Structure
// ======================================

async function saveDocumentStructure(
    documentId,
    structureData
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (resolve, reject) {

            const transaction =
                db.transaction(
                    DOCUMENT_STRUCTURE_STORE_NAME,
                    "readwrite"
                );


            const store =
                transaction.objectStore(
                    DOCUMENT_STRUCTURE_STORE_NAME
                );


            const request =
                store.put(
                    structureData,
                    String(documentId)
                );


            request.onsuccess =
                function () {

                    resolve(
                        structureData
                    );

                };


            request.onerror =
                function () {

                    reject(
                        request.error
                    );

                };


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}
// ======================================
// Get Document Structure
// ======================================

async function getDocumentStructure(
    documentId
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (resolve, reject) {

            const transaction =
                db.transaction(
                    DOCUMENT_STRUCTURE_STORE_NAME,
                    "readonly"
                );


            const store =
                transaction.objectStore(
                    DOCUMENT_STRUCTURE_STORE_NAME
                );


            const request =
                store.get(
                    String(documentId)
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


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}
// ======================================
// Search Document Context
// ======================================

async function searchDocumentContext(
    documentId,
    query
) {

    const searchTerm =
        normalizeSearchText(
            query
        );


    if (!searchTerm) {

        return {

            query:
                "",

            count:
                0,

            results:
                []

        };

    }


    // ==================================
    // استرجاع النص
    // ==================================

    const textData =
        await getDocumentText(
            documentId
        );


    if (!textData) {

        throw new Error(
            "لا يوجد نص محفوظ لهذا المستند."
        );

    }


    // ==================================
    // استرجاع البنية
    // ==================================

    const structureData =
        await getDocumentStructure(
            documentId
        );


    if (!structureData) {

        throw new Error(
            "لا توجد بنية محفوظة لهذا المستند."
        );

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


    const results = [];


    // ==================================
    // البحث فقرة فقرة
    // ==================================

    paragraphs.forEach(
        function (paragraph) {

            if (
                !paragraph ||
                !paragraph.text
            ) {

                return;

            }


            const originalText =
                String(
                    paragraph.text
                );


            const normalizedParagraph =
                normalizeSearchText(
                    originalText
                );


            if (
                !normalizedParagraph.includes(
                    searchTerm
                )
            ) {

                return;

            }


            // ==================================
            // العثور على العنوان الأقرب
            // ==================================

            let nearestHeading =
                null;


            headings.forEach(
                function (heading) {

                    if (
                        heading.index <
                        paragraph.index
                    ) {

                        if (
                            !nearestHeading ||
                            heading.index >
                            nearestHeading.index
                        ) {

                            nearestHeading =
                                heading;

                        }

                    }

                }
            );


            // ==================================
            // استخراج المقتطف
            // ==================================

            const position =
                normalizedParagraph.indexOf(
                    searchTerm
                );


            const start =
                Math.max(
                    0,
                    position - 100
                );


            const end =
                Math.min(
                    normalizedParagraph.length,
                    position +
                    searchTerm.length +
                    160
                );


            const context =
                normalizedParagraph.substring(
                    start,
                    end
                );


            results.push({

                paragraphIndex:
                    paragraph.index,

                paragraphId:
                    paragraph.id,

                text:
                    originalText,

                context:
                    context,

                heading:
                    nearestHeading
                        ? nearestHeading.text
                        : "",

                headingLevel:
                    nearestHeading
                        ? nearestHeading.style
                        : ""

            });

        }
    );


    return {

        query:
            searchTerm,

        count:
            results.length,

        results:
            results

    };

}
// ======================================
// Save Document Index
// ======================================

async function saveDocumentIndex(
    documentId,
    indexData
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (resolve, reject) {

            const transaction =
                db.transaction(
                    DOCUMENT_INDEX_STORE_NAME,
                    "readwrite"
                );


            const store =
                transaction.objectStore(
                    DOCUMENT_INDEX_STORE_NAME
                );


            const request =
                store.put(
                    indexData,
                    String(documentId)
                );


            request.onsuccess =
                function () {

                    resolve(
                        indexData
                    );

                };


            request.onerror =
                function () {

                    reject(
                        request.error
                    );

                };


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}
// ======================================
// Get Document Index
// ======================================

async function getDocumentIndex(
    documentId
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (resolve, reject) {

            const transaction =
                db.transaction(
                    DOCUMENT_INDEX_STORE_NAME,
                    "readonly"
                );


            const store =
                transaction.objectStore(
                    DOCUMENT_INDEX_STORE_NAME
                );


            const request =
                store.get(
                    String(documentId)
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


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}
// ======================================
// Test Document Index
// ======================================

async function testCurrentDocumentIndex() {

    if (!currentDocument) {

        console.warn(
            "لا يوجد مستند نشط."
        );

        return;

    }

    try {

        const index =
            await getDocumentIndex(
                currentDocument.id
            );

        console.log(
            "الفهرس الحقيقي للمستند:",
            index
        );

        if (!index) {

            console.warn(
                "لا يوجد فهرس محفوظ لهذا المستند."
            );

            return;

        }

        console.log(
            "عدد الكلمات:",
            index.tokenCount
        );

        console.log(
            "عدد الكلمات الفريدة:",
            index.uniqueTerms
        );

        console.log(
            "الكلمات المفهرسة:",
            Object.keys(
                index.terms || {}
            ).slice(
                0,
                20
            )
        );

    }
    catch (error) {

        console.error(
            "فشل اختبار الفهرس:",
            error
        );

    }

}
// ======================================
// Search Indexed Document
// ======================================

async function searchIndexedDocument(
    documentId,
    query
) {

    const searchTerm =
        normalizeSearchText(
            query
        );


    if (!searchTerm) {

        return {

            query:
                "",

            count:
                0,

            results:
                []

        };

    }


    const indexData =
        await getDocumentIndex(
            documentId
        );


    if (!indexData) {

        throw new Error(
            "لا يوجد فهرس لهذا المستند."
        );

    }


    const textData =
        await getDocumentText(
            documentId
        );


    if (!textData) {

        throw new Error(
            "لا يوجد نص مفهرس لهذا المستند."
        );

    }


    const originalText =
        String(
            textData.text || ""
        );


    const normalizedText =
        normalizeSearchText(
            originalText
        );


    const results = [];


    let position =
        normalizedText.indexOf(
            searchTerm
        );


    while (
        position !== -1
    ) {

        const start =
            Math.max(
                0,
                position - 80
            );


        const end =
            Math.min(
                normalizedText.length,
                position +
                searchTerm.length +
                120
            );


        results.push({

            position:
                position,

            context:
                normalizedText.substring(
                    start,
                    end
                )

        });


        position =
            normalizedText.indexOf(
                searchTerm,
                position +
                searchTerm.length
            );

    }


    return {

        query:
            searchTerm,

        count:
            results.length,

        results:
            results,

        indexTokenCount:
            indexData.tokenCount,

        indexUniqueTerms:
            indexData.uniqueTerms

    };

}
// ======================================
// Delete Working Word File
// ======================================

async function deleteWorkingWordFile(
    fileId
) {

    const db =
        await openDocumentDatabase();


    return new Promise(
        function (resolve, reject) {

            const transaction =
                db.transaction(
                    DOCUMENT_STORE_NAME,
                    "readwrite"
                );


            const store =
                transaction.objectStore(
                    DOCUMENT_STORE_NAME
                );


            const request =
                store.delete(
                    fileId
                );


            request.onsuccess =
                function () {

                    resolve();

                };


            request.onerror =
                function () {

                    reject(
                        request.error
                    );

                };


            transaction.oncomplete =
                function () {

                    db.close();

                };

        }
    );

}
// ======================================
// Convert Blob/File to Base64
// ======================================

function fileToBase64(file) {

    return new Promise(
        function (resolve, reject) {

            const reader =
                new FileReader();

            reader.onload =
                function () {

                    const result =
                        String(
                            reader.result || ""
                        );

                    const commaIndex =
                        result.indexOf(",");

                    if (
                        commaIndex === -1
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
                            commaIndex + 1
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

            reader.readAsDataURL(file);

        }
    );

}

// ======================================
// Create Document
// ======================================

function createDocument(
    file,
    projectId,
    order
) {

    const now =
        new Date().toISOString();


    const documentId =
        Date.now();


    const documentItem = {

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
            String(documentId),

        order:
            typeof order === "number"
                ? order
                : 0,

        type:
            "word",
        
        readStatus:
            "new",
        
        indexStatus:
        "new",
        
        createdAt:
            now,

        updatedAt:
            now

    };


    documents.push(
        documentItem
    );


    saveDocuments();


    return documentItem;

}

// ======================================
// Update Document Read Status
// ======================================

function updateDocumentReadStatus(
    documentItem,
    status
) {

    if (!documentItem)
        return;


    documentItem.readStatus =
        status;


    documentItem.updatedAt =
        new Date().toISOString();


    if (
        status ===
        "read"
    ) {

        documentItem.readAt =
            new Date().toISOString();

    }


    saveDocuments();

}
// ======================================
// Update Document Index Status
// ======================================

function updateDocumentIndexStatus(
    documentItem,
    status
) {

    if (!documentItem)
        return;


    documentItem.indexStatus =
        status;


    documentItem.updatedAt =
        new Date().toISOString();


    saveDocuments();

}
// ======================================
// Get Project Documents
// ======================================

function getProjectDocuments(
    projectId
) {

    if (!projectId)
        return [];

    return documents
        .filter(function (documentItem) {

            return (
                documentItem &&
                documentItem.projectId ===
                    projectId
            );

        })
        .sort(function (a, b) {

            return (
                a.order -
                b.order
            );

        });

}

// ======================================
// Set Active Document
// ======================================

function setCurrentDocument(documentItem) {

    if (!documentItem) {

        currentDocument = null;

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


    // ==================================
    // القراءة تكون مرة واحدة فقط
    // ==================================

    if (
        documentItem.readStatus ===
        "read"
    ) {

        // المستند مقروء مسبقًا
        renderDocuments();

        return;

    }


    // ==================================
    // المستند يحتاج إلى قراءة
    // ==================================

    readCurrentWordDocument(
        documentItem
    )
    .then(function (text) {

        console.log(
            "محتوى نسخة العمل:",
            text
        );

        renderDocuments();

    })
    .catch(function (error) {

        console.error(
            "تعذر قراءة نسخة العمل:",
            error
        );

        renderDocuments();

    });


    // إظهار "جارٍ القراءة..." فورًا
    renderDocuments();

}
// ======================================
// Read Current Working Document
// ======================================

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


    try {

        if (
            !Office.context.requirements.isSetSupported(
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
                async function (context) {

                    const workingDocument =
                        context.application.createDocument(
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


        // ======================================
        // تمت القراءة
        // ======================================

        updateDocumentReadStatus(
            documentItem,
            "read"
        );


        // ======================================
        // بدء الفهرسة
        // ======================================

        updateDocumentIndexStatus(
            documentItem,
            "indexing"
        );


        try {

            const indexData =
                buildDocumentIndex(
                    documentItem.id,
                    text
                );

            await saveDocumentIndex(
                documentItem.id,
                indexData
            );
            
            const structureData =
                await buildDocumentStructure(
                    documentItem
                );


            await saveDocumentStructure(
                documentItem.id,
                structureData
            );

            // ==================================
            // حفظ إحصاءات الفهرس مع المستند
            // ==================================

            documentItem.indexTokenCount =
                indexData.tokenCount;

            documentItem.indexUniqueTerms =
                indexData.uniqueTerms;

            documentItem.indexUpdatedAt =
                indexData.updatedAt;


            // ==================================
            // تثبيت حالة الفهرسة
            // ==================================

            updateDocumentIndexStatus(
                documentItem,
                "indexed"
            );


        }
        catch (indexError) {

            updateDocumentIndexStatus(
                documentItem,
                "error"
            );


            throw indexError;

        }


        return text;

    }

    catch (error) {

        updateDocumentReadStatus(
            documentItem,
            "error"
        );


        throw error;

    }

}
// ======================================
// Update Document Timestamp
// ======================================

function touchDocument(
    documentItem
) {

    if (!documentItem)
        return;

    documentItem.updatedAt =
        new Date().toISOString();

}


// ======================================
// Attach Document To Project
// ======================================

function attachDocumentToProject(
    project,
    documentItem
) {

    if (
        !project ||
        !documentItem
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
            documentItem.id
        )
    ) {

        project.documents.push(
            documentItem.id
        );

        project.updatedAt =
            new Date().toISOString();

        saveProjects();

    }

}
// ======================================
// Set Active Project
// ======================================

function setCurrentProject(project) {

    if (!project) {

        currentProject = null;

        renderDocuments();

        return;

    }


    currentProject =
        project;


    renderDocuments();

}


// ======================================
// Add Document Button
// Choose Real DOCX File
// ======================================

if (
    addDocumentBtn &&
    wordDocumentPicker
) {

    addDocumentBtn.onclick =
        function (e) {

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


            // فتح نافذة اختيار الملف
            wordDocumentPicker.value =
                "";


            wordDocumentPicker.click();

        };


    // ==================================
    // عند اختيار الملف
    // ==================================

    wordDocumentPicker.onchange =
        async function () {

            try {

                const file =
                    wordDocumentPicker.files &&
                    wordDocumentPicker.files[0];


                if (!file)
                    return;


                // نقبل DOCX فقط
                const isDocx =
                    /\.docx$/i.test(
                        file.name
                    );


                if (!isDocx) {

                    console.warn(
                        "الملف المختار ليس DOCX."
                    );

                    return;

                }


                if (!currentProject) {

                    return;

                }


                // ==========================
                // تحديد الترتيب
                // ==========================

                const projectDocuments =
                    getProjectDocuments(
                        currentProject.id
                    );


                const nextOrder =
                    projectDocuments.length + 1;


                // ==========================
                // إنشاء سجل المستند
                // ==========================

                const documentItem =
                    createDocument(
                        file,
                        currentProject.id,
                        nextOrder
                    );


                // ==========================
                // حفظ نسخة الملف
                // ==========================

                await saveWorkingWordFile(
                    documentItem.storageId,
                    file
                );


                // ==========================
                // ربط المستند بالمشروع
                // ==========================

                attachDocumentToProject(
                    currentProject,
                    documentItem
                );


                // ==========================
                // جعله المستند النشط
                // ==========================

                setCurrentDocument(
                    documentItem
                );


                // ==========================
                // إعادة العرض
                // ==========================

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

            catch (error) {

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
// ======================================
// Project Icon
// ======================================

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
// ======================================
// ايقونة محادثات
// ======================================
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

function saveProjects() {

    localStorage.setItem(
        "WORD_AI_PROJECTS",
        JSON.stringify(projects)
    );

}


// ======================================
// Chat System
// ======================================

let chats = [];

try {


chats =
    JSON.parse(
        localStorage.getItem(
            "WORD_AI_CHATS"
        )
    ) || [];


}
catch (e) {


chats = [];


}

let currentChat = null;

function saveChats() {


localStorage.setItem(
    "WORD_AI_CHATS",
    JSON.stringify(chats)
);


}

// ======================================
// Sidebar
// Expand / Collapse Sections
// ======================================

function initializeSidebarSections() {


const headers =
    document.querySelectorAll(
        ".section-title[data-target], .section-toggle[data-target]"
    );


headers.forEach(function (header) {

    const targetId =
        header.getAttribute("data-target");

    if (!targetId)
        return;


    const target =
        document.getElementById(targetId);

    if (!target)
        return;


    target.classList.remove("open");
    header.classList.remove("open");


    header.onclick =
        function (e) {

            e.preventDefault();
            e.stopPropagation();


            const isOpen =
                target.classList.contains(
                    "open"
                );


            if (isOpen) {

                target.classList.remove(
                    "open"
                );

                header.classList.remove(
                    "open"
                );

            }
            else {

                target.classList.add(
                    "open"
                );

                header.classList.add(
                    "open"
                );

            }

        };

});


}

// ======================================
// Close All Sidebar Sections
// ======================================

function closeAllSidebarSections() {


document
    .querySelectorAll(
        ".section-title[data-target], .section-toggle[data-target]"
    )
    .forEach(function (header) {

        const targetId =
            header.getAttribute("data-target");

        if (!targetId)
            return;


        const target =
            document.getElementById(targetId);


        if (target) {

            target.classList.remove(
                "open"
            );

        }


        header.classList.remove(
            "open"
        );

    });


}

// ======================================
// Projects
// Render Projects
// ======================================

function renderProjects() {


if (!projectsList)
    return;


projectsList.innerHTML = "";


projects.forEach(function (project) {

    const item =
        document.createElement("div");


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
        function (e) {

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
            function (e) {

                e.stopPropagation();


                document
                    .querySelectorAll(
                        ".project-options-menu"
                    )
                    .forEach(function (m) {

                        m.classList.remove(
                            "open"
                        );

                    });


                const options =
                    item.querySelector(
                        ".project-options-menu"
                    );


                if (!options)
                    return;


                options.classList.add(
                    "open"
                );


                const rect =
                    menu.getBoundingClientRect();


                const menuWidth = 140;

                const menuHeight =
                    options.offsetHeight || 80;

                const margin = 8;


                const viewportWidth =
                    window.innerWidth;

                const viewportHeight =
                    window.innerHeight;


                const spaceLeft =
                    rect.left;

                const spaceBelow =
                    viewportHeight -
                    rect.bottom;

                const spaceAbove =
                    rect.top;


                let left;
                let top;


                if (
                    spaceLeft >=
                    menuWidth + margin
                ) {

                    left =
                        rect.left -
                        menuWidth -
                        margin;

                }
                else {

                    left =
                        rect.right +
                        margin;

                }


                if (
                    spaceBelow <
                        menuHeight + margin &&
                    spaceAbove >=
                        menuHeight + margin
                ) {

                    top =
                        rect.top -
                        menuHeight -
                        margin;

                }
                else {

                    top =
                        rect.bottom +
                        margin;

                }


                if (left < margin) {

                    left =
                        margin;

                }


                if (
                    left + menuWidth >
                    viewportWidth - margin
                ) {

                    left =
                        viewportWidth -
                        menuWidth -
                        margin;

                }


                if (top < margin) {

                    top =
                        margin;

                }


                if (
                    top + menuHeight >
                    viewportHeight - margin
                ) {

                    top =
                        viewportHeight -
                        menuHeight -
                        margin;

                }


                options.style.position =
                    "fixed";

                options.style.left =
                    left + "px";

                options.style.top =
                    top + "px";

                options.style.right =
                    "auto";

                options.style.bottom =
                    "auto";

                options.style.zIndex =
                    "999999";

            };

    }


    const renameProject =
        item.querySelector(
            ".rename-project"
        );


    if (renameProject) {

        renameProject.onclick =
            function (e) {

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


                const title =
                    item.querySelector(
                        ".project-title"
                    );


                if (!title)
                    return;


                const oldName =
                    project.name;


                title.innerHTML = `

                    <input
                        class="edit-project-title"
                        value="${oldName}">

                `;


                const edit =
                    title.querySelector(
                        ".edit-project-title"
                    );


                if (!edit)
                    return;


                edit.focus();


                edit.setSelectionRange(
                    edit.value.length,
                    edit.value.length
                );


                edit.onkeydown =
                    function (event) {

                        if (
                            event.key ===
                            "Enter"
                        ) {

                            const value =
                                edit.value.trim();


                            if (value !== "") {

                                project.name =
                                    value;

                            }
                            else {

                                project.name =
                                    oldName;

                            }


                            saveProjects();

                            renderProjects();

                            renderExpandedProjects();

                        }


                        if (
                            event.key ===
                            "Escape"
                        ) {

                            project.name =
                                oldName;

                            renderProjects();

                        }

                    };

            };

    }


    const deleteProject =
        item.querySelector(
            ".delete-project"
        );


    if (deleteProject) {

        deleteProject.onclick =
            function (e) {

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
                        function () {

                            projects =
                                projects.filter(
                                    function (p) {

                                        return (
                                            p.id !==
                                            project.id
                                        );

                                    }
                                );


                            if (
                                currentProject &&
                                currentProject.id ===
                                project.id
                            ) {

                                currentProject =
                                    null;

                            }


                            saveProjects();

                            renderProjects();

                            renderExpandedProjects();

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

});


}

// ======================================
// Projects Button
// ======================================

if (projectsBtn) {


projectsBtn.onclick =
    function (e) {

        e.stopPropagation();


        if (!projectsPopup)
            return;


        if (chatPopup)
            chatPopup.classList.remove(
                "open"
            );


        if (searchPopup)
            searchPopup.classList.remove(
                "open"
            );


        if (settingsWindow)
            settingsWindow.classList.remove(
                "open"
            );


        projectsPopup.classList.toggle(
            "open"
        );


        renderProjects();

    };


}

// ======================================
// New Project
// ======================================

if (newProjectBtn) {


newProjectBtn.onclick =
    function (e) {

        e.stopPropagation();


        if (!projectsPopup)
            return;


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


        const boxWidth = 240;

        const boxHeight = 120;

        const screenMargin = 12;


        let left =
            buttonRect.left;

        let top =
            buttonRect.bottom + 8;


        const actualBoxWidth =
            box.offsetWidth ||
            boxWidth;


        if (
            left + actualBoxWidth >
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
            top + boxHeight >
            window.innerHeight -
            screenMargin
        ) {

            top =
                buttonRect.top -
                boxHeight -
                8;

        }


        if (left < 10) {

            left = 10;

        }


        box.style.position =
            "fixed";

        box.style.left =
            left + "px";

        box.style.top =
            top + "px";

        box.style.zIndex =
            "999999";


        const inputProject =
            box.querySelector(
                ".new-project-name"
            );


        if (inputProject) {

            inputProject.focus();

        }


        const saveProject =
            box.querySelector(
                ".save-project"
            );


        if (saveProject) {

            saveProject.onclick =
                function () {

                    const name =
                        inputProject
                            ? inputProject.value.trim()
                            : "";


                    if (name !== "") {

                        const now =
    new Date().toISOString();


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
            function (e) {

                e.stopPropagation();

            };

    };


}

// ======================================
// Expanded Sidebar
// القائمة الجانبية المنبثقة
// ======================================

if (sidebarToggleBtn && expandedSidebar) {


sidebarToggleBtn.onclick = function (e) {

    e.preventDefault();
    e.stopPropagation();

    if (
        !expandedSidebar ||
        !sidebarToggleBtn ||
        !expandedSidebarToggleSlot
    ) {
        return;
    }

    const isOpening =
        !expandedSidebar.classList.contains("open");


    /* =========================================
       فتح القائمة
       ========================================= */

    if (isOpening) {

        expandedSidebar.classList.add("open");

        document.body.classList.add(
            "expanded-sidebar-open"
        );

        sidebarToggleBtn.title =
            "إخفاء القائمة";

        sidebarToggleBtn.classList.add(
            "sidebar-open"
        );

        /* نقل نفس الزر إلى داخل القائمة */
        expandedSidebarToggleSlot.appendChild(
            sidebarToggleBtn
        );

    }


    /* =========================================
       إغلاق القائمة
       ========================================= */

    else {

        expandedSidebar.classList.remove("open");

        document.body.classList.remove(
            "expanded-sidebar-open"
        );

        sidebarToggleBtn.title =
            "إظهار القائمة";

        sidebarToggleBtn.classList.remove(
            "sidebar-open"
        );

        /* إعادة الزر إلى مكانه الأصلي */
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

// ======================================
// Render Projects in Expanded Sidebar
// ======================================

function renderExpandedProjects() {


const list =
    document.getElementById(
        "expanded-projects-list"
    );


if (!list)
    return;


list.innerHTML =
    "";


projects.forEach(
    function (project) {

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
        function (e) {

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

// ======================================
// Sidebar — Recent Chats under New Chat
// نفس آخر 8 محادثات التي يعرضها نظام المحادثات
// ======================================

function renderSidebarChats() {


const list =
    document.getElementById(
        "new-chat-list"
    );

if (!list)
    return;

list.innerHTML = "";

// نفس منطق آخر 8 محادثات
if (chats.length === 0) {

    list.innerHTML = `
        <div class="empty-chat">
            لا توجد محادثات
        </div>
    `;

    return;
}

chats
    .slice(0, 8)
    .forEach(function (chat) {

        const item =
            document.createElement("div");

        item.className =
            "recent-chat-item";

        item.innerHTML = `
        <span class="chat-title">
            ${chatIcon}
            ${chat.title}
        </span>
        `;

        // فتح المحادثة عند الضغط
        item.onclick =
            function (e) {

                e.stopPropagation();

                currentChat =
                    chat;

                renderChat();

                // إغلاق القائمة المنسدلة إن كانت مفتوحة
                if (projectsPopup) {

                    projectsPopup
                        .classList
                        .remove("open");

                }

                if (chatPopup) {

                    chatPopup
                        .classList
                        .remove("open");

                }

                if (searchPopup) {

                    searchPopup
                        .classList
                        .remove("open");

                }

            };

        list.appendChild(item);

    });


}



// ======================================
// Create New Chat
// ======================================

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


// لا نضيف المحادثة إلى chats هنا
// لأنها لم تصبح محادثة فعلية بعد


// تحديث قائمة آخر 8 محادثات فقط
renderSidebarChats();

renderRecentChats();


// تنظيف مربع الإدخال
if (input) {

    input.value =
        "";

    input.style.height =
        "auto";

}


// عرض شاشة المحادثة الجديدة
if (chatArea) {

    chatArea.innerHTML = `

        <div class="welcome">

            <div class="ai-symbol">
                ✦
            </div>

            <h2>
                مرحبًا بك
            </h2>

            <p>
                ابدأ محادثة جديدة
            </p>

        </div>

    `;

}


}

// ======================================
// New Chat Button
// ======================================

if (newChatBtn) {


newChatBtn.onclick =
    function (e) {

        e.stopPropagation();


        if (projectsPopup)
            projectsPopup.classList.remove(
                "open"
            );


        if (chatPopup)
            chatPopup.classList.remove(
                "open"
            );


        if (searchPopup)
            searchPopup.classList.remove(
                "open"
            );


        if (settingsWindow)
            settingsWindow.classList.remove(
                "open"
            );


        createNewChat();

    };


}

// =====================================================
// =====================================================
// AI SYSTEM
// التعديل يبدأ من هنا
// =====================================================
// =====================================================

// ======================================
// قراءة إعدادات الذكاء الاصطناعي
// ======================================

function getSavedSettings() {


try {

    return (
        JSON.parse(
            localStorage.getItem(
                "AI_SETTINGS"
            )
        ) || {}
    );

}
catch (e) {

    return {};

}


}

// ======================================
// حفظ إعدادات الذكاء الاصطناعي
// ======================================

function saveAISettings(data) {


localStorage.setItem(
    "AI_SETTINGS",
    JSON.stringify(data)
);


}

// ======================================
// معلومات المزود
// ======================================

function updateProviderInfo() {


if (!providerInfo || !provider)
    return;


const value =
    provider.value;


if (value === "openrouter") {

    providerInfo.innerHTML =
        "OpenRouter: سيتم جلب النماذج المجانية المتاحة من حسابك.";

    return;

}


if (value === "openai") {

    providerInfo.innerHTML =
        "OpenAI: سيتم جلب النماذج المتاحة من حسابك.";

    return;

}


if (value === "gemini") {

    providerInfo.innerHTML =
        "Gemini: سيتم جلب النماذج التي تدعم generateContent.";

    return;

}


providerInfo.innerHTML =
    "سيتم تحديد رابط الاتصال حسب مزود الذكاء الاصطناعي.";


}

// ======================================
// تحميل الإعدادات
// ======================================

function loadSettings() {


const data =
    getSavedSettings();


if (provider) {

    provider.value =
        data.provider ||
        "openrouter";

}


if (apiKey) {

    apiKey.value =
        data.key ||
        "";

}


if (modelSelect) {

    const savedModel =
        data.model ||
        "";


    modelSelect.innerHTML =
        "";


    if (savedModel !== "") {

        const option =
            document.createElement(
                "option"
            );


        option.value =
            savedModel;


        option.textContent =
            savedModel;


        modelSelect.appendChild(
            option
        );


        modelSelect.value =
            savedModel;

    }
    else {

        const option =
            document.createElement(
                "option"
            );


        option.value =
            "";


        option.textContent =
            "أدخل المفتاح ثم حدّث النماذج";


        modelSelect.appendChild(
            option
        );

    }

}


updateProviderInfo();


}

// ======================================
// Settings Button
// ======================================

if (settingsBtn) {


settingsBtn.onclick =
    function (e) {

        e.preventDefault();
        e.stopPropagation();


        if (projectsPopup)
            projectsPopup.classList.remove(
                "open"
            );


        if (chatPopup)
            chatPopup.classList.remove(
                "open"
            );


        if (searchPopup)
            searchPopup.classList.remove(
                "open"
            );


        if (settingsWindow) {

            settingsWindow.classList.add(
                "open"
            );

        }


        loadSettings();

    };


}

// ======================================
// Provider Change
// ======================================

if (provider) {


provider.onchange =
    function () {

        updateProviderInfo();


        if (modelSelect) {

            modelSelect.innerHTML = `

                <option value="">
                    أدخل المفتاح ثم حدّث النماذج
                </option>

            `;

        }

        if (settingsStatus) {

            settingsStatus.innerHTML =
                "";

        }

    };


}

// ======================================
// Close Settings
// ======================================

if (closeSettings) {


closeSettings.onclick =
    function (e) {

        e.preventDefault();
        e.stopPropagation();


        if (settingsWindow) {

            settingsWindow.classList.remove(
                "open"
            );

        }

    };


}

// ======================================
// Show / Hide API Key
// ======================================

if (showKey && apiKey) {


showKey.onclick =
    function (e) {

        e.preventDefault();
        e.stopPropagation();


        if (
            apiKey.type ===
            "password"
        ) {

            apiKey.type =
                "text";

            showKey.innerHTML =
                "🙈";

        }
        else {

            apiKey.type =
                "password";

            showKey.innerHTML =
                "👁";

        }

    };


}

// ======================================
// Save Settings
// ======================================

if (saveSettings) {


saveSettings.onclick =
    function (e) {

        e.preventDefault();
        e.stopPropagation();


        const settings = {

            provider:
                provider
                    ? provider.value
                    : "openrouter",

            key:
                apiKey
                    ? apiKey.value.trim()
                    : "",

            model:
                modelSelect
                    ? modelSelect.value.trim()
                    : ""

        };


        if (!settings.key) {

            if (settingsStatus) {

                settingsStatus.innerHTML =
                    "⚠ يرجى إدخال مفتاح API.";

            }

            return;

        }


        if (!settings.model) {

            if (settingsStatus) {

                settingsStatus.innerHTML =
                    "⚠ يرجى تحديد نموذج الذكاء الاصطناعي.";

            }

            return;

        }


        saveAISettings(
            settings
        );


        if (settingsStatus) {

            settingsStatus.innerHTML =
                "✓ تم حفظ إعدادات الذكاء الاصطناعي";

        }

    };


}

// =====================================================
// AI MODELS
// =====================================================

// ======================================
// تحديث النماذج
// ======================================

if (refreshModels) {


refreshModels.onclick =
    async function (e) {

        e.preventDefault();
        e.stopPropagation();


        refreshModels.disabled =
            true;


        try {

            await loadModels();

        }

        catch (error) {

            if (settingsStatus) {

                settingsStatus.innerHTML =
                    "⚠ " +
                    (
                        error.message ||
                        "تعذر تحديث النماذج"
                    );

            }

        }

        finally {

            refreshModels.disabled =
                false;

        }

    };


}

// ======================================
// الدالة الرئيسية لتحميل النماذج
// ======================================

async function loadModels() {


const selectedProvider =
    provider
        ? provider.value
        : "openrouter";


if (selectedProvider === "openrouter") {

    await loadOpenRouterModels();

    return;

}


if (selectedProvider === "openai") {

    await loadOpenAIModels();

    return;

}


if (selectedProvider === "gemini") {

    await loadGeminiModels();

    return;

}


throw new Error(
    "مزود الذكاء الاصطناعي غير معروف."
);


}

// ======================================
// OpenRouter Models
// ======================================

async function loadOpenRouterModels() {


const key =
    apiKey
        ? apiKey.value.trim()
        : "";


if (!key) {

    throw new Error(
        "يرجى إدخال مفتاح OpenRouter أولاً."
    );

}


if (settingsStatus) {

    settingsStatus.innerHTML =
        "⏳ جاري تحميل نماذج OpenRouter المجانية...";

}


const response =
    await fetch(
        "https://openrouter.ai/api/v1/models",
        {

            method:
                "GET",

            headers: {

                "Authorization":
                    "Bearer " + key,

                "Content-Type":
                    "application/json",

                "HTTP-Referer":
                    window.location.href,

                "X-Title":
                    "Research Tools"

            }

        }
    );


const result =
    await readJSON(response);


if (!response.ok) {

    throw new Error(
        getAPIError(
            result,
            "فشل الاتصال بـ OpenRouter."
        )
    );

}


if (
    !result.data ||
    !Array.isArray(result.data)
) {

    throw new Error(
        "لم تصل قائمة النماذج من OpenRouter."
    );

}


const freeModels =
    result.data.filter(
        function (item) {

            if (
                !item ||
                !item.id
            ) {

                return false;

            }


            if (
                !item.pricing
            ) {

                return false;

            }


            return (
                String(item.pricing.prompt) === "0" &&
                String(item.pricing.completion) === "0"
            );

        }
    );


populateModels(
    freeModels.map(
        function (item) {

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


if (settingsStatus) {

    settingsStatus.innerHTML =
        "✓ تم تحديث النماذج المجانية: " +
        freeModels.length;

}


}

// ======================================
// OpenAI Models
// ======================================

async function loadOpenAIModels() {


const key =
    apiKey
        ? apiKey.value.trim()
        : "";


if (!key) {

    throw new Error(
        "يرجى إدخال مفتاح OpenAI أولاً."
    );

}


if (settingsStatus) {

    settingsStatus.innerHTML =
        "⏳ جاري تحميل نماذج OpenAI...";

}


const response =
    await fetch(
        "https://api.openai.com/v1/models",
        {

            method:
                "GET",

            headers: {

                "Authorization":
                    "Bearer " + key,

                "Content-Type":
                    "application/json"

            }

        }
    );


const result =
    await readJSON(response);


if (!response.ok) {

    throw new Error(
        getAPIError(
            result,
            "فشل الاتصال بـ OpenAI."
        )
    );

}


if (
    !result.data ||
    !Array.isArray(result.data)
) {

    throw new Error(
        "لم تصل قائمة نماذج OpenAI."
    );

}


const models =
    result.data.filter(
        function (item) {

            if (!item || !item.id)
                return false;


            const id =
                item.id.toLowerCase();


            return (
                id.startsWith("gpt-") ||
                id.startsWith("o1") ||
                id.startsWith("o3") ||
                id.startsWith("o4")
            );

        }
    );


models.sort(
    function (a, b) {

        return a.id.localeCompare(
            b.id
        );

    }
);


populateModels(
    models.map(
        function (item) {

            return {

                id:
                    item.id,

                name:
                    item.id

            };

        }
    )
);


if (settingsStatus) {

    settingsStatus.innerHTML =
        "✓ تم تحديث نماذج OpenAI: " +
        models.length;

}


}

// ======================================
// Gemini Models
// ======================================

async function loadGeminiModels() {


const key =
    apiKey
        ? apiKey.value.trim()
        : "";


if (!key) {

    throw new Error(
        "يرجى إدخال مفتاح Gemini أولاً."
    );

}


if (settingsStatus) {

    settingsStatus.innerHTML =
        "⏳ جاري تحميل نماذج Gemini...";

}


const url =
    "https://generativelanguage.googleapis.com/v1beta/models?key=" +
    encodeURIComponent(key);


const response =
    await fetch(
        url,
        {

            method:
                "GET"

        }
    );


const result =
    await readJSON(response);


if (!response.ok) {

    throw new Error(
        getAPIError(
            result,
            "فشل الاتصال بـ Gemini."
        )
    );

}


if (
    !result.models ||
    !Array.isArray(result.models)
) {

    throw new Error(
        "لم تصل قائمة نماذج Gemini."
    );

}


const models =
    result.models.filter(
        function (item) {

            if (!item)
                return false;


            if (!item.name)
                return false;


            if (
                !item.supportedGenerationMethods
            ) {

                return false;

            }


            return (
                item.supportedGenerationMethods
                    .includes(
                        "generateContent"
                    )
            );

        }
    );


populateModels(
    models.map(
        function (item) {

            const cleanId =
                String(item.name)
                    .replace(
                        /^models\//,
                        ""
                    );


            return {

                id:
                    cleanId,

                name:
                    item.displayName
                        ? item.displayName +
                          " — " +
                          cleanId
                        : cleanId

            };

        }
    )
);


if (settingsStatus) {

    settingsStatus.innerHTML =
        "✓ تم تحديث نماذج Gemini: " +
        models.length;

}


}

// ======================================
// وضع النماذج داخل القائمة
// ======================================

function populateModels(models) {


if (!modelSelect)
    return;


const saved =
    getSavedSettings();


const savedModel =
    saved.model ||
    "";


modelSelect.innerHTML =
    "";


if (
    !models ||
    models.length === 0
) {

    const option =
        document.createElement(
            "option"
        );


    option.value =
        "";


    option.textContent =
        "لا توجد نماذج متاحة";


    modelSelect.appendChild(
        option
    );


    return;

}


models.forEach(
    function (item) {

        const option =
            document.createElement(
                "option"
            );


        option.value =
            item.id;


        option.textContent =
            item.name;


        modelSelect.appendChild(
            option
        );

    }
);


if (savedModel) {

    const exists =
        Array.from(
            modelSelect.options
        ).some(
            function (option) {

                return (
                    option.value ===
                    savedModel
                );

            }
        );


    if (exists) {

        modelSelect.value =
            savedModel;

    }

}


}

// =====================================================
// AI CONNECTION TEST
// =====================================================

async function testAIConnection() {


const data = {

    provider:
        provider
            ? provider.value
            : "openrouter",

    key:
        apiKey
            ? apiKey.value.trim()
            : "",

    model:
        modelSelect
            ? modelSelect.value.trim()
            : ""

};


if (!data.key) {

    throw new Error(
        "يرجى إدخال مفتاح API أولاً."
    );

}


if (!data.model) {

    throw new Error(
        "يرجى تحديد نموذج الذكاء الاصطناعي أولاً."
    );

}


// --------------------------------------
// OpenRouter
// --------------------------------------

if (
    data.provider ===
    "openrouter"
) {

    const response =
        await fetch(
            "https://openrouter.ai/api/v1/chat/completions",
            {

                method:
                    "POST",

                headers: {

                    "Authorization":
                        "Bearer " +
                        data.key,

                    "Content-Type":
                        "application/json",

                    "HTTP-Referer":
                        window.location.href,

                    "X-Title":
                        "Research Tools"

                },

                body:
                    JSON.stringify({

                        model:
                            data.model,

                        messages: [

                            {

                                role:
                                    "user",

                                content:
                                    "أجب بكلمة واحدة فقط: متصل"

                            }

                        ],

                        max_tokens:
                            10

                    })

            }
        );


    const result =
        await readJSON(response);


    if (!response.ok) {

        throw new Error(
            getAPIError(
                result,
                "فشل الاتصال بـ OpenRouter."
            )
        );

    }


    return (
        "✓ تم الاتصال بـ OpenRouter بنجاح"
    );

}


// --------------------------------------
// OpenAI
// --------------------------------------

if (
    data.provider ===
    "openai"
) {

    const response =
        await fetch(
            "https://api.openai.com/v1/chat/completions",
            {

                method:
                    "POST",

                headers: {

                    "Authorization":
                        "Bearer " +
                        data.key,

                    "Content-Type":
                        "application/json"

                },

                body:
                    JSON.stringify({

                        model:
                            data.model,

                        messages: [

                            {

                                role:
                                    "user",

                                content:
                                    "أجب بكلمة واحدة فقط: متصل"

                            }

                        ],

                        max_tokens:
                            10

                    })

            }
        );


    const result =
        await readJSON(response);


    if (!response.ok) {

        throw new Error(
            getAPIError(
                result,
                "فشل الاتصال بـ OpenAI."
            )
        );

    }


    return (
        "✓ تم الاتصال بـ OpenAI بنجاح"
    );

}


// --------------------------------------
// Gemini
// --------------------------------------

if (
    data.provider ===
    "gemini"
) {

    const model =
        normalizeGeminiModel(
            data.model
        );


    const url =
        "https://generativelanguage.googleapis.com/v1beta/models/" +
        encodeURIComponent(model) +
        ":generateContent?key=" +
        encodeURIComponent(data.key);


    const response =
        await fetch(
            url,
            {

                method:
                    "POST",

                headers: {

                    "Content-Type":
                        "application/json"

                },

                body:
                    JSON.stringify({

                        contents: [

                            {

                                parts: [

                                    {

                                        text:
                                            "أجب بكلمة واحدة فقط: متصل"

                                    }

                                ]

                            }

                        ]

                    })

            }
        );


    const result =
        await readJSON(response);


    if (!response.ok) {

        throw new Error(
            getAPIError(
                result,
                "فشل الاتصال بـ Gemini."
            )
        );

    }


    return (
        "✓ تم الاتصال بـ Gemini بنجاح"
    );

}


throw new Error(
    "مزود الذكاء الاصطناعي غير معروف."
);


}

// ======================================
// Test Connection Button
// ======================================

if (testConnection) {


testConnection.onclick =
    async function (e) {

        e.preventDefault();
        e.stopPropagation();


        testConnection.disabled =
            true;


        if (settingsStatus) {

            settingsStatus.innerHTML =
                "⏳ جاري اختبار الاتصال...";

        }


        try {

            const message =
                await testAIConnection();


            if (settingsStatus) {

                settingsStatus.innerHTML =
                    message;

            }

        }

        catch (error) {

            if (settingsStatus) {

                settingsStatus.innerHTML =
                    "⚠ " +
                    (
                        error.message ||
                        "تعذر الاتصال"
                    );

            }

        }

        finally {

            testConnection.disabled =
                false;

        }

    };


}

// =====================================================
// AI REQUEST
// إرسال المحادثة كاملة مع الرسالة الجديدة
// =====================================================

async function askAI(text) {

    const data =
        getSavedSettings();

    const selectedProvider =
        (
            data.provider ||
            "openrouter"
        ).toLowerCase();

    const key =
        data.key ||
        "";

    const model =
        data.model ||
        "";


    if (!key.trim()) {

        throw new Error(
            "لم يتم إدخال مفتاح الذكاء الاصطناعي من الإعدادات."
        );

    }


    if (!model.trim()) {

        throw new Error(
            "لم يتم تحديد نموذج الذكاء الاصطناعي من الإعدادات."
        );

    }


    // =================================================
    // بناء سياق المحادثة
    // =================================================

    const conversationMessages = [];


    if (
        currentChat &&
        Array.isArray(currentChat.messages)
    ) {

        currentChat.messages.forEach(
            function (msg) {

                if (
                    !msg ||
                    !msg.text
                ) {

                    return;

                }


                conversationMessages.push({

                    role:
                        msg.role === "ai"
                            ? "assistant"
                            : "user",

                    content:
                        String(msg.text)

                });

            }
        );

    }


    // =================================================
    // إضافة السؤال الجديد
    // =================================================

    conversationMessages.push({

        role:
            "user",

        content:
            text

    });


    // ==================================
    // OpenRouter
    // ==================================

    if (
        selectedProvider ===
        "openrouter"
    ) {

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
                            key

                    },

                    body:
                        JSON.stringify({

                            model:
                                model,

                            messages:
                                conversationMessages

                        })

                }
            );


        const result =
            await readJSON(
                response
            );


        if (!response.ok) {

            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ OpenRouter."
                )
            );

        }


        return extractOpenAIStyleAnswer(
            result,
            "OpenRouter"
        );

    }


    // ==================================
    // OpenAI
    // ==================================

    if (
        selectedProvider ===
        "openai"
    ) {

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
                                conversationMessages

                        })

                }
            );


        const result =
            await readJSON(
                response
            );


        if (!response.ok) {

            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ OpenAI."
                )
            );

        }


        return extractOpenAIStyleAnswer(
            result,
            "OpenAI"
        );

    }


    // ==================================
    // Gemini
    // ==================================

    if (
        selectedProvider ===
        "gemini"
    ) {

        const cleanModel =
            normalizeGeminiModel(
                model
            );


        const url =
            "https://generativelanguage.googleapis.com/v1beta/models/" +
            encodeURIComponent(
                cleanModel
            ) +
            ":generateContent?key=" +
            encodeURIComponent(
                key
            );


        // Gemini يستخدم contents بدل messages
        const contents =
            conversationMessages.map(
                function (msg) {

                    return {

                        role:
                            msg.role === "assistant"
                                ? "model"
                                : "user",

                        parts: [

                            {

                                text:
                                    msg.content

                            }

                        ]

                    };

                }
            );


        const response =
            await fetch(
                url,
                {

                    method:
                        "POST",

                    headers: {

                        "Content-Type":
                            "application/json"

                    },

                    body:
                        JSON.stringify({

                            contents:
                                contents

                        })

                }
            );


        const result =
            await readJSON(
                response
            );


        if (!response.ok) {

            throw new Error(
                getAPIError(
                    result,
                    "فشل الاتصال بـ Gemini."
                )
            );

        }


        return extractGeminiAnswer(
            result
        );

    }


    throw new Error(
        "مزود الذكاء الاصطناعي غير معروف: " +
        selectedProvider
    );

}

// ======================================
// Gemini Model Normalization
// ======================================

function normalizeGeminiModel(model) {


return String(
    model || ""
).replace(
    /^models\//,
    ""
);


}

// ======================================
// قراءة JSON بأمان
// ======================================

async function readJSON(response) {


try {

    return await response.json();

}
catch (e) {

    return {};

}


}

// ======================================
// استخراج رسالة الخطأ
// ======================================

function getAPIError(
result,
fallback
) {


if (
    result &&
    result.error
) {

    if (
        typeof result.error ===
        "string"
    ) {

        return result.error;

    }


    if (
        result.error.message
    ) {

        return result.error.message;

    }

}


return fallback;


}

// ======================================
// استخراج رد OpenAI / OpenRouter
// ======================================

function extractOpenAIStyleAnswer(
result,
providerName
) {


if (
    result &&
    result.choices &&
    result.choices.length > 0
) {

    const choice =
        result.choices[0];


    if (
        choice.message &&
        typeof choice.message.content ===
        "string"
    ) {

        return choice.message.content;

    }

}


throw new Error(
    "لم يصل رد صالح من " +
    providerName +
    "."
);


}

// ======================================
// استخراج رد Gemini
// ======================================

function extractGeminiAnswer(result) {


if (
    result &&
    result.candidates &&
    result.candidates.length > 0
) {

    const candidate =
        result.candidates[0];


    if (
        candidate.content &&
        candidate.content.parts &&
        Array.isArray(
            candidate.content.parts
        )
    ) {

        const textParts =
            candidate.content.parts
                .filter(
                    function (part) {

                        return (
                            part &&
                            typeof part.text ===
                            "string"
                        );

                    }
                )
                .map(
                    function (part) {

                        return part.text;

                    }
                );


        if (textParts.length > 0) {

            return textParts.join(
                "\n"
            );

        }

    }

}


throw new Error(
    "لم يصل رد صالح من Gemini."
);


}

// =====================================================
// =====================================================
// باقي النظام — كما هو
// =====================================================
// =====================================================
function formatAIMessage(text) {

    if (!text)
        return "";

    try {

        return marked.parse(text, {
            breaks: true,
            gfm: true
        });

    }

    catch (error) {

        console.error(
            "Markdown formatting error:",
            error
        );

        return String(text)
            .replace(/\n/g, "<br>");

    }
}
// ======================================
// Render Chat
// ======================================

function renderChat() {

if (!chatArea)
    return;


chatArea.innerHTML =
    "";


if (!currentChat) {

    chatArea.innerHTML = `

        <div class="welcome">

            <div class="ai-symbol">
                ✦
            </div>

            <h2>
                مرحبًا بك
            </h2>

            <p>
                ابدأ محادثة جديدة
            </p>

        </div>

    `;


    return;

}


currentChat.messages.forEach(
    function (msg) {

        const div =
            document.createElement(
                "div"
            );


        div.className =
            "message " +
            (
                msg.role === "user"
                    ? "user-message"
                    : "ai-message"
            );


        if (msg.role === "user") {

    div.textContent =
        msg.text || "";

}
else {

    div.innerHTML =
        formatAIMessage(
            msg.text || ""
        );

}


        chatArea.appendChild(
            div
        );

    }
);


chatArea.scrollTop =
    chatArea.scrollHeight;


}

// ======================================
// Render Chat List
// ======================================

function renderChatList() {


const list =
    document.getElementById(
        "chat-list"
    );


if (!list)
    return;


list.innerHTML =
    "";


if (chats.length === 0) {

    list.innerHTML =
        "<div class='empty-chat'>لا توجد محادثات</div>";

    return;

}


chats
.forEach(function (chat) {

        const item =
            document.createElement(
                "div"
            );


        item.className =
            "chat-history-item";


        item.innerHTML = `
    <span class="chat-title">
        ${chatIcon}
        ${chat.title}
    </span>

    <button
        class="chat-menu"
        type="button">
        ⋮
    </button>

    <div class="chat-options-menu">

        <div class="rename-chat">
            ✏ إعادة تسمية
        </div>

        <div class="delete-chat">
            🗑 حذف
        </div>

    </div>
`;


        const title =
            item.querySelector(
                ".chat-title"
            );


        if (title) {

            title.onclick =
                function (e) {

                    e.stopPropagation();


                    currentChat =
                        chat;


                    renderChat();


                    if (
                        expandedSidebar
                    ) {

                        expandedSidebar.classList.remove(
                            "open"
                        );

                    }

                };

        }


        const menu =
            item.querySelector(
                ".chat-menu"
            );


        if (menu) {

            menu.onclick =
                function (e) {

                    e.stopPropagation();


                    document
                        .querySelectorAll(
                            ".chat-options-menu"
                        )
                        .forEach(
                            function (m) {

                                m.classList.remove(
                                    "open"
                                );

                            }
                        );


                    const options =
                        item.querySelector(
                            ".chat-options-menu"
                        );


                    if (!options)
                        return;


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


                    const viewportWidth =
                        window.innerWidth;

                    const viewportHeight =
                        window.innerHeight;


                    const spaceLeft =
                        rect.left;

                    const spaceAbove =
                        rect.top;

                    const spaceBelow =
                        viewportHeight -
                        rect.bottom;


                    let left;
                    let top;


                    if (
                        spaceLeft >=
                        menuWidth + margin
                    ) {

                        left =
                            rect.left -
                            menuWidth -
                            margin;

                    }
                    else {

                        left =
                            rect.right +
                            margin;

                    }


                    if (
                        spaceBelow <
                            menuHeight + margin &&
                        spaceAbove >=
                            menuHeight + margin
                    ) {

                        top =
                            rect.top -
                            menuHeight -
                            margin;

                    }
                    else {

                        top =
                            rect.bottom +
                            margin;

                    }


                    if (left < margin) {

                        left =
                            margin;

                    }


                    if (
                        left + menuWidth >
                        viewportWidth - margin
                    ) {

                        left =
                            viewportWidth -
                            menuWidth -
                            margin;

                    }


                    if (top < margin) {

                        top =
                            margin;

                    }


                    if (
                        top + menuHeight >
                        viewportHeight - margin
                    ) {

                        top =
                            viewportHeight -
                            menuHeight -
                            margin;

                    }


                    options.style.position =
                        "fixed";

                    options.style.left =
                        left + "px";

                    options.style.top =
                        top + "px";

                    options.style.right =
                        "auto";

                    options.style.bottom =
                        "auto";

                    options.style.zIndex =
                        "999999";

                };

        }


        const renameBtn =
            item.querySelector(
                ".rename-chat"
            );


        if (renameBtn) {

            renameBtn.onclick =
                function (e) {

                    e.stopPropagation();


                    const options =
                        item.querySelector(
                            ".chat-options-menu"
                        );


                    if (options) {

                        options.classList.remove(
                            "open"
                        );

                    }


                    const titleSpan =
                        item.querySelector(
                            ".chat-title"
                        );


                    if (!titleSpan)
                        return;


                    const oldName =
                        chat.title;


                    titleSpan.innerHTML = `

                        <input
                            class="edit-chat-title"
                            value="${oldName}">

                    `;


                    const editInput =
                        titleSpan.querySelector(
                            ".edit-chat-title"
                        );


                    if (!editInput)
                        return;


                    editInput.focus();


                    editInput.setSelectionRange(
                        editInput.value.length,
                        editInput.value.length
                    );


                    editInput.onkeydown =
                        function (event) {

                            if (
                                event.key ===
                                "Enter"
                            ) {

                                const value =
                                    editInput.value.trim();


                                if (value !== "") {

                                    chat.title =
                                        value;

                                }
                                else {

                                    chat.title =
                                        oldName;

                                }


                                saveChats();

                                renderChatList();

                                renderSidebarChats();

                                renderRecentChats();

                            }


                            if (
                                event.key ===
                                "Escape"
                            ) {

                                chat.title =
                                    oldName;


                                renderChatList();

                            }

                        };

                };

        }


        const deleteBtn =
            item.querySelector(
                ".delete-chat"
            );


        if (deleteBtn) {

            deleteBtn.onclick =
                function (e) {

                    e.stopPropagation();


                    const options =
                        item.querySelector(
                            ".chat-options-menu"
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
                        "delete-confirm";


                    confirmBox.innerHTML = `

                        <div class="confirm-dialog">

                            <p>
                                هل تريد حذف المحادثة:
                                <br>

                                <strong>
                                    ${chat.title}
                                </strong>

                                ؟
                            </p>

                            <button
                                class="confirm-delete"
                                type="button">
                                حذف
                            </button>

                            <button
                                class="cancel-delete"
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
                            ".confirm-delete"
                        );


                    if (confirmDelete) {

                        confirmDelete.onclick =
                            function () {

                                chats =
                                    chats.filter(
                                        function (c) {

                                            return (
                                                c.id !==
                                                chat.id
                                            );

                                        }
                                    );


                                if (
                                    currentChat &&
                                    currentChat.id ===
                                    chat.id
                                ) {

                                    currentChat =
                                        null;

                                    renderChat();

                                }


                                saveChats();

                                renderChatList();

                                renderSidebarChats();

                                renderRecentChats();

                                confirmBox.remove();

                            };

                    }


                    const cancelDelete =
                        confirmBox.querySelector(
                            ".cancel-delete"
                        );


                    if (cancelDelete) {

                        cancelDelete.onclick =
                            function () {

                                confirmBox.remove();

                            };

                    }

                };

        }


        list.appendChild(
            item
        );

    });


}

// ======================================
// Chat Popup
// ======================================

function renderRecentChats() {


if (!recentChatList)
    return;


recentChatList.innerHTML =
    "";


if (chats.length === 0) {

    recentChatList.innerHTML =
        "<div class='empty-chat'>لا توجد محادثات</div>";

    return;

}


chats
    .slice(0, 8)
    .forEach(
        function (chat) {

            const div =
                document.createElement(
                    "div"
                );


            div.className =
                "recent-chat-item";


            div.innerHTML = `

    <span class="chat-title">
        ${chatIcon}
        ${chat.title}
    </span>

`;


            div.onclick =
                function () {

                    currentChat =
                        chat;


                    renderChat();


                    if (chatPopup) {

                        chatPopup.classList.remove(
                            "open"
                        );

                    }

                };


            recentChatList.appendChild(
                div
            );

        }
    );


}

// ======================================
// Chat Button
// ======================================

if (chatBtn) {


chatBtn.onclick =
    function (e) {

        e.stopPropagation();


        if (!chatPopup)
            return;


        if (projectsPopup)
            projectsPopup.classList.remove(
                "open"
            );


        if (searchPopup)
            searchPopup.classList.remove(
                "open"
            );


        if (settingsWindow)
            settingsWindow.classList.remove(
                "open"
            );


        if (
            chatPopup.classList.contains(
                "open"
            )
        ) {

            chatPopup.classList.remove(
                "open"
            );

            return;

        }


        // عرض جميع المحادثات
        renderChatList();


        chatPopup.classList.add(
            "open"
        );

    };


}

// ======================================
// Search
// ======================================

if (searchBtn) {


searchBtn.onclick =
    function (e) {

        e.stopPropagation();


        if (!searchPopup)
            return;


        if (projectsPopup)
            projectsPopup.classList.remove(
                "open"
            );


        if (chatPopup)
            chatPopup.classList.remove(
                "open"
            );


        if (settingsWindow)
            settingsWindow.classList.remove(
                "open"
            );


        searchPopup.classList.toggle(
            "open"
        );


        if (
            searchPopup.classList.contains(
                "open"
            )
        ) {

            if (searchInput) {

                searchInput.focus();

            }

        }

    };


}

// ======================================
// Search Input
// ======================================

if (searchInput) {


searchInput.oninput =
    function () {

        const keyword =
            searchInput.value
                .trim()
                .toLowerCase();


        if (!searchResults)
            return;


        searchResults.innerHTML =
            "";


        if (keyword === "")
            return;


        chats.forEach(
            function (chat) {

                if (
                    chat.title
                        .toLowerCase()
                        .includes(
                            keyword
                        )
                ) {

                    const item =
                        document.createElement(
                            "div"
                        );


                    item.className =
                        "search-result-item";


                    item.innerHTML = `
                    <span class="chat-title">
                        ${chatIcon}
                        ${chat.title}
                    </span>
                    `;


                    item.onclick =
                        function () {

                            currentChat =
                                chat;


                            renderChat();


                            if (searchPopup) {

                                searchPopup
                                    .classList
                                    .remove(
                                        "open"
                                    );

                            }


                            searchInput.value =
                                "";

                            searchResults.innerHTML =
                                "";

                        };


                    searchResults.appendChild(
                        item
                    );

                }

            }
        );

    };


}

// ======================================
// Send Message
// ======================================

async function sendMessage() {

    if (!input)
        return;

    const text =
        input.value.trim();

    if (text === "")
        return;


    // =================================================
    // إنشاء محادثة جديدة إذا لم توجد محادثة حالية
    // =================================================

    if (!currentChat) {

        currentChat = {

            id:
                Date.now(),

            title:
                text.substring(0, 30),

            messages:
                [],

            isTemporary:
                true

        };

    }


    // =================================================
    // تحويل المحادثة المؤقتة إلى محادثة محفوظة
    // =================================================

    if (currentChat.isTemporary) {

        currentChat.isTemporary =
            false;
        currentChat.projectId =
            currentProject
                ? currentProject.id
                : null;
        currentChat.title =
            text.substring(0, 30);


        // منع إضافة نفس المحادثة أكثر من مرة
        const alreadyExists =
            chats.some(function (chat) {

                return (
                    chat.id ===
                    currentChat.id
                );

            });


        if (!alreadyExists) {

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

        // حفظ المحادثة مباشرة
        saveChats();

    }


    // =================================================
    // إضافة رسالة المستخدم
    // =================================================

    currentChat.messages.push({

        role:
            "user",

        text:
            text

    });


    // حفظ بعد رسالة المستخدم مباشرة
    saveChats();


    // =================================================
    // تحديث الواجهات
    // =================================================

    renderChat();

    renderChatList();

    renderSidebarChats();

    renderRecentChats();


    // =================================================
    // تنظيف مربع الإدخال
    // =================================================

    input.value =
        "";

    input.style.height =
        "auto";


    // =================================================
    // رسالة الانتظار
    // =================================================

    const loading =
        document.createElement(
            "div"
        );


    loading.className =
        "message ai-message";


    loading.innerHTML =
        "⏳ جاري التفكير...";


    if (chatArea) {

        chatArea.appendChild(
            loading
        );

        chatArea.scrollTop =
            chatArea.scrollHeight;

    }


    // =================================================
    // إرسال الطلب إلى الذكاء الاصطناعي
    // =================================================

    try {

        const answer =
            await askAI(
                text
            );


        // إزالة رسالة الانتظار
        if (loading) {

            loading.remove();

        }


        // =================================================
        // إضافة رد الذكاء الاصطناعي
        // =================================================

        currentChat.messages.push({

            role:
                "ai",

            text:
                answer || ""

        });


        // =================================================
        // حفظ المحادثة بعد رد الذكاء الاصطناعي
        // =================================================

        saveChats();


        // =================================================
        // تحديث كل القوائم
        // =================================================

        renderChat();

        renderChatList();

        renderSidebarChats();

        renderRecentChats();

    }

    catch (error) {

        // إزالة رسالة الانتظار
        if (loading) {

            loading.remove();

        }


        // =================================================
        // تسجيل الخطأ داخل المحادثة
        // =================================================

        currentChat.messages.push({

            role:
                "ai",

            text:
                "خطأ: " +
                (
                    error.message ||
                    "حدث خطأ غير معروف"
                )

        });


        // حفظ حتى لا تضيع المحادثة عند حدوث خطأ
        saveChats();


        // تحديث الواجهات
        renderChat();

        renderChatList();

        renderSidebarChats();

        renderRecentChats();

    }

}

// ======================================
// Send Button
// ======================================

if (sendBtn) {


sendBtn.onclick =
    function (e) {

        e.preventDefault();

        sendMessage();

    };


}

// ======================================
// Keyboard
// ======================================

if (input) {


input.onkeydown =
    function (e) {

        if (
            e.key === "Enter" &&
            !e.shiftKey
        ) {

            e.preventDefault();

            sendMessage();

        }

    };


}

// ======================================
// Initial Render
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

});


const sidebar = document.querySelector(".sidebar");
const pinSidebar = document.getElementById("pin-sidebar");

if (pinSidebar && sidebar) {

    let sidebarPinned =
        localStorage.getItem("sidebarPinned") === "true";

    if (sidebarPinned) {
        sidebar.classList.add("pinned");
        pinSidebar.classList.add("pinned");

        // ربط تثبيت الشريط بمساحة الشات
        document.body.classList.add("sidebar-is-pinned");
    }

    pinSidebar.addEventListener("click", function (e) {

        e.stopPropagation();

        sidebarPinned = !sidebarPinned;

        sidebar.classList.toggle("pinned", sidebarPinned);
        pinSidebar.classList.toggle("pinned", sidebarPinned);

        // تمديد/انكماش مساحة الشات مع التثبيت
        document.body.classList.toggle(
            "sidebar-is-pinned",
            sidebarPinned
        );

        localStorage.setItem(
            "sidebarPinned",
            sidebarPinned ? "true" : "false"
        );
    });

}
