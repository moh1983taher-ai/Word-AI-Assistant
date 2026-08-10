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

const input =
document.getElementById("user-input");

const sendBtn =
document.getElementById("send-btn");

const chatArea =
document.getElementById("chat-area");

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

let currentProject = null;

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
            ▣ ${project.name}
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
        function () {

            currentProject =
                project;


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

                        projects.unshift({

                            id:
                                Date.now(),

                            name:
                                name,

                            chats:
                                []

                        });


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

    // فتح / إغلاق القائمة
    expandedSidebar.classList.toggle("open");

    // تحديث شكل الزر
    if (expandedSidebar.classList.contains("open")) {

        sidebarToggleBtn.innerHTML = "×";
        sidebarToggleBtn.title = "إخفاء القائمة";

    } else {

        sidebarToggleBtn.innerHTML = "☰";
        sidebarToggleBtn.title = "إظهار القائمة";

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
                ▣ ${project.name}
            </span>

        `;


        item.onclick =
            function (e) {

                e.stopPropagation();


                currentProject =
                    project;

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
            ◯
            <span>
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
        true

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
                        key,

                    

                },

                body:
                    JSON.stringify({

                        model:
                            model,

                        messages: [

                            {

                                role:
                                    "user",

                                content:
                                    text

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

                        messages: [

                            {

                                role:
                                    "user",

                                content:
                                    text

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
        encodeURIComponent(cleanModel) +
        ":generateContent?key=" +
        encodeURIComponent(key);


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

                                role:
                                    "user",

                                parts: [

                                    {

                                        text:
                                            text

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


        div.textContent =
            msg.text || "";


        div.innerHTML =
            div.innerHTML.replace(
                /\n/g,
                "<br>"
            );


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
                ◯ ${chat.title}
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

                ◯
                <span>
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


                    item.innerHTML =
                        "◯ " +
                        chat.title;


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


    if (currentChat.isTemporary) {

currentChat.isTemporary =
    false;

currentChat.title =
    text.substring(0, 30);

saveChats();

renderChatList();

renderSidebarChats();

renderRecentChats();


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


if (chatArea) {

    chatArea.appendChild(
        loading
    );

}


try {

    const answer =
        await askAI(text);


    loading.remove();


    currentChat.messages.push({

        role:
            "ai",

        text:
            answer

    });


    saveChats();


    renderChat();

    renderChatList();

    renderSidebarChats();

    renderRecentChats();

}

catch (error) {

    loading.remove();


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


    saveChats();


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
}

pinSidebar.addEventListener("click", function (e) {

    e.stopPropagation();

    sidebarPinned = !sidebarPinned;

    sidebar.classList.toggle("pinned", sidebarPinned);
    pinSidebar.classList.toggle("pinned", sidebarPinned);

    localStorage.setItem(
        "sidebarPinned",
        sidebarPinned ? "true" : "false"
    );
});


}