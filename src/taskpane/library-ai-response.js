/* =========================================================
 * أدوات الباحث — Gemini response compatibility layer
 *
 * Loaded after taskpane.js so it replaces only the response
 * extractor. Provider requests and DuckAI handling are unchanged.
 * ========================================================= */
(function (global) {
    "use strict";

    global.extractLibraryAIText = function extractLibraryAIText(result, provider) {
        if (provider === "gemini") {
            const candidates = Array.isArray(result && result.candidates)
                ? result.candidates
                : [];

            const blockReason = result &&
                result.promptFeedback &&
                result.promptFeedback.blockReason;

            if (!candidates.length) {
                const reason = blockReason
                    ? "؛ سبب الحجب: " + blockReason
                    : "";
                throw new Error(
                    "لم يُرجع Gemini أي مرشح نصي" + reason +
                    ". مفاتيح الاستجابة: " +
                    Object.keys(result || {}).join(", ")
                );
            }

            let totalTextParts = 0;
            let thoughtTextParts = 0;
            const candidateStates = [];

            for (let i = 0; i < candidates.length; i++) {
                const candidate = candidates[i] || {};
                const parts = candidate.content &&
                    Array.isArray(candidate.content.parts)
                    ? candidate.content.parts
                    : [];

                const finalTexts = [];

                for (let j = 0; j < parts.length; j++) {
                    const part = parts[j];

                    if (!part || typeof part.text !== "string" ||
                        !part.text.trim()) {
                        continue;
                    }

                    totalTextParts++;

                    // لا نعيد أجزاء التفكير الداخلي إلى المستخدم أو إلى محلل JSON.
                    if (part.thought === true) {
                        thoughtTextParts++;
                        continue;
                    }

                    finalTexts.push(part.text);
                }

                candidateStates.push({
                    finishReason: candidate.finishReason || "غير محدد",
                    parts: parts.length
                });

                const text = finalTexts.join("\n").trim();
                if (text) {
                    return text;
                }
            }

            const states = candidateStates.map(function (item, index) {
                return "مرشح " + (index + 1) +
                    ": finishReason=" + item.finishReason +
                    ", parts=" + item.parts;
            }).join(" | ");

            const reasonText = blockReason
                ? "؛ promptFeedback=" + blockReason
                : "";

            throw new Error(
                "استجابة Gemini لا تحتوي على نص نهائي قابل للاستخدام" +
                reasonText +
                ". textParts=" + totalTextParts +
                ", thoughtParts=" + thoughtTextParts +
                ". " + states
            );
        }

        // الإبقاء على سلوك مزودي OpenAI-compatible كما هو.
        const text = result &&
            result.choices &&
            result.choices[0] &&
            result.choices[0].message &&
            result.choices[0].message.content;

        if (typeof text === "string" && text.trim()) {
            return text.trim();
        }

        throw new Error("لم يصل رد صالح من " + provider + ".");
    };

    console.log("تم تحميل معالج استجابات Gemini المحسّن.");
})(window);
