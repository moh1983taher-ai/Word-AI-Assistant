(function (global) {
    "use strict";

    const TRACKED_CHANGE_TAG = /<(?:[A-Za-z_][\w.-]*:)?(?:ins|del|moveFrom|moveTo)(?:\s|\/?>)/i;
    const PROTECTED_TOKEN_PATTERN = /\[\[WB_KEEP_[A-Za-z0-9_-]+\]\]/g;

    function hasTrackedChanges(ooxml) {
        return TRACKED_CHANGE_TAG.test(String(ooxml || ""));
    }

    function extractParagraphXml(ooxml) {
        const source = String(ooxml || "");
        const pattern = /<((?:[A-Za-z_][\w.-]*:)?p)(?:\s[^>]*)?>[\s\S]*?<\/\1\s*>/g;
        const paragraphs = [];
        let match;
        while ((match = pattern.exec(source)) !== null) {
            paragraphs.push(match[0]);
        }
        return paragraphs;
    }

    function stripParagraphTerminator(value) {
        return String(value || "").replace(/(?:\r\n|\r|\n)+$/, "");
    }

    function mapParagraphRanges(fullText, paragraphTexts) {
        const source = String(fullText || "");
        const records = Array.isArray(paragraphTexts) ? paragraphTexts : [];
        const mappings = [];
        let cursor = 0;

        records.forEach(function (record, index) {
            const rawText = record && typeof record === "object" && "text" in record
                ? record.text
                : record;
            const paragraphText = stripParagraphTerminator(rawText);
            if (!paragraphText) return;

            const start = source.indexOf(paragraphText, cursor);
            if (start < 0) {
                throw new Error("تعذر مطابقة بنية فقرات Word مع النص المستخرج؛ أُوقف التدقيق لمنع تعديل مواضع غير صحيحة.");
            }
            const end = start + paragraphText.length;
            mappings.push({
                paragraphIndex: Number.isInteger(record && record.index) ? record.index : index,
                start: start,
                end: end,
                text: paragraphText
            });
            cursor = end;
        });

        return mappings;
    }

    function findParagraphRange(paragraphRanges, startValue, lengthValue) {
        const ranges = Array.isArray(paragraphRanges) ? paragraphRanges : [];
        const start = Number(startValue);
        const length = Math.max(0, Number(lengthValue) || 0);
        if (!Number.isInteger(start) || start < 0) return null;

        const matches = ranges.filter(function (range) {
            return start >= range.start && start + length <= range.end;
        });
        return matches.length === 1 ? matches[0] : null;
    }

    function buildProtectedParagraphRanges(fullText, paragraphRecords) {
        const records = Array.isArray(paragraphRecords) ? paragraphRecords : [];
        const protectedIndices = new Set();

        records.forEach(function (record, index) {
            const hasApiChanges = Number(record && record.trackedChangesCount) > 0 ||
                record && record.hasTrackedChanges === true;
            if (hasApiChanges || hasTrackedChanges(record && record.ooxml)) {
                protectedIndices.add(Number.isInteger(record.index) ? record.index : index);
            }
        });

        if (!protectedIndices.size) return [];

        records.forEach(function (record, index) {
            const paragraphText = stripParagraphTerminator(record && record.text);
            const paragraphIndex = Number.isInteger(record && record.index) ? record.index : index;
            if (protectedIndices.has(paragraphIndex) && !paragraphText) {
                throw new Error("تعذر تحديد فقرة تحتوي على تغييرات متعقبة؛ أُوقف التدقيق حمايةً للمستند.");
            }
        });

        const mappings = mapParagraphRanges(fullText, records);
        const ranges = mappings.filter(function (mapping) {
            return protectedIndices.has(mapping.paragraphIndex);
        });

        if (ranges.length !== protectedIndices.size) {
            throw new Error("تعذر تحديد جميع الفقرات ذات التغييرات المتعقبة؛ أُوقف التدقيق حمايةً للمستند.");
        }
        return ranges;
    }

    function maskChunk(sourceValue, protectedRanges, baseOffset, revisionId) {
        const source = String(sourceValue || "");
        const ranges = Array.isArray(protectedRanges) ? protectedRanges.slice() : [];
        const offset = Math.max(0, Number(baseOffset) || 0);
        const chunkEnd = offset + source.length;
        const overlapping = ranges
            .filter(function (range) {
                return Number(range.start) < chunkEnd && Number(range.end) > offset;
            })
            .sort(function (a, b) {
                return Number(a.start) - Number(b.start);
            });

        if (!overlapping.length) {
            return { text: source, source: source, offset: offset, segments: [] };
        }

        const id = String(revisionId || "revision").replace(/[^A-Za-z0-9_-]/g, "_");
        const segments = [];
        let output = "";
        let cursor = 0;

        overlapping.forEach(function (range, index) {
            const absoluteStart = Number(range.start);
            const absoluteEnd = Number(range.end);
            if (!Number.isInteger(absoluteStart) || !Number.isInteger(absoluteEnd) ||
                absoluteStart < offset || absoluteEnd > chunkEnd || absoluteEnd <= absoluteStart) {
                throw new Error("تقاطع حدّ الدفعة مع فقرة ذات تغييرات متعقبة؛ أُوقف التدقيق حتى لا ينكشف جزء من النص غير المستقر.");
            }

            const localStart = absoluteStart - offset;
            const localEnd = absoluteEnd - offset;
            const protectedText = String(range.text || "");
            if (source.slice(localStart, localEnd) !== protectedText) {
                throw new Error("تغيرت فقرة محمية بعد قراءتها؛ أُوقف التدقيق حمايةً للنص.");
            }
            if (localStart < cursor) {
                throw new Error("تداخلت حدود الفقرات المحمية؛ أُوقف التدقيق لتجنب إخفاء نطاق غير صحيح.");
            }

            let token = "[[WB_KEEP_" + id + "_" +
                String(Number.isInteger(range.paragraphIndex) ? range.paragraphIndex : index) + "_" + index + "]]";
            while (source.includes(token)) token = token.slice(0, -2) + "_X]]";

            output += source.slice(cursor, localStart) + token;
            segments.push({
                token: token,
                paragraphIndex: range.paragraphIndex,
                start: localStart,
                end: localEnd,
                text: protectedText
            });
            cursor = localEnd;
        });

        output += source.slice(cursor);
        return { text: output, source: source, offset: offset, segments: segments };
    }

    function mapMaskedOffsetToSource(maskState, maskedOffset) {
        if (!maskState || !Array.isArray(maskState.segments) || !maskState.segments.length) {
            return Number(maskedOffset);
        }

        const target = Number(maskedOffset);
        if (!Number.isInteger(target) || target < 0 || target > maskState.text.length) {
            throw new Error("تعذر تحويل موضع التصحيح من النص المحمي إلى موضعه الأصلي.");
        }

        let sourceCursor = 0;
        let maskedCursor = 0;

        for (const segment of maskState.segments) {
            const sourceLength = segment.start - sourceCursor;
            const maskedEnd = maskedCursor + sourceLength;

            if (target <= maskedEnd) {
                return sourceCursor + (target - maskedCursor);
            }

            maskedCursor = maskedEnd + segment.token.length;
            sourceCursor = segment.end;
        }

        return sourceCursor + (target - maskedCursor);
    }

    function countOccurrences(source, needle) {
        if (!needle) return 0;
        let count = 0;
        let cursor = 0;
        while ((cursor = source.indexOf(needle, cursor)) !== -1) {
            count++;
            cursor += needle.length;
        }
        return count;
    }

    function restoreChunk(value, maskState) {
        const output = String(value || "");
        if (!maskState || !Array.isArray(maskState.segments) || !maskState.segments.length) {
            return output;
        }

        const positions = [];
        maskState.segments.forEach(function (segment) {
            if (countOccurrences(output, segment.token) !== 1) {
                throw new Error("لم يحافظ النموذج على علامة حماية إحدى الفقرات ذات التغييرات المتعقبة؛ أُلغي تجهيز النتيجة.");
            }
            const position = output.indexOf(segment.token);
            const originalPosition = maskState.text.indexOf(segment.token);
            const lineBreakCount = function (value) {
                return (String(value || "").match(/\r\n|\r|\n/g) || []).length;
            };
            if (lineBreakCount(output.slice(0, position)) !==
                lineBreakCount(maskState.text.slice(0, originalPosition))) {
                throw new Error("نقل النموذج فقرة محمية إلى موضع آخر؛ أُلغي تجهيز النتيجة.");
            }
            positions.push(position);
        });

        for (let index = 1; index < positions.length; index++) {
            if (positions[index] <= positions[index - 1]) {
                throw new Error("غيّر النموذج ترتيب الفقرات المحمية؛ أُلغي تجهيز النتيجة.");
            }
        }

        const expectedTokens = new Set(maskState.segments.map(function (segment) {
            return segment.token;
        }));
        const foundTokens = output.match(PROTECTED_TOKEN_PATTERN) || [];
        if (foundTokens.length !== expectedTokens.size ||
            foundTokens.some(function (token) { return !expectedTokens.has(token); })) {
            throw new Error("غيّر النموذج علامة حماية أو أضاف علامة غير معروفة؛ أُلغي تجهيز النتيجة.");
        }

        let restored = output;
        maskState.segments.forEach(function (segment) {
            restored = restored.replace(segment.token, segment.text);
        });
        return restored;
    }

    function assertNoProtectedEdits(edits, maskState) {
        if (!maskState || !Array.isArray(maskState.segments) || !maskState.segments.length) return;

        (Array.isArray(edits) ? edits : []).forEach(function (edit) {
            const start = Number(edit && edit.start);
            const end = start + String(edit && edit.original || "").length;
            const overlaps = maskState.segments.some(function (segment) {
                return start < segment.end && end > segment.start;
            });
            if (overlaps) {
                throw new Error("اقترح النموذج تعديلًا داخل فقرة لها تغييرات متعقبة؛ رُفض التصحيح حمايةً للمراجعات السابقة.");
            }
        });
    }

    const api = {
        hasTrackedChanges: hasTrackedChanges,
        extractParagraphXml: extractParagraphXml,
        mapParagraphRanges: mapParagraphRanges,
        findParagraphRange: findParagraphRange,
        mapMaskedOffsetToSource: mapMaskedOffsetToSource,
        buildProtectedParagraphRanges: buildProtectedParagraphRanges,
        maskChunk: maskChunk,
        restoreChunk: restoreChunk,
        assertNoProtectedEdits: assertNoProtectedEdits
    };

    global.DocumentRevisionPreflight = api;
    if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
