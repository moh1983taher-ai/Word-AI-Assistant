(function (global) {
    "use strict";

    const DIGIT = "[0-9٠-٩۰-۹]";
    const NUMBER = DIGIT + "{1,6}(?:\\s*\\/\\s*" + DIGIT + "{1,6})?(?:\\s*[-–—]\\s*" + DIGIT + "{1,6})?";
    const LOCATION_LABEL = "(?:(?:ج(?:زء)?|المجلد(?:ات)?|ص(?:فحة|فحات)?|صفحات?)\\s*)?";
    const LOCATION_SEPARATOR = "(?:\\s*[،,؛;]\\s*|\\s+)";
    const LOCATION_SUFFIX = new RegExp(
        LOCATION_SEPARATOR + LOCATION_LABEL + NUMBER +
        "(?:" + LOCATION_SEPARATOR + LOCATION_LABEL + NUMBER + ")*\\s*$",
        "u"
    );

    function normalizeIdentityText(value) {
        return String(value ?? "")
            .normalize("NFC")
            .toLowerCase()
            .replace(/[\u064B-\u065F\u0670\u0640]/gu, "")
            .replace(/[\u200B-\u200F\u202A-\u202E\uFEFF]/gu, "")
            .replace(/[أإآٱ]/gu, "ا")
            .replace(/ى/gu, "ي")
            .replace(/[ؤ]/gu, "و")
            .replace(/[ئ]/gu, "ي")
            .replace(/ة/gu, "ه")
            .replace(/[\s\u00A0\p{P}\p{S}]+/gu, "");
    }

    function referenceKey(reference) {
        if (!reference || typeof reference !== "object") return "";
        const author = normalizeIdentityText(reference.author);
        const title = normalizeIdentityText(getIdentityText(reference.title));
        if (!title) return "";
        return author + "::" + title;
    }

    function uniqueReferenceEntries(entries) {
        const seen = new Set();
        const result = [];
        for (const item of Array.isArray(entries) ? entries : []) {
            const key = referenceKey(item && item.reference);
            if (key && seen.has(key)) continue;
            if (key) seen.add(key);
            result.push(item);
        }
        return result;
    }

    function getIdentityText(value) {
        const text = String(value ?? "").trim();
        if (!text) return "";
        const match = LOCATION_SUFFIX.exec(text);
        if (!match || match.index <= 0) return text;
        return text.slice(0, match.index)
            .trim()
            .replace(/[،,؛;:\s]+$/gu, "")
            .trim();
    }

    function hasUnparsedLocationSuffix(value) {
        const text = String(value ?? "").trim();
        if (!text) return false;
        return /(?:^|[،,؛;\s])(?:ج(?:زء)?|المجلد(?:ات)?|ص(?:فحة|فحات)?|صفحات?)\s*[0-9٠-٩۰-۹]/u.test(text) ||
            /(?:^|[،,؛;\s])[0-9٠-٩۰-۹]{1,6}\s*\/\s*[0-9٠-٩۰-۹]{1,6}(?:\s*[،,؛;]\s*[0-9٠-٩۰-۹]{1,6}\s*\/\s*[0-9٠-٩۰-۹]{1,6})?\s*$/u.test(text) ||
            /(?:^|[،,؛;\s])[0-9٠-٩۰-۹]{1,6}\s*[-–—]\s*[0-9٠-٩۰-۹]{1,6}\s*$/u.test(text);
    }

    function findExactOccurrences(text, needle) {
        const source = String(text ?? "");
        const target = String(needle ?? "");
        if (!target) return [];
        const positions = [];
        let from = 0;
        while (from <= source.length - target.length) {
            const index = source.indexOf(target, from);
            if (index < 0) break;
            positions.push(index);
            from = index + Math.max(1, target.length);
        }
        return positions;
    }

        function findValidatedSourceSpan(noteText, candidateText) {
        const source = String(noteText ?? "");
        const candidate = String(candidateText ?? "").trim();
        if (!source || !candidate) return "";
        const target = normalizeIdentityText(candidate);
        if (target.length < 4) return "";

        const normalized = [];
        const rawMap = [];
        for (let index = 0; index < source.length; index++) {
            const part = normalizeIdentityText(source[index]);
            if (!part) continue;
            normalized.push(part);
            for (let offset = 0; offset < part.length; offset++) rawMap.push(index);
        }

        const haystack = normalized.join("");
        let from = 0;
        while (from <= haystack.length - target.length) {
            const position = haystack.indexOf(target, from);
            if (position < 0) break;
            const rawStart = rawMap[position];
            const rawEnd = rawMap[position + target.length - 1] + 1;
            from = position + Math.max(1, target.length);
            if (!Number.isInteger(rawStart) || !Number.isInteger(rawEnd) || rawEnd <= rawStart) continue;

            const previous = rawStart > 0 ? source[rawStart - 1] : "";
            const next = source.slice(rawEnd).match(/^\\s*/u)?.[0].length || 0;
            const following = source[rawEnd + next] || "";
            if (previous && /[\\p{L}\\p{N}]/u.test(previous)) continue;
            if (following && /[\\p{L}\\p{N}]/u.test(following)) continue;
            return source.slice(rawStart, rawEnd);
        }
        return "";
    }

function getNoteIndex(suggestion, source) {
        const materialId = String(suggestion && suggestion.materialId || "");
        const pattern = source === "footnote" ? /^footnote-(\d+)$/ : /^endnote-(\d+)$/;
        const match = materialId.match(pattern);
        if (match) {
            const value = Number(match[1]);
            return Number.isInteger(value) && value > 0 ? value - 1 : -1;
        }
        const noteNumber = Number(suggestion && suggestion.noteNumber);
        return Number.isInteger(noteNumber) && noteNumber > 0 ? noteNumber - 1 : -1;
    }

    function normalizeBibliographicTitle(value, author, locationText = "") {
        let title = String(value ?? "").trim();
        const location = String(locationText ?? "").trim();
        if (title && location) {
            const suffixCandidates = [location];
            if (/^ص/u.test(location)) {
                const pageValue = location.replace(/^ص\s*/u, "").trim();
                if (pageValue) {
                    suffixCandidates.push("ص" + pageValue);
                    suffixCandidates.push("صفحة " + pageValue);
                }
            }
            for (const suffix of new Set(suffixCandidates)) {
                const position = title.lastIndexOf(suffix);
                if (position < 0) continue;
                const tail = title.slice(position + suffix.length);
                const hasBoundary = position === 0 || /[\s،,؛:]/u.test(title[position - 1]);
                if (tail.trim() === "" && hasBoundary) {
                    title = title.slice(0, position).replace(/[،,؛:\s]+$/u, "").trim();
                    break;
                }
            }
        }
        const authorKey = normalizeIdentityText(author);
        if (!title) return "";

        if (authorKey) {
            const parts = title.split(/[،,؛;]+/u).map(part => part.trim()).filter(Boolean);
            const authorFiltered = parts.filter(part => normalizeIdentityText(part) !== authorKey);
            if (authorFiltered.length && authorFiltered.length !== parts.length) {
                title = authorFiltered.join("، ");
            }
        }

        const tokens = title.split(/\s+/u).filter(Boolean);
        const cleaned = [];

        function tokenInfo(token) {
            const raw = String(token || "");
            const core = raw.replace(/^[،,؛:()[\]{}"'\u0060]+/u, "").replace(/[،,؛:()[\]{}"'\u0060]+$/u, "");
            const hasWaw = /^و(?=.{2,})/u.test(core);
            return {
                key: normalizeIdentityText(hasWaw ? core.slice(1) : core),
                hasWaw: hasWaw,
                suffix: (raw.match(/[،,؛:.)\]}]+$/u) || [""])[0]
            };
        }

        for (const token of tokens) {
            const current = tokenInfo(token);
            const previous = cleaned.length ? tokenInfo(cleaned[cleaned.length - 1]) : null;
            const duplicate = previous && current.key &&
                current.key === previous.key &&
                (current.hasWaw || previous.hasWaw);

            if (duplicate) {
                if (current.suffix && !/[،,؛:.)\]}]+$/u.test(cleaned[cleaned.length - 1])) {
                    cleaned[cleaned.length - 1] += current.suffix;
                }
                continue;
            }
            cleaned.push(token);
        }

        return cleaned.join(" ").replace(/\s+([،,؛:.)])/gu, "$1").trim();
    }

    function buildPairs(noteText, suggestion) {
        const references = Array.isArray(suggestion && suggestion.references) ? suggestion.references : [];
        const suggestedTexts = Array.isArray(suggestion && suggestion.suggestedTexts) ? suggestion.suggestedTexts : [];
        const sourceTexts = Array.isArray(suggestion && suggestion.sourceTexts) ? suggestion.sourceTexts : [];
        const count = Math.min(references.length, suggestedTexts.length, sourceTexts.length);
        const pairs = [];
        pairs.collapsedCount = 0;
        pairs.unchangedCount = 0;

        // The same work can legitimately occur more than once in a single note.
        // Preserve each occurrence using the ordinal of its complete source anchor.
        const sourceOccurrenceCounts = new Map();

        for (let index = 0; index < count; index++) {
            const reference = references[index];
            const sourceText = String(sourceTexts[index] ?? "").trim();
            const replacementText = getIdentityText(suggestedTexts[index]);
            const identityText = getIdentityText(sourceText);
            if (!sourceText || !identityText || !replacementText) continue;

            const sourceKey = normalizeIdentityText(sourceText);
            if (!sourceKey) continue;

            const sourceOccurrenceIndex = sourceOccurrenceCounts.get(sourceKey) || 0;
            const sourceMatches = findExactOccurrences(noteText, sourceText);
            if (sourceMatches.length <= sourceOccurrenceIndex) continue;

            const sourceStart = sourceMatches[sourceOccurrenceIndex];
            sourceOccurrenceCounts.set(sourceKey, sourceOccurrenceIndex + 1);

            const identityOffset = sourceText.indexOf(identityText);
            if (identityOffset < 0) continue;

            const start = sourceStart + identityOffset;
            const end = start + identityText.length;
            if (noteText.slice(start, end) !== identityText) continue;

            if (normalizeIdentityText(identityText) === normalizeIdentityText(replacementText)) {
                pairs.unchangedCount++;
                continue;
            }

            pairs.push({
                sourceText: sourceText,
                identityText: identityText,
                replacementText: replacementText,
                start: start,
                end: end,
                expectedText: identityText,
                sourceOccurrenceIndex: sourceOccurrenceIndex,
                referenceKey: referenceKey(reference) || sourceKey
            });
        }

        pairs.sort((a, b) => a.start - b.start || b.end - a.end);
        const nonOverlapping = [];
        for (const pair of pairs) {
            const previous = nonOverlapping[nonOverlapping.length - 1];
            if (previous && pair.start < previous.end) continue;
            nonOverlapping.push(pair);
        }
        nonOverlapping.collapsedCount = pairs.collapsedCount;
        nonOverlapping.unchangedCount = pairs.unchangedCount;
        return nonOverlapping;
    }

    function selectOccurrence(matches, occurrenceIndex) {
        if (!Array.isArray(matches) || !Number.isInteger(occurrenceIndex) || occurrenceIndex < 0) {
            return null;
        }
        return matches[occurrenceIndex] || null;
    }

    function findUniqueBibliographicAnchor(noteText, reference) {
        const text = String(noteText ?? "");
        if (!text.trim() || !reference || typeof reference !== "object") return "";

        const author = String(reference.author ?? "").trim();
        const title = normalizeBibliographicTitle(getIdentityText(reference.title ?? ""), author);
        if (!title) return "";

        const candidates = [];
        const addCandidate = value => {
            const candidate = String(value ?? "").trim();
            const key = normalizeIdentityText(candidate);
            if (key.length < 8 || candidates.some(item => normalizeIdentityText(item) === key)) return;
            candidates.push(candidate);
        };

        if (author) {
            addCandidate(author + "، " + title);
            addCandidate(title + "، " + author);
            addCandidate(author + " " + title);
            addCandidate(title + " " + author);
        }
        addCandidate(title);

        let normalized = "";
        const rawMap = [];
        for (let index = 0; index < text.length; index++) {
            const part = normalizeIdentityText(text[index]);
            normalized += part;
            for (let offset = 0; offset < part.length; offset++) rawMap.push(index);
        }

        candidates.sort((a, b) => normalizeIdentityText(b).length - normalizeIdentityText(a).length);

        for (const candidate of candidates) {
            const target = normalizeIdentityText(candidate);
            const positions = [];
            let from = 0;
            while (from <= normalized.length - target.length) {
                const position = normalized.indexOf(target, from);
                if (position < 0) break;
                positions.push(position);
                from = position + Math.max(1, target.length);
            }

            // A fallback is safe only when the bibliographic identity occurs once.
            // Repeated occurrences must be anchored by an explicit source variant.
            if (positions.length !== 1) continue;

            const position = positions[0];
            const rawStart = rawMap[position];
            const rawEnd = rawMap[position + target.length - 1] + 1;
            if (Number.isInteger(rawStart) && Number.isInteger(rawEnd) && rawEnd > rawStart) {
                const validated = findValidatedSourceSpan(text, text.slice(rawStart, rawEnd));
                if (validated) return validated;
            }
        }

        return "";
    }

    function searchOptions() {
        return {
            matchCase: false,
            matchWholeWord: false,
            matchWildcards: false,
            ignorePunct: true,
            ignoreSpace: true
        };
    }

    async function apply(suggestions) {
        const items = Array.isArray(suggestions) ? suggestions : [];
        if (!items.length) return { applied: 0, skipped: 0, unchanged: 0 };

        const summaries = { applied: 0, skipped: 0, unchanged: 0 };

        await Word.run(async function (context) {
            const body = context.document.body;
            const footnotes = body.footnotes;
            const endnotes = body.endnotes;
            footnotes.load("items/body");
            endnotes.load("items/body");
            await context.sync();

            const noteEntries = [];
            function addNotes(collection, source) {
                collection.items.forEach(function (note, index) {
                    const range = note.body.getRange();
                    range.load("text");
                    noteEntries.push({
                        source: source,
                        index: index,
                        range: range,
                        text: ""
                    });
                });
            }
            addNotes(footnotes, "footnote");
            addNotes(endnotes, "endnote");
            if (noteEntries.length) await context.sync();

            noteEntries.forEach(function (entry) {
                entry.text = String(entry.range.text ?? "").trim();
            });

            const noteByKey = new Map(noteEntries.map(entry => [
                entry.source + ":" + entry.index,
                entry
            ]));

            const operations = [];
            const options = searchOptions();

            for (const suggestion of items) {
                const source = suggestion && suggestion.source;
                if (source !== "footnote" && source !== "endnote") {
                    summaries.skipped++;
                    continue;
                }

                const noteIndex = getNoteIndex(suggestion, source);
                const entry = noteByKey.get(source + ":" + noteIndex);
                if (!entry || normalizeIdentityText(entry.text) !==
                    normalizeIdentityText(suggestion.originalText)) {
                    summaries.skipped += Math.max(
                        1,
                        Math.min(
                            Array.isArray(suggestion && suggestion.references) ? suggestion.references.length : 0,
                            Array.isArray(suggestion && suggestion.sourceTexts) ? suggestion.sourceTexts.length : 0,
                            Array.isArray(suggestion && suggestion.suggestedTexts) ? suggestion.suggestedTexts.length : 0
                        )
                    );
                    continue;
                }

                const pairs = buildPairs(entry.text, suggestion);
                const requestedCount = Math.min(
                    Array.isArray(suggestion && suggestion.references) ? suggestion.references.length : 0,
                    Array.isArray(suggestion && suggestion.sourceTexts) ? suggestion.sourceTexts.length : 0,
                    Array.isArray(suggestion && suggestion.suggestedTexts) ? suggestion.suggestedTexts.length : 0
                );
                summaries.skipped += Math.max(0, requestedCount - pairs.length -
                    (pairs.collapsedCount || 0) - (pairs.unchangedCount || 0));
                summaries.unchanged += pairs.unchangedCount || 0;

                for (const pair of pairs) {
                    operations.push({ entry: entry, pair: pair });
                }
            }

            if (!operations.length) return;

            // Bound the number of queued Word searches. Each batch is fully
            // resolved before the next one, while all mutations remain deferred.
            const SEARCH_BATCH_SIZE = 60;
            const resolved = [];

            for (let offset = 0; offset < operations.length; offset += SEARCH_BATCH_SIZE) {
                const batch = operations.slice(offset, offset + SEARCH_BATCH_SIZE);
                const fullSearches = batch.map(function (operation) {
                    return {
                        entry: operation.entry,
                        pair: operation.pair,
                        fullSearch: operation.entry.range.search(
                            operation.pair.sourceText,
                            options
                        ).load("items")
                    };
                });

                await context.sync();

                const fullCandidates = [];
                for (const operation of fullSearches) {
                    for (const match of operation.fullSearch.items) {
                        match.load("text");
                        fullCandidates.push(match);
                    }
                }
                if (fullCandidates.length) await context.sync();

                const identitySearches = [];
                for (const operation of fullSearches) {
                    const expectedFull = normalizeIdentityText(operation.pair.sourceText);
                    const matches = operation.fullSearch.items.filter(match =>
                        normalizeIdentityText(match.text) === expectedFull
                    );
                    const fullMatch = selectOccurrence(matches, operation.pair.sourceOccurrenceIndex);
                    if (!fullMatch) {
                        summaries.skipped++;
                        continue;
                    }

                    const identitySearch = fullMatch.search(
                        operation.pair.identityText,
                        options
                    ).load("items");

                    identitySearches.push({
                        entry: operation.entry,
                        pair: operation.pair,
                        identitySearch: identitySearch
                    });
                }

                if (identitySearches.length) await context.sync();

                const identityCandidates = [];
                for (const operation of identitySearches) {
                    for (const match of operation.identitySearch.items) {
                        match.load("text");
                        identityCandidates.push(match);
                    }
                }
                if (identityCandidates.length) await context.sync();

                for (const operation of identitySearches) {
                    const expectedIdentity = normalizeIdentityText(operation.pair.identityText);
                    const matches = operation.identitySearch.items.filter(match =>
                        normalizeIdentityText(match.text) === expectedIdentity
                    );
                    if (matches.length !== 1) {
                        summaries.skipped++;
                        continue;
                    }

                    const range = matches[0];
                    if (String(range.text ?? "") !== operation.pair.expectedText) {
                        summaries.skipped++;
                        continue;
                    }

                    resolved.push({
                        entry: operation.entry,
                        pair: operation.pair,
                        range: range,
                        insertedRange: null
                    });
                }
            }

            // Re-read every affected note immediately before queuing mutations.
            const affectedEntries = Array.from(new Set(resolved.map(item => item.entry)));
            for (const entry of affectedEntries) entry.range.load("text");
            if (affectedEntries.length) await context.sync();

            for (const entry of affectedEntries) {
                if (String(entry.range.text ?? "").trim() !== entry.text) {
                    throw new Error("تغير نص إحدى الحواشي بعد إعداد المعاينة؛ أُلغي الاستبدال حمايةً للمستند. أعد التحليل ثم حاول مجددًا.");
                }
            }

            // Sort from the end of each note to the beginning; distinct ranges remain stable.
            resolved.sort((a, b) =>
                a.entry.source.localeCompare(b.entry.source) ||
                a.entry.index - b.entry.index ||
                b.pair.start - a.pair.start
            );

            for (const item of resolved) {
                item.insertedRange = item.range.insertText(
                    item.pair.replacementText,
                    "Replace"
                );
                item.insertedRange.load("text");
            }

            if (resolved.length) await context.sync();

            for (const item of resolved) {
                if (normalizeIdentityText(item.insertedRange.text) ===
                    normalizeIdentityText(item.pair.replacementText)) {
                    summaries.applied++;
                } else {
                    summaries.skipped++;
                }
            }
        });

        return {
            applied: summaries.applied,
            skipped: summaries.skipped,
            unchanged: summaries.unchanged
        };
    }

    global.FootnoteCitationEngine = {
        apply: apply,
        buildPairs: buildPairs,
        findExactOccurrences: findExactOccurrences,
        findValidatedSourceSpan: findValidatedSourceSpan,
        getIdentityText: getIdentityText,
        hasUnparsedLocationSuffix: hasUnparsedLocationSuffix,
        normalizeIdentityText: normalizeIdentityText,
        referenceKey: referenceKey,
        uniqueReferenceEntries: uniqueReferenceEntries,
        normalizeBibliographicTitle: normalizeBibliographicTitle,
        findUniqueBibliographicAnchor: findUniqueBibliographicAnchor,
        selectOccurrence: selectOccurrence
    };

    if (typeof module !== "undefined" && module.exports) {
        module.exports = global.FootnoteCitationEngine;
    }
})(typeof window !== "undefined" ? window : globalThis);
