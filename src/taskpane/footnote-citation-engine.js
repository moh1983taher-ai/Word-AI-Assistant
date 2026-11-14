(function (global) {
    "use strict";

    const DIGIT = "[0-9٠-٩۰-۹]";
    const NUMBER = DIGIT + "{1,6}(?:\\s*\\/\\s*" + DIGIT + "{1,6})?(?:\\s*[-–—]\\s*" + DIGIT + "{1,6})?";
    const LOCATION_LABEL = "(?:(?:ج(?:زء)?|المجلد(?:ات)?|ص(?:فحة|فحات)?|صفحات?)\\s*)?";
    const LOCATION_SEPARATOR = "(?:\\s*[،,؛;]\\s*|\\s+)";
    const LOCATION_SUFFIX = new RegExp(
        LOCATION_SEPARATOR + LOCATION_LABEL + NUMBER +
        "(?:" + LOCATION_SEPARATOR + LOCATION_LABEL + NUMBER + ")*" +
        "(?:\\s*(?:[،,؛;]\\s*)?(?:وما بعدها|وما بعده|وما يليها|وما يليه|وما بعد ذلك))?\\s*$",
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
        let text = String(value ?? "").trim();
        if (!text) return "";

        // Outer punctuation enclosing the selected citation is not part of
        // the bibliographic identity; it must remain in the surrounding note.
        const wrappers = [
            ["(", ")"], ["[", "]"], ["（", "）"], ["［", "］"],
            ["{", "}"], ["【", "】"], ["«", "»"], ["“", "”"]
        ];
        for (const [opening, closing] of wrappers) {
            if (!text.startsWith(opening) || !text.endsWith(closing)) continue;
            let depth = 0;
            let closesAtEnd = true;
            for (let index = 0; index < text.length; index++) {
                if (text[index] === opening) depth++;
                if (text[index] === closing) depth--;
                if (depth === 0 && index < text.length - 1) {
                    closesAtEnd = false;
                    break;
                }
            }
            if (closesAtEnd && depth === 0) {
                text = text.slice(opening.length, -closing.length).trim();
                break;
            }
        }

        // Keep note counters outside the replacement, whether or not an
        // editorial lead-in such as "انظر:" follows the counter.
        text = text.replace(
            /^(?:(?:\(\s*[0-9٠-٩۰-۹]{1,4}\s*\)|\[\s*[0-9٠-٩۰-۹]{1,4}\s*\]|[0-9٠-٩۰-۹]{1,4}[.)])\s*)+/u,
            ""
        ).trimStart();
        text = text.replace(
            /^[0-9٠-٩۰-۹]{1,4}(?=\s+(?:انظر|ينظر|راجع|يراجع|قارن|نقلاً عن|نقلًا عن)(?:\s|[:：]|$))\s+/u,
            ""
        );
        text = text.replace(
            /^(?:(?:انظر|ينظر|راجع|يراجع|قارن|نقلاً عن|نقلًا عن)\s*[:：]?\s*)+/u,
            ""
        ).trim();

        // Terminal punctuation and an unmatched closing bracket delimit the
        // citation; they are not characters to be consumed by replacement.
        text = text.replace(/[،,؛;:.!?؟۔]+$/u, "").trim();
        const closingPairs = {
            ")": "(", "]": "[", "}": "{", "）": "（", "］": "［", "】": "【"
        };
        for (const [closing, opening] of Object.entries(closingPairs)) {
            if (!text.endsWith(closing)) continue;
            const openingCount = [...text].filter(char => char === opening).length;
            const closingCount = [...text].filter(char => char === closing).length;
            if (closingCount > openingCount) text = text.slice(0, -closing.length).trimEnd();
        }

        const match = LOCATION_SUFFIX.exec(text);
        if (!match || match.index <= 0) return text;
        return text.slice(0, match.index)
            .trim()
            .replace(/[،,؛;:.\s]+$/gu, "")
            .trim();
    }

    function isHadithTakhrijReference(reference, occurrence = null) {
        if (!reference || typeof reference !== "object") return false;
        const type = String(reference.type || reference.referenceType || reference.kind || "").trim().toLowerCase();
        if (["hadith", "حديث", "hadith-takhrij", "takhrij"].includes(type)) return true;

        const variantTexts = (Array.isArray(reference.variants) ? reference.variants : []).map(item => {
            if (typeof item === "string") return item;
            if (!item || typeof item !== "object") return "";
            return String(item.text || item.originalText || item.cleanedText || item.citation || item.value || "");
        });

        const evidence = [
            reference.author, reference.title, reference.originalText,
            reference.cleanedText, reference.notes,
            occurrence && occurrence.sourceText,
            occurrence && occurrence.originalText,
            ...variantTexts
        ].filter(Boolean).join(" ");

        return /(?:^|[\s،,؛:])(?:أخرجه|أخرج|رواه|روى|خرَّجه|خرّجه|خرجه|تخريج(?:\s+الحديث)?|عزاه)(?=$|[\s،,؛:])|حديث\s*(?:رقم|برقم)|رقم\s+الحديث|متفق\s+عليه/u.test(evidence);
    }

    function hasUnparsedLocationSuffix(value) {
        const text = String(value ?? "").trim();
        if (!text) return false;
        return /(?:^|[،,؛;\s])(?:ج(?:زء)?|المجلد(?:ات)?|ص(?:فحة|فحات)?|صفحات?)\s*[0-9٠-٩۰-۹]/u.test(text) ||
            /(?:^|[،,؛;\s])[0-9٠-٩۰-۹]{1,6}\s*\/\s*[0-9٠-٩۰-۹]{1,6}(?:\s*[،,؛;]\s*[0-9٠-٩۰-۹]{1,6}\s*\/\s*[0-9٠-٩۰-۹]{1,6})?\s*$/u.test(text) ||
            /(?:^|[،,؛;\s])[0-9٠-٩۰-۹]{1,6}\s*[-–—]\s*[0-9٠-٩۰-۹]{1,6}(?:\s*[،,؛;]?\s*وما بعدها)?\s*$/u.test(text);
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
            const next = source.slice(rawEnd).match(/^\s*/u)?.[0].length || 0;
            const following = source[rawEnd + next] || "";
            if (previous && /[\p{L}\p{N}]/u.test(previous)) {
                const beforePrevious = rawStart > 1 ? source[rawStart - 2] : "";
                const isAttachedKaf = previous === "ك" &&
                    /^ال/u.test(source.slice(rawStart, rawStart + 2)) &&
                    (!beforePrevious || !/[\p{L}\p{N}]/u.test(beforePrevious));
                if (!isAttachedKaf) continue;
            }
            if (following && /[\p{L}\p{N}]/u.test(following)) continue;
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


    function findNormalizedTextSpans(text, value) {
        const source = String(text ?? "");
        const rawValue = String(value ?? "").trim();
        const primaryTarget = normalizeIdentityText(rawValue);
        if (!source || primaryTarget.length < 2) return [];

        // Some Arabic prepositions assimilate the definite article in writing:
        // "للأزهري" is a valid surface form of "الأزهري". Try that spelling
        // variant only when the target itself begins with the definite article.
        const targets = [primaryTarget];
        if (primaryTarget.startsWith("ال") && primaryTarget.length > 3) {
            targets.push(primaryTarget.slice(1));
        }

        let normalized = "";
        const rawMap = [];
        for (let index = 0; index < source.length; index++) {
            const part = normalizeIdentityText(source[index]);
            if (!part) continue;
            normalized += part;
            for (let offset = 0; offset < part.length; offset++) rawMap.push(index);
        }

        const spans = [];
        const spanKeys = new Set();
        const isWordCharacter = character =>
            Boolean(character) && /[\p{L}\p{N}]/u.test(character);
        const attachedPrefixes = new Set(["و", "ك", "ب", "ل"]);

        for (const target of targets) {
            let from = 0;
            while (from <= normalized.length - target.length) {
                const position = normalized.indexOf(target, from);
                if (position < 0) break;
                from = position + Math.max(1, target.length);

                const start = rawMap[position];
                const end = rawMap[position + target.length - 1] + 1;
                if (!Number.isInteger(start) || !Number.isInteger(end) || end <= start) continue;

                const previous = start > 0 ? source[start - 1] : "";
                const beforePrevious = start > 1 ? source[start - 2] : "";
                const beforeBeforePrevious = start > 2 ? source[start - 3] : "";
                const attachedPrefix = (
                    attachedPrefixes.has(previous) &&
                    (!beforePrevious || !isWordCharacter(beforePrevious))
                ) || (
                    previous === "ل" &&
                    beforePrevious === "ل" &&
                    (!beforeBeforePrevious || !isWordCharacter(beforeBeforePrevious))
                );
                if (isWordCharacter(previous) && !attachedPrefix) continue;
                if (isWordCharacter(source[end] || "")) continue;

                const key = start + ":" + end;
                if (!spanKeys.has(key)) {
                    spanKeys.add(key);
                    spans.push({ start, end, text: source.slice(start, end) });
                }
            }
        }

        return spans.sort((a, b) => a.start - b.start || a.end - b.end);
    }



    function resolveBibliographicIdentitySpan(sourceText, reference) {
        const source = String(sourceText ?? "");
        if (!source || !reference || typeof reference !== "object") return null;
        const title = String(reference.title ?? "").trim();
        if (!title) return null;

        const titleSpans = findNormalizedTextSpans(source, title);
        if (!titleSpans.length) return null;

        const authorCandidates = [
            reference.sourceAuthor,
            reference.authorFull,
            reference.author
        ].map(value => String(value ?? "").trim()).filter(Boolean);

        let authorSpans = [];
        const seenAuthorSpans = new Set();
        for (const candidate of authorCandidates) {
            const matches = findNormalizedTextSpans(source, candidate);
            const usable = matches.filter(match =>
                !titleSpans.some(titleSpan =>
                    match.start < titleSpan.end && titleSpan.start < match.end
                )
            );
            for (const match of usable) {
                const key = match.start + ":" + match.end;
                if (!seenAuthorSpans.has(key)) {
                    seenAuthorSpans.add(key);
                    authorSpans.push(match);
                }
            }
            if (authorSpans.length) break;
        }

        const isWordCharacter = character =>
            Boolean(character) && /[\p{L}\p{N}]/u.test(character);
        function expandAttachedAuthorPrefix(start) {
            const previous = start > 0 ? source[start - 1] : "";
            const beforePrevious = start > 1 ? source[start - 2] : "";
            // The doubled lam is a spelling form of the definite-article
            // construction in "للأزهري". Do not absorb other attached
            // particles (notably ك in "كالزركشي") into the author's name.
            if (
                previous === "ل" &&
                (!beforePrevious || !isWordCharacter(beforePrevious))
            ) return start - 1;
            return start;
        }

        const separatorOnly = /^[\s\u00A0،,؛;:()[\]{}«»“”"'־\-–—]*$/u;
        const pairs = [];
        for (const titleSpan of titleSpans) {
            for (const authorSpan of authorSpans) {
                const authorIdentityStart = expandAttachedAuthorPrefix(authorSpan.start);
                let start, end, separator;
                if (titleSpan.end <= authorIdentityStart) {
                    start = titleSpan.start;
                    end = authorSpan.end;
                    separator = source.slice(titleSpan.end, authorIdentityStart);
                } else if (authorSpan.end <= titleSpan.start) {
                    start = authorIdentityStart;
                    end = titleSpan.end;
                    separator = source.slice(authorSpan.end, titleSpan.start);
                } else {
                    continue;
                }
                if (separatorOnly.test(separator)) {
                    pairs.push({
                        start, end,
                        titleStart: titleSpan.start,
                        titleEnd: titleSpan.end,
                        authorStart: authorIdentityStart,
                        authorEnd: authorSpan.end,
                        titleText: titleSpan.text,
                        authorText: source.slice(authorSpan.start, authorSpan.end),
                        includesAuthor: true,
                        authorFound: true
                    });
                }
            }
        }

        if (pairs.length === 1) return pairs[0];
        if (pairs.length > 1 || titleSpans.length !== 1) return null;

        const titleSpan = titleSpans[0];
        return {
            start: titleSpan.start,
            end: titleSpan.end,
            titleStart: titleSpan.start,
            titleEnd: titleSpan.end,
            authorStart: null,
            authorEnd: null,
            titleText: titleSpan.text,
            authorText: "",
            includesAuthor: false,
            authorFound: authorSpans.length > 0
        };
    }



    function formatBibliographicIdentity(reference, format, includeAuthor = true, authorOverride) {
        if (!reference || typeof reference !== "object") return "";
        const title = String(reference.title ?? "").trim();
        const author = String(
            authorOverride !== undefined ? authorOverride : reference.author ?? ""
        ).trim();

        switch (String(format || "author-title")) {
            case "reference-only":
                return title;
            case "title-author":
            case "title-author-first":
                return [title, includeAuthor ? author : ""].filter(Boolean).join("، ");
            case "author-title":
            default:
                return [includeAuthor ? author : "", title].filter(Boolean).join("، ");
        }
    }

    function buildCitationPreview(sourceText, reference, replacementText) {
        const source = String(sourceText ?? "");
        const identity = resolveBibliographicIdentitySpan(source, reference);
        if (!identity || (identity.authorFound && !identity.includesAuthor)) return source;
        return source.slice(0, identity.start) +
            String(replacementText ?? "") +
            source.slice(identity.end);
    }

    function getReplacementTexts(suggestion) {
        if (Array.isArray(suggestion && suggestion.replacementTexts)) {
            return suggestion.replacementTexts;
        }
        return Array.isArray(suggestion && suggestion.suggestedTexts)
            ? suggestion.suggestedTexts
            : [];
    }

    function buildPairs(noteText, suggestion) {
        const references = Array.isArray(suggestion && suggestion.references) ? suggestion.references : [];
        const replacementTexts = getReplacementTexts(suggestion);
        const sourceTexts = Array.isArray(suggestion && suggestion.sourceTexts) ? suggestion.sourceTexts : [];
        const count = Math.min(references.length, replacementTexts.length, sourceTexts.length);
        const pairs = [];
        pairs.collapsedCount = 0;
        pairs.unchangedCount = 0;

        // The same work can legitimately occur more than once in a single note.
        // Preserve each occurrence using the ordinal of its complete source anchor.
        const sourceOccurrenceCounts = new Map();

        for (let index = 0; index < count; index++) {
            const reference = references[index];
            const sourceText = String(sourceTexts[index] ?? "").trim();
            const replacementText = getIdentityText(replacementTexts[index]);
            if (!sourceText || !replacementText) continue;

            const sourceKey = sourceText;
            const sourceOccurrenceIndex = sourceOccurrenceCounts.get(sourceKey) || 0;
            let matchedSourceText = sourceText;
            let sourceMatches = findExactOccurrences(noteText, matchedSourceText);

            // If the AI supplied a bibliographic identity without an attached
            // preposition, allow the unique validated anchor finder to locate it.
            if (sourceMatches.length <= sourceOccurrenceIndex && sourceOccurrenceIndex === 0) {
                const fallbackAnchor = findUniqueBibliographicAnchor(noteText, reference);
                if (fallbackAnchor) {
                    const fallbackMatches = findExactOccurrences(noteText, fallbackAnchor);
                    if (fallbackMatches.length === 1) {
                        matchedSourceText = fallbackAnchor;
                        sourceMatches = fallbackMatches;
                    }
                }
            }
            if (sourceMatches.length <= sourceOccurrenceIndex) continue;

            const sourceStart = sourceMatches[sourceOccurrenceIndex];
            sourceOccurrenceCounts.set(sourceKey, sourceOccurrenceIndex + 1);

            // Resolve the title and author independently. Only the exact identity
            // span may be replaced; editorial and publication metadata stay outside.
            const identity = resolveBibliographicIdentitySpan(matchedSourceText, reference);
            if (!identity) continue;

            // If an author is present but is separated from the title by metadata
            // or other text, do not replace a partial range or duplicate the author.
            if (identity.authorFound && !identity.includesAuthor) continue;

            const replaceableText = matchedSourceText.slice(identity.start, identity.end);
            const start = sourceStart + identity.start;
            const end = sourceStart + identity.end;
            if (!replaceableText || noteText.slice(start, end) !== replaceableText) continue;

            if (normalizeIdentityText(replaceableText) === normalizeIdentityText(replacementText)) {
                pairs.unchangedCount++;
                continue;
            }

            pairs.push({
                sourceText: matchedSourceText,
                identityText: replaceableText,
                replacementText: replacementText,
                start: start,
                end: end,
                expectedText: replaceableText,
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

        const authors = Array.from(new Set([
            String(reference.sourceAuthor ?? "").trim(),
            String(reference.authorFull ?? "").trim(),
            String(reference.author ?? "").trim()
        ].filter(Boolean)));
        const title = normalizeBibliographicTitle(
            getIdentityText(reference.title ?? ""),
            authors[0] || ""
        );
        if (!title || !authors.length) return "";

        const candidates = [];
        const addCandidate = value => {
            const candidate = String(value ?? "").trim();
            const key = normalizeIdentityText(candidate);
            if (key.length < 8 || candidates.some(item => normalizeIdentityText(item) === key)) return;
            candidates.push(candidate);
        };

        for (const author of authors) {
            addCandidate(author + "، " + title);
            addCandidate(title + "، " + author);
            addCandidate(author + " " + title);
            addCandidate(title + " " + author);
        }

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
        // Search the exact source extracted from this same note snapshot. Ignoring
        // punctuation or spaces can enlarge/shift Word's returned range.
        // If Word cannot find the exact span, the operation is skipped safely.
        return {
            matchCase: true,
            matchWholeWord: false,
            matchWildcards: false,
            ignorePunct: false,
            ignoreSpace: false
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
                            getReplacementTexts(suggestion).length
                        )
                    );
                    continue;
                }

                const pairs = buildPairs(entry.text, suggestion);
                const requestedCount = Math.min(
                    Array.isArray(suggestion && suggestion.references) ? suggestion.references.length : 0,
                    Array.isArray(suggestion && suggestion.sourceTexts) ? suggestion.sourceTexts.length : 0,
                    getReplacementTexts(suggestion).length
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
                    const expectedFull = operation.pair.sourceText;
                    const matches = operation.fullSearch.items.filter(match =>
                        String(match.text ?? "") === expectedFull
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
                    const expectedIdentity = operation.pair.expectedText;
                    const matches = operation.identitySearch.items.filter(match =>
                        String(match.text ?? "") === expectedIdentity
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
        findNormalizedTextSpans: findNormalizedTextSpans,
        resolveBibliographicIdentitySpan: resolveBibliographicIdentitySpan,
        formatBibliographicIdentity: formatBibliographicIdentity,
        buildCitationPreview: buildCitationPreview,
        isHadithTakhrijReference: isHadithTakhrijReference,
        selectOccurrence: selectOccurrence
    };

    if (typeof module !== "undefined" && module.exports) {
        module.exports = global.FootnoteCitationEngine;
    }
})(typeof window !== "undefined" ? window : globalThis);
