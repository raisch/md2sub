// md2sub.mjs
// -------------------------------------------------------------
// Convert Markdown -> DOCX in "modern short-story submission" format.
// Fix: docx TextRun instances don't reliably expose `.options` across versions.
// We now track "run parts" ourselves (plain JS objects) and only create TextRuns
// at the last moment. This avoids undefined `.options.text` crashes.
// -------------------------------------------------------------

import fs from "node:fs/promises";
import path from "node:path";

import { unified } from "unified";
import remarkParse from "remark-parse";
import remarkGfm from "remark-gfm";
import { visit } from "unist-util-visit";

import {
    AlignmentType,
    BorderStyle,
    Document,
    Footer,
    Header,
    LineRuleType,
    Packer,
    Paragraph,
    SimpleField,
    Table,
    TableCell,
    TableLayoutType,
    TableRow,
    TextRun,
    WidthType,
} from "docx";

// -----------------------------
// DOCX unit constants (per spec)
// -----------------------------
const TWIPS_PER_INCH = 1440;
const HALF_INCH_TWIPS = 720;
const FONT_SIZE_12PT = 24; // docx uses half-points (12pt -> 24)
const DOUBLE_LINE_SPACING = 480; // approx "double" in docx (twentieths of a point)
const US_LETTER = { width: 12240, height: 15840 }; // 8.5" x 11" in twips
// Usable line width on US Letter with 1" margins:
// 8.5" = 12240 twips, minus 2" margins = 2*1440 => 12240 - 2880 = 9360
const USABLE_WIDTH_TWIPS = 9360;

// -----------------------------
// Additional layout constants (tuned to look good in Pages)
// -----------------------------
const BLANK_LINES_BEFORE_BODY = 2;

// -----------------------------
// CLI + basic validation
// -----------------------------
function usageAndExit(msg) {
    const text =
        (msg ? `Error: ${msg}\n\n` : "") +
        "Usage:\n  node md-to-submission-docx.mjs input.md [output.docx]\n\n" +
        "Example:\n  node md-to-submission-docx.mjs story.md submission.docx\n";
    console.error(text);
    process.exit(1);
}

const [, , inputPathArg, outputPathArg] = process.argv;
if (!inputPathArg) usageAndExit("input.md is required.");

const inputPath = path.resolve(process.cwd(), inputPathArg);
const outputPath = path.resolve(process.cwd(), outputPathArg || "submission.docx");

// -----------------------------
// Parse top-of-file constants block
// -----------------------------
function parseFrontMatterBlock(md) {
    const lines = md.replace(/\r\n/g, "\n").split("\n");

    const meta = {
        Author: "",
        LastName: "",
        Address: "",
        City: "",
        State: "",
        PostalCode: "",
        Phone: "",
        Email: "",
        ShortTitle: "",
    };

    let i = 0;
    let sawAnyKey = false;
    let currentKey = null;

    const isLikelyContentStart = (line) => {
        const t = line.trim();
        if (!t) return false;
        if (t.startsWith("#")) return true;
        if (t === "---" || t === "***" || t === "___") return true;
        if (/^(\*|-|\+)\s+/.test(t)) return true;
        if (/^\d+\.\s+/.test(t)) return true;
        if (/^\|.*\|$/.test(t)) return true;
        return false;
    };

    for (; i < lines.length; i++) {
        const line = lines[i];

        if (sawAnyKey && line.trim() === "") {
            i++;
            break;
        }

        if (!sawAnyKey && isLikelyContentStart(line)) {
            i = 0;
            break;
        }

        const isIndented = /^\s{2,}\S/.test(line);
        if (isIndented && currentKey) {
            meta[currentKey] = (meta[currentKey] ? meta[currentKey] + "\n" : "") + line.trim();
            continue;
        }

        const m = line.match(/^([A-Za-z][A-Za-z0-9]*)\s*:\s*(.*)$/);
        if (m) {
            const key = m[1];
            const value = m[2] ?? "";
            sawAnyKey = true;
            currentKey = key;
            meta[key] = value.trim();
            continue;
        }

        if (sawAnyKey) break;
    }

    const markdownBody = lines.slice(i).join("\n");
    return { meta, markdownBody };
}

// -----------------------------
// Markdown -> MDAST
// -----------------------------
async function parseMarkdown(mdText) {
    const processor = unified().use(remarkParse).use(remarkGfm);
    return processor.parse(mdText);
}

// -----------------------------
// Inline conversion helpers
// We build "run parts" as plain objects and only later create TextRuns.
// -----------------------------
/**
 * @typedef {{ text: string, bold?: boolean, italics?: boolean }} RunPart
 */

function mergeStyleIntoParts(parts, style) {
    return parts.map((p) => ({ ...p, ...style }));
}

function inlineNodeToParts(node) {
    switch (node.type) {
        case "text":
            return [{ text: node.value ?? "" }];

        case "strong": {
            const parts = [];
            for (const c of node.children || []) parts.push(...inlineNodeToParts(c));
            return mergeStyleIntoParts(parts, { bold: true });
        }

        case "emphasis": {
            const parts = [];
            for (const c of node.children || []) parts.push(...inlineNodeToParts(c));
            return mergeStyleIntoParts(parts, { italics: true });
        }

        case "inlineCode":
            // Treat inline code as plain text (no special formatting)
            return [{ text: node.value ?? "" }];

        case "link": {
            const labelParts = [];
            for (const c of node.children || []) labelParts.push(...inlineNodeToParts(c));
            const labelText = labelParts.map((p) => p.text).join("");
            const url = node.url || "";
            const combined = url ? `${labelText} (${url})` : labelText;
            return [{ text: combined }];
        }

        case "delete": {
            const parts = [];
            for (const c of node.children || []) parts.push(...inlineNodeToParts(c));
            return parts;
        }

        case "break":
            return [{ text: "\n" }];

        default: {
            if (node.children && Array.isArray(node.children)) {
                const parts = [];
                for (const c of node.children) parts.push(...inlineNodeToParts(c));
                return parts;
            }
            return [{ text: "" }];
        }
    }
}

function phrasingChildrenToParts(children = []) {
    const parts = [];
    for (const c of children) parts.push(...inlineNodeToParts(c));
    return parts;
}

function partsToTextRuns(parts) {
    const safe = parts.filter((p) => typeof p?.text === "string");
    return safe.map(
        (p) =>
            new TextRun({
                text: p.text,
                bold: !!p.bold,
                italics: !!p.italics,
                font: "Times New Roman",
                size: 24,
            })
    );
}

function partsAreAllWhitespace(parts) {
    return parts.every((p) => (p?.text ?? "").trim() === "");
}

// -----------------------------
// Paragraph style helpers
// -----------------------------
function baseParagraphOptions() {
    return {
        spacing: {
            line: DOUBLE_LINE_SPACING,          // 480
            lineRule: LineRuleType.AUTO,        // Pages respects this more than implicit defaults
            before: 0,
            after: 0,
        },
    };
}

function centeredParagraph(children) {
    return new Paragraph({
        ...baseParagraphOptions(),
        style: "Centered", // IMPORTANT for Pages
        alignment: AlignmentType.CENTER,
        indent: { left: 0, right: 0, firstLine: 0 },
        children,
    });
}

function singleSpaceParagraphOptions() {
    return {
        spacing: {
            line: 240,                 // single spacing
            lineRule: LineRuleType.AUTO,
            before: 0,
            after: 0,
        },
    };
}

function bodyParagraph(parts) {
    return new Paragraph({
        ...baseParagraphOptions(),
        style: "Normal", // IMPORTANT for Pages
        indent: { firstLine: HALF_INCH_TWIPS, left: 0, right: 0 },
        children: partsToTextRuns(parts),
    });
}

function sceneBreakParagraph() {
    return centeredParagraph([
        new TextRun({
            text: "#",
            font: "Times New Roman",
            size: 24
        })
    ]);
}


function contactLineParagraph(text) {
    return new Paragraph({
        ...singleSpaceParagraphOptions(),
        style: "Normal",
        alignment: AlignmentType.LEFT,
        indent: { firstLine: 0, left: 0, right: 0 },
        children: [new TextRun({ text, font: "Times New Roman", size: 24 })],
    });
}

function blankLine() {
    return new Paragraph({
        ...baseParagraphOptions(),
        children: [new TextRun({ text: "" })],
    });
}

// -----------------------------
// Word count helpers
// -----------------------------
function countWordsFromMdast(tree, { skipTitleByline = true } = {}) {
    const kids = tree.children || [];

    // If you already skip the first H1 (title) and first H2 (byline) in your conversion,
    // mirror that logic here so the count matches "body words".
    const skip = new Set();
    if (skipTitleByline) {
        let sawH1 = false;
        for (let i = 0; i < kids.length; i++) {
            const n = kids[i];
            if (!sawH1 && n.type === "heading" && n.depth === 1) {
                sawH1 = true;
                skip.add(i);

                // skip immediate following empty paragraphs
                for (let j = i + 1; j < kids.length; j++) {
                    const m = kids[j];
                    const isEmptyPara =
                        m.type === "paragraph" &&
                        (!m.children || m.children.every((c) => c.type === "text" && !c.value.trim()));
                    if (isEmptyPara) {
                        skip.add(j);
                        continue;
                    }
                    if (m.type === "heading" && m.depth === 2) skip.add(j); // byline
                    break;
                }
                break;
            }
        }
    }

    let text = "";
    for (let i = 0; i < kids.length; i++) {
        if (skip.has(i)) continue;
        const node = kids[i];

        // Ignore thematic breaks in word count
        if (node.type === "thematicBreak") continue;

        // Collect plain text from all textual descendants
        visit(node, (n) => {
            if (n.type === "text") text += " " + n.value;
            if (n.type === "inlineCode") text += " " + n.value;
            if (n.type === "code") text += " " + (n.value || "");
        });
    }

    // Count words
    const words = (text.match(/[A-Za-z0-9]+(?:'[A-Za-z0-9]+)?/g) || []).length;
    return words;
}

function roundToNearest100(n) {
    return Math.round(n / 100) * 100;
}

function authorAndWordCountLine(authorName, approxWordCount) {
    const left = authorName || "";
    const right = approxWordCount ? `about ${approxWordCount} words` : "";

    // Manuscript font is monospaced? No—Times is proportional.
    // But for Pages reliability, we approximate the right margin using spaces.
    // Tune TARGET_COLS if you want tighter/looser alignment in Pages.
    const TARGET_COLS = 92;

    const leftLen = left.length;
    const rightLen = right.length;

    const spacesNeeded = Math.max(2, TARGET_COLS - (leftLen + rightLen));
    const pad = " ".repeat(spacesNeeded);

    return new Paragraph({
        ...singleSpaceParagraphOptions(),
        style: "Normal",
        alignment: AlignmentType.LEFT,
        indent: { firstLine: 0, left: 0, right: 0 },
        children: [
            new TextRun({ text: left + pad + right, font: "Times New Roman", size: 24 }),
        ],
    });
}


// -----------------------------
// Lists (simple manuscript style)
// -----------------------------
function listItemParagraph(prefix, parts, level = 0) {
    const leftIndent = HALF_INCH_TWIPS + level * 360;
    const hanging = 360;
    return new Paragraph({
        ...baseParagraphOptions(),
        indent: { left: leftIndent, hanging },
        children: [new TextRun({ text: prefix + " " }), ...partsToTextRuns(parts)],
    });
}

// -----------------------------
// Tables (minimal GFM)
// -----------------------------
function mdTableToDocxTable(tableNode) {
    const rowsMd = tableNode.children || [];
    const colCount = rowsMd[0]?.children?.length || 1;

    // Even split across the usable line width
    const base = Math.floor(USABLE_WIDTH_TWIPS / colCount);
    const widths = Array(colCount).fill(base);
    // Put any remainder into the last column to make the sum exact.
    widths[colCount - 1] += USABLE_WIDTH_TWIPS - widths.reduce((a, b) => a + b, 0);

    const rows = rowsMd.map((rowNode) => {
        const cellsMd = rowNode.children || [];

        const cells = [];
        for (let i = 0; i < colCount; i++) {
            const cellNode = cellsMd[i];
            const cellParts = [];

            if (cellNode) {
                for (const c of cellNode.children || []) cellParts.push(...inlineNodeToParts(c));
            }

            cells.push(
                new TableCell({
                    // Keep setting cell width too (some importers use it), but the key is columnWidths below.
                    width: { size: widths[i], type: WidthType.DXA },

                    children: [
                        new Paragraph({
                            ...baseParagraphOptions(),
                            style: "Normal",
                            indent: { firstLine: 0, left: 0, right: 0 },
                            children: partsToTextRuns(cellParts.length ? cellParts : [{ text: "" }]),
                        }),
                    ],
                })
            );
        }

        return new TableRow({ children: cells });
    });

    return new Table({
        // Force a real table grid (THIS is what Pages tends to honor)
        columnWidths: widths,

        width: { size: USABLE_WIDTH_TWIPS, type: WidthType.DXA },
        layout: TableLayoutType.FIXED,

        // Prevent Pages from centering it
        alignment: AlignmentType.LEFT,

        // Optional: remove borders (keep if you want manuscript-clean tables)
        borders: {
            top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        },

        rows,
    });
}



// -----------------------------
// Block conversion
// -----------------------------
function convertMdastToDocxBlocks(tree) {
    const rootChildren = tree.children || [];

    let titleText = "";
    let bylineText = "";
    const skipIndices = new Set();

    for (let idx = 0; idx < rootChildren.length; idx++) {
        const n = rootChildren[idx];
        if (!titleText && n.type === "heading" && n.depth === 1) {
            titleText = extractPlainText(n);
            skipIndices.add(idx);

            for (let j = idx + 1; j < rootChildren.length; j++) {
                const m = rootChildren[j];
                if (m.type === "paragraph" && isNodeVisiblyEmpty(m)) {
                    skipIndices.add(j);
                    continue;
                }
                if (m.type === "heading" && m.depth === 2) {
                    bylineText = extractPlainText(m);
                    skipIndices.add(j);
                }
                break;
            }
            break;
        }
    }

    const blocks = [];

    for (let idx = 0; idx < rootChildren.length; idx++) {
        if (skipIndices.has(idx)) continue;
        const node = rootChildren[idx];

        switch (node.type) {
            case "paragraph": {
                const parts = phrasingChildrenToParts(node.children || []);
                if (parts.length === 0 || partsAreAllWhitespace(parts)) {
                    blocks.push(blankLine());
                } else {
                    blocks.push(bodyParagraph(parts));
                }
                break;
            }

            case "heading": {
                const parts = phrasingChildrenToParts(node.children || []);
                const isH1 = node.depth === 1;

                const styledParts = isH1 ? mergeStyleIntoParts(parts, { bold: true }) : parts;
                blocks.push(centeredParagraph(partsToTextRuns(styledParts)));
                break;
            }


            case "thematicBreak":
                blocks.push(sceneBreakParagraph());
                break;

            case "list": {
                const ordered = !!node.ordered;
                const start = typeof node.start === "number" ? node.start : 1;
                const items = node.children || [];
                let n = start;

                for (const item of items) {
                    const itemParts = listItemToParts(item);
                    const prefix = ordered ? `${n}.` : "•";
                    blocks.push(listItemParagraph(prefix, itemParts, 0));
                    n++;
                }
                break;
            }

            case "table":
                blocks.push(mdTableToDocxTable(node));
                break;

            case "code": {
                const codeText = node.value || "";
                const lines = codeText.split("\n");
                for (const ln of lines) {
                    blocks.push(
                        new Paragraph({
                            ...baseParagraphOptions(),
                            indent: { left: HALF_INCH_TWIPS },
                            children: [new TextRun({ text: ln })],
                        })
                    );
                }
                break;
            }

            case "blockquote": {
                for (const child of node.children || []) {
                    if (child.type === "paragraph") {
                        const parts = phrasingChildrenToParts(child.children || []);
                        blocks.push(
                            new Paragraph({
                                ...baseParagraphOptions(),
                                indent: { left: HALF_INCH_TWIPS, firstLine: 0 },
                                children: partsToTextRuns(parts),
                            })
                        );
                    }
                }
                break;
            }

            default:
                break;
        }
    }

    return { titleText, bylineText, bodyBlocks: blocks };
}

function extractPlainText(node) {
    let out = "";
    visit(node, (n) => {
        if (n.type === "text") out += n.value;
        if (n.type === "inlineCode") out += n.value;
    });
    return out.trim();
}

function isNodeVisiblyEmpty(paragraphNode) {
    const text = extractPlainText(paragraphNode);
    return text.trim().length === 0;
}

function listItemToParts(listItemNode) {
    const parts = [];
    const children = listItemNode.children || [];
    let firstPara = true;

    for (const c of children) {
        if (c.type === "paragraph") {
            if (!firstPara) parts.push({ text: " " });
            parts.push(...phrasingChildrenToParts(c.children || []));
            firstPara = false;
        } else if (c.type === "list") {
            parts.push({ text: " " });
            const nestedItems = c.children || [];
            let idx = 0;
            for (const ni of nestedItems) {
                if (idx > 0) parts.push({ text: " " });
                parts.push({ text: c.ordered ? `${(c.start || 1) + idx}. ` : "• " });
                parts.push(...listItemToParts(ni));
                idx++;
            }
        }
    }

    return parts;
}

// -----------------------------
// Header (pages 2-N only)
// -----------------------------
function makeHeader(lastName, shortTitle) {
    const safeLast = (lastName || "").trim() || "LASTNAME";
    const safeShort = (shortTitle || "").trim() || "Short Title";

    return new Header({
        children: [
            new Paragraph({
                ...baseParagraphOptions(),
                style: "Normal",
                alignment: AlignmentType.RIGHT,
                children: [
                    new TextRun({
                        text: `${safeLast} / ${safeShort} / `,
                        font: "Times New Roman",
                        size: 24,
                    }),
                    // Pages is more reliable with a SimpleField PAGE than PageNumber.CURRENT
                    new SimpleField("PAGE"),
                ],
            }),
        ],
    });
}

function makeEmptyHeader() {
    // Pages sometimes fails to honor "different first page" when the first header part is empty.
    return new Header({
        children: [
            new Paragraph({
                ...baseParagraphOptions(),
                children: [new TextRun({ text: "" })],
            }),
        ],
    });
}


function makeEmptyFooter() {
    return new Footer({ children: [] });
}

// -----------------------------
// First page front matter
// -----------------------------
function buildFirstPageFrontMatter(meta, titleText, bylineText, approxWordCount) {
    const pieces = [];

    const contactLines = [];


    // First line: Author (left) + about N words (right)
    if (meta.Author) {
        pieces.push(authorAndWordCountLine(meta.Author, approxWordCount));
    }

    if (meta.Address) {
        const addrLines = meta.Address.split("\n").map((s) => s.trim()).filter(Boolean);
        contactLines.push(...addrLines);
    }

    const cityStateZip = [meta.City, meta.State, meta.PostalCode].filter(Boolean).join(" ").trim();
    if (cityStateZip) contactLines.push(cityStateZip);

    if (meta.Phone) contactLines.push(meta.Phone);
    if (meta.Email) contactLines.push(meta.Email);

    if (contactLines.length) {
        for (const line of contactLines) pieces.push(contactLineParagraph(line));
    }

    // Approximate "halfway down" with blank lines.
    const BLANK_LINES_BEFORE_TITLE = 5;
    for (let i = 0; i < BLANK_LINES_BEFORE_TITLE; i++) pieces.push(blankLine());

    const safeTitle = (titleText || "Untitled").toUpperCase();
    pieces.push(
        centeredParagraph([
            new TextRun({
                text: safeTitle,
                bold: true,
                font: "Times New Roman",
                size: 24,
            }),
        ])
    );

    if (bylineText) {
        pieces.push(
            centeredParagraph([
                new TextRun({
                    text: bylineText,
                    font: "Times New Roman",
                    size: 24,
                }),
            ])
        );
    }

    for (let i = 0; i < BLANK_LINES_BEFORE_BODY; i++) {
        pieces.push(blankLine());
    }

    return pieces;
}

// -----------------------------
// Main
// -----------------------------
async function main() {
    let mdRaw;
    try {
        mdRaw = await fs.readFile(inputPath, "utf8");
    } catch (err) {
        usageAndExit(`Could not read input file: ${inputPath}\n${err.message}`);
    }

    const { meta, markdownBody } = parseFrontMatterBlock(mdRaw);
    const tree = await parseMarkdown(markdownBody);

    const rawWordCount = countWordsFromMdast(tree, { skipTitleByline: true });
    const approxWordCount = roundToNearest100(rawWordCount);

    const { titleText, bylineText, bodyBlocks } = convertMdastToDocxBlocks(tree);

    const doc = new Document({
        styles: {
            default: {
                document: {
                    run: { font: "Times New Roman", size: 24 },
                },
            },
            paragraphStyles: [
                {
                    id: "Normal",
                    name: "Normal",
                    basedOn: "Normal",
                    quickFormat: true,
                    run: { font: "Times New Roman", size: 24 },
                    paragraph: {
                        spacing: { line: 480, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
                        indent: { firstLine: 720, left: 0, right: 0 },
                    },
                },
                {
                    id: "Centered",
                    name: "Centered",
                    basedOn: "Normal",
                    quickFormat: true,
                    run: { font: "Times New Roman", size: 24 },
                    paragraph: {
                        spacing: { line: 480, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
                        indent: { firstLine: 0, left: 0, right: 0 },
                    },
                },
            ],
        },

        // sections: [
        //     // SECTION 1: Title page only (NO header)
        //     {
        //         properties: {
        //             page: {
        //                 size: US_LETTER,
        //                 margin: {
        //                     top: TWIPS_PER_INCH,
        //                     right: TWIPS_PER_INCH,
        //                     bottom: TWIPS_PER_INCH,
        //                     left: TWIPS_PER_INCH,
        //                 },
        //             },
        //         },
        //         headers: { default: makeEmptyHeader() },
        //         footers: { default: makeEmptyFooter() },
        //         children: buildFirstPageFrontMatter(meta, titleText, bylineText),
        //     },

        //     // SECTION 2: Body (header enabled on ALL pages in this section)
        //     {
        //         properties: {
        //             page: {
        //                 size: US_LETTER,
        //                 margin: {
        //                     top: TWIPS_PER_INCH,
        //                     right: TWIPS_PER_INCH,
        //                     bottom: TWIPS_PER_INCH,
        //                     left: TWIPS_PER_INCH,
        //                 },
        //             },
        //         },
        //         headers: { default: makeHeader(meta.LastName, meta.ShortTitle) },
        //         footers: { default: makeEmptyFooter() },
        //         children: bodyBlocks,
        //     },
        // ],

        sections: [
            {
                properties: {
                    page: {
                        size: US_LETTER,
                        margin: {
                            top: TWIPS_PER_INCH,
                            right: TWIPS_PER_INCH,
                            bottom: TWIPS_PER_INCH,
                            left: TWIPS_PER_INCH,
                        },
                    },

                    // Key: header on pages 2+ only
                    titlePage: true,
                },

                headers: {
                    first: makeEmptyHeader(), // page 1 header blank
                    default: makeHeader(meta.LastName, meta.ShortTitle), // pages 2+
                },

                footers: {
                    first: makeEmptyFooter(),
                    default: makeEmptyFooter(),
                },

                // Title/byline + body all in one continuous flow
                children: [
                    ...buildFirstPageFrontMatter(meta, titleText, bylineText, approxWordCount),
                    ...bodyBlocks,
                ],
            },
        ],

    });

    const buf = await Packer.toBuffer(doc);
    await fs.writeFile(outputPath, buf);
    console.log(`Wrote ${outputPath}`);
}

main().catch((err) => {
    console.error("Fatal error:", err);
    process.exit(1);
});
