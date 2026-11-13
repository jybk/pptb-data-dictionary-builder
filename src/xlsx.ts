/**
 * Minimal .xlsx writer: a few sheets of text/number cells with a styled, frozen,
 * filterable header row. Keeps the tool free of heavyweight spreadsheet libraries.
 */
import { strToU8, zipSync } from "fflate";

export interface SheetSpec {
    name: string;
    headers: string[];
    /** Column widths in characters */
    widths: number[];
    rows: (string | number)[][];
    /** Add an autofilter over the header row (default true) */
    filter?: boolean;
}

function xml(text: string): string {
    return text
        .replace(/[^\x09\x0A\x0D\x20-퟿-�]/g, "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

function colName(index: number): string {
    let n = index + 1;
    let name = "";
    while (n > 0) {
        const r = (n - 1) % 26;
        name = String.fromCharCode(65 + r) + name;
        n = Math.floor((n - 1) / 26);
    }
    return name;
}

/** Excel sheet names: max 31 chars, none of \ / ? * [ ] :, unique within the workbook */
export function safeSheetNames(names: string[]): string[] {
    const used = new Set<string>();
    return names.map((raw, i) => {
        const base = (raw.replace(/[\\/?*[\]:]/g, " ").replace(/^'+|'+$/g, "").trim() || `Sheet${i + 1}`).slice(0, 31);
        let name = base;
        for (let n = 2; used.has(name.toLowerCase()); n++) {
            const suffix = ` (${n})`;
            name = base.slice(0, 31 - suffix.length) + suffix;
        }
        used.add(name.toLowerCase());
        return name;
    });
}

function cell(ref: string, value: string | number, style: number): string {
    if (typeof value === "number" && Number.isFinite(value)) return `<c r="${ref}" s="${style}"><v>${value}</v></c>`;
    const text = String(value ?? "");
    if (!text) return `<c r="${ref}" s="${style}"/>`;
    // Excel caps a cell at 32,767 characters
    return `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${xml(text.slice(0, 32767))}</t></is></c>`;
}

function sheetXml(sheet: SheetSpec): string {
    const lastCol = colName(Math.max(sheet.headers.length, 1) - 1);
    const cols = sheet.widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${Math.min(Math.max(w, 4), 80) + 1}" customWidth="1"/>`).join("");
    const head = `<row r="1">${sheet.headers.map((h, i) => cell(`${colName(i)}1`, h, 1)).join("")}</row>`;
    const body = sheet.rows.map((row, r) => `<row r="${r + 2}">${row.map((v, c) => cell(`${colName(c)}${r + 2}`, v, 2)).join("")}</row>`).join("");
    const filter = sheet.filter === false ? "" : `<autoFilter ref="A1:${lastCol}${sheet.rows.length + 1}"/>`;
    return (
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        `<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` +
        `<sheetFormatPr defaultRowHeight="15"/>` +
        (cols ? `<cols>${cols}</cols>` : "") +
        `<sheetData>${head}${body}</sheetData>` +
        filter +
        `</worksheet>`
    );
}

const STYLES =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>` +
    `<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E5F"/><bgColor indexed="64"/></patternFill></fill></fills>` +
    `<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFD0D7DE"/></left><right style="thin"><color rgb="FFD0D7DE"/></right><top style="thin"><color rgb="FFD0D7DE"/></top><bottom style="thin"><color rgb="FFD0D7DE"/></bottom><diagonal/></border></borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellXfs count="3">` +
    `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
    `<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>` +
    `<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>` +
    `</cellXfs>` +
    `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
    `</styleSheet>`;

export function buildXlsx(sheets: SheetSpec[]): Uint8Array {
    const names = safeSheetNames(sheets.map((s) => s.name));
    const files: Record<string, Uint8Array> = {};
    const head = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;

    files["[Content_Types].xml"] = strToU8(
        `${head}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
            `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
            `<Default Extension="xml" ContentType="application/xml"/>` +
            `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
            `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
            sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("") +
            `</Types>`,
    );
    files["_rels/.rels"] = strToU8(
        `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
            `</Relationships>`,
    );
    files["xl/workbook.xml"] = strToU8(
        `${head}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
            `<sheets>${names.map((n, i) => `<sheet name="${xml(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets>` +
            `</workbook>`,
    );
    files["xl/_rels/workbook.xml.rels"] = strToU8(
        `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("") +
            `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
            `</Relationships>`,
    );
    files["xl/styles.xml"] = strToU8(STYLES);
    sheets.forEach((sheet, i) => {
        files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(sheetXml(sheet));
    });
    return zipSync(files, { level: 6 });
}
