/** Export formats for a built data dictionary. */
import type { ChoiceRow, ColumnDef, DictionaryRow, TabInfo } from "./dictionary";
import { buildXlsx, type SheetSpec } from "./xlsx";

export interface ExportContext {
    /** Short name of the form or view; becomes the worksheet name in Excel */
    name: string;
    title: string;
    /** Label/value pairs describing where the dictionary came from */
    meta: [string, string][];
    columns: ColumnDef[];
    rows: DictionaryRow[];
    tabs: TabInfo[];
    choices: ChoiceRow[];
    splitByTab: boolean;
}

/*
 * Every export takes a list of contexts, one per form or view. With one context the
 * output is a plain dictionary; with several, each row or section also says which
 * form it belongs to.
 */

function matrix(ctx: ExportContext, rows: DictionaryRow[] = ctx.rows): (string | number)[][] {
    return rows.map((r) => ctx.columns.map((c) => r.values[c.key] ?? ""));
}

/**
 * Spreadsheet apps run text that starts with = + - @ (or a tab or carriage return) as a
 * formula when they open a CSV or receive a paste. A leading apostrophe makes it plain text.
 * This also keeps ranges such as "-2,147,483,648 to 2,147,483,647" from turning into errors.
 */
export function neutralizeFormula(value: string | number): string | number {
    return typeof value === "string" && /^[=+\-@\t\r]/.test(value) ? `'${value}` : value;
}

function delimited(ctxs: ExportContext[], sep: string): string {
    const quote = (v: string | number) => {
        const s = String(neutralizeFormula(v));
        return s.includes(sep) || /["\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const many = ctxs.length > 1;
    const lines: (string | number)[][] = [[...(many ? ["Form"] : []), ...ctxs[0].columns.map((c) => c.header)]];
    for (const ctx of ctxs) for (const row of matrix(ctx)) lines.push(many ? [ctx.name, ...row] : row);
    return lines.map((line) => line.map(quote).join(sep)).join("\r\n");
}

/** UTF-8 with BOM so Excel opens accented labels correctly */
export function toCsv(ctxs: ExportContext[]): string {
    return "\uFEFF" + delimited(ctxs, ",");
}

/** Tab-separated, for pasting straight into Excel, Word or a wiki table */
export function toTsv(ctxs: ExportContext[]): string {
    return delimited(ctxs, "\t");
}

export function toJson(ctxs: ExportContext[]): string {
    const one = (ctx: ExportContext) => ({
        name: ctx.name,
        title: ctx.title,
        source: Object.fromEntries(ctx.meta),
        fields: ctx.rows.map((r) => Object.fromEntries(ctx.columns.map((c) => [c.header, r.values[c.key] ?? ""]))),
        choices: ctx.choices,
    });
    return JSON.stringify(ctxs.length > 1 ? { forms: ctxs.map(one) } : one(ctxs[0]), null, 2);
}

export function toMarkdown(ctxs: ExportContext[]): string {
    // "<" is escaped so text from Dataverse cannot become HTML when the Markdown is rendered
    const plain = (v: string | number) => String(v).replace(/</g, "&lt;");
    const cellText = (v: string | number) => plain(v).replace(/\\/g, "\\\\").replace(/\|/g, "\\|").replace(/\r?\n/g, "<br>");
    const one = (ctx: ExportContext) => {
        // The tab is the heading of each table, so the Tab column would only repeat it
        const columns = ctx.columns.filter((c) => c.key !== "tab");
        const out: string[] = [`# ${plain(ctx.title)}`, ""];
        for (const [k, v] of ctx.meta) out.push(`- **${k}:** ${plain(v)}`);
        for (const tab of ctx.tabs) {
            const rows = ctx.rows.filter((r) => r.groupKey === tab.key);
            if (!rows.length) continue;
            out.push("", `## ${plain(tab.label)}`, "");
            out.push(`| ${columns.map((c) => c.header).join(" | ")} |`);
            out.push(`| ${columns.map(() => "---").join(" | ")} |`);
            for (const r of rows) out.push(`| ${columns.map((c) => cellText(r.values[c.key] ?? "")).join(" | ")} |`);
        }
        return out.join("\n") + "\n";
    };
    return ctxs.map(one).join("\n---\n\n");
}

export function toXlsx(ctxs: ExportContext[]): Uint8Array {
    const many = ctxs.length > 1;
    const sheets: SheetSpec[] = [];
    for (const ctx of ctxs) {
        const headers = ctx.columns.map((c) => c.header);
        const widths = ctx.columns.map((c) => c.width);
        // The list of other forms reads best on one line, so its cells do not wrap
        const noWrap = ctx.columns.map((c) => c.key === "alsoOn");
        // One worksheet per tab only makes sense for a single form; several forms get one worksheet each
        if (!many && ctx.splitByTab && ctx.tabs.length > 1) {
            for (const tab of ctx.tabs) {
                const rows = ctx.rows.filter((r) => r.groupKey === tab.key);
                if (rows.length) sheets.push({ name: tab.label, headers, widths, noWrap, rows: matrix(ctx, rows) });
            }
        } else {
            sheets.push({ name: ctx.name || "Data Dictionary", headers, widths, noWrap, rows: matrix(ctx) });
        }
    }
    const choices = ctxs.flatMap((ctx) => ctx.choices.map((c) => [...(many ? [ctx.name] : []), c.table, c.logicalName, c.label, c.value, c.choice]));
    if (choices.length) {
        sheets.push({
            name: "Allowed Values",
            headers: [...(many ? ["Form"] : []), "Table", "Logical Name", "Label", "Value", "Choice Label"],
            widths: [...(many ? [24] : []), 18, 26, 26, 12, 34],
            rows: choices,
        });
    }
    const about: (string | number)[][] = [];
    ctxs.forEach((ctx, i) => {
        if (many) about.push(...(i ? [["", ""]] : []), ["Form", ctx.name]);
        about.push(...ctx.meta.map(([k, v]) => [k, v]));
    });
    sheets.push({ name: "About", headers: ["Property", "Value"], widths: [22, 80], rows: about, filter: false });
    return buildXlsx(sheets);
}

export function fileSlug(text: string): string {
    return (
        text
            .normalize("NFKD")
            .replace(/[^\w\s-]/g, "")
            .trim()
            .replace(/[\s_]+/g, "-")
            .toLowerCase()
            .slice(0, 60) || "data-dictionary"
    );
}
