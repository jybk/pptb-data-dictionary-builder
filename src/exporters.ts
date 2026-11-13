/** Export formats for a built data dictionary. */
import type { ChoiceRow, ColumnDef, DictionaryRow, TabInfo } from "./dictionary";
import { buildXlsx, type SheetSpec } from "./xlsx";

export interface ExportContext {
    title: string;
    /** Label/value pairs describing where the dictionary came from */
    meta: [string, string][];
    columns: ColumnDef[];
    rows: DictionaryRow[];
    tabs: TabInfo[];
    choices: ChoiceRow[];
    splitByTab: boolean;
}

function matrix(ctx: ExportContext, rows: DictionaryRow[] = ctx.rows): (string | number)[][] {
    return rows.map((r) => ctx.columns.map((c) => r.values[c.key] ?? ""));
}

function delimited(ctx: ExportContext, sep: string): string {
    const quote = (v: string | number) => {
        const s = String(v);
        return s.includes(sep) || /["\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    return [ctx.columns.map((c) => c.header), ...matrix(ctx)].map((line) => line.map(quote).join(sep)).join("\r\n");
}

/** UTF-8 with BOM so Excel opens accented labels correctly */
export function toCsv(ctx: ExportContext): string {
    return "﻿" + delimited(ctx, ",");
}

/** Tab-separated, for pasting straight into Excel, Word or a wiki table */
export function toTsv(ctx: ExportContext): string {
    return delimited(ctx, "\t");
}

export function toJson(ctx: ExportContext): string {
    return JSON.stringify(
        {
            title: ctx.title,
            source: Object.fromEntries(ctx.meta),
            fields: ctx.rows.map((r) => Object.fromEntries(ctx.columns.map((c) => [c.header, r.values[c.key] ?? ""]))),
            choices: ctx.choices,
        },
        null,
        2,
    );
}

export function toMarkdown(ctx: ExportContext): string {
    const cellText = (v: string | number) => String(v).replace(/\\/g, "\\\\").replace(/\|/g, "\\|").replace(/\r?\n/g, "<br>");
    // The tab is the heading of each table, so the Tab column would only repeat it
    const columns = ctx.columns.filter((c) => c.key !== "tab");
    const out: string[] = [`# ${ctx.title}`, ""];
    for (const [k, v] of ctx.meta) out.push(`- **${k}:** ${v}`);
    for (const tab of ctx.tabs) {
        const rows = ctx.rows.filter((r) => r.groupKey === tab.key);
        if (!rows.length) continue;
        out.push("", `## ${tab.label}`, "");
        out.push(`| ${columns.map((c) => c.header).join(" | ")} |`);
        out.push(`| ${columns.map(() => "---").join(" | ")} |`);
        for (const r of rows) out.push(`| ${columns.map((c) => cellText(r.values[c.key] ?? "")).join(" | ")} |`);
    }
    return out.join("\n") + "\n";
}

export function toXlsx(ctx: ExportContext): Uint8Array {
    const headers = ctx.columns.map((c) => c.header);
    const widths = ctx.columns.map((c) => c.width);
    const sheets: SheetSpec[] = [];
    if (ctx.splitByTab && ctx.tabs.length > 1) {
        for (const tab of ctx.tabs) {
            const rows = ctx.rows.filter((r) => r.groupKey === tab.key);
            if (rows.length) sheets.push({ name: tab.label, headers, widths, rows: matrix(ctx, rows) });
        }
    } else {
        sheets.push({ name: "Data Dictionary", headers, widths, rows: matrix(ctx) });
    }
    if (ctx.choices.length) {
        sheets.push({
            name: "Allowed Values",
            headers: ["Table", "Logical Name", "Label", "Value", "Choice Label"],
            widths: [18, 26, 26, 12, 34],
            rows: ctx.choices.map((c) => [c.table, c.logicalName, c.label, c.value, c.choice]),
        });
    }
    sheets.push({ name: "About", headers: ["Property", "Value"], widths: [22, 80], rows: ctx.meta.map(([k, v]) => [k, v]), filter: false });
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
