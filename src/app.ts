/// <reference types="@pptb/types" />
import { listSurfaces, loadForm, loadTable, loadView, type SurfaceOption, type SurfaceRecord } from "./dataverse";
import { buildDictionary, COLUMNS, type ColumnDef, type Dictionary, type DictionaryRow, type TableMeta } from "./dictionary";
import { fileSlug, toCsv, toJson, toMarkdown, toTsv, toXlsx, type ExportContext } from "./exporters";
import { parseFormXml, parseViewLayout, type ParsedLayout } from "./layoutParser";
import { hostMismatch, parseTarget, type ParsedTarget } from "./urlParser";

interface Notice {
    level: "info" | "warning" | "error";
    text: string;
}

interface Options {
    includeHidden: boolean;
    includeControls: boolean;
    uniqueOnly: boolean;
    splitByTab: boolean;
}

/** One built form or view. Several can be collected and exported together. */
interface FormEntry {
    /** Editable name; the worksheet name in Excel */
    name: string;
    target: ParsedTarget;
    surface: SurfaceRecord;
    layout: ParsedLayout;
    tables: Map<string, TableMeta>;
    lcid: number;
    tabFilter: Set<string>;
    pickerOptions: SurfaceOption[] | null;
    pickerTable: string;
}

const LANGUAGES: Record<number, string> = {
    1030: "Danish",
    1031: "German",
    1033: "English",
    1035: "Finnish",
    1036: "French",
    1040: "Italian",
    1041: "Japanese",
    1043: "Dutch",
    1044: "Norwegian",
    1045: "Polish",
    1046: "Portuguese (Brazil)",
    1053: "Swedish",
    2052: "Chinese (Simplified)",
    2070: "Portuguese",
    3082: "Spanish",
};

const state = {
    connection: null as { name: string; url: string } | null,
    /** Built forms and views, in the order they were added */
    forms: [] as FormEntry[],
    /** Index of the form on screen; -1 when there is none */
    current: -1,
    /** What the load in progress does when it succeeds */
    mode: "replace" as "replace" | "add",
    /** The form/view chooser is on screen instead of the result */
    picking: false,
    target: null as ParsedTarget | null,
    notices: [] as Notice[],
    pickerOptions: null as SurfaceOption[] | null,
    pickerTable: "",
    surface: null as SurfaceRecord | null,
    layout: null as ParsedLayout | null,
    tables: new Map<string, TableMeta>(),
    dict: null as Dictionary | null,
    options: { includeHidden: false, includeControls: false, uniqueOnly: false, splitByTab: false } as Options,
    lcid: 1033,
    columns: new Set(COLUMNS.filter((c) => c.on).map((c) => c.key)),
    tabFilter: new Set<string>(),
    search: "",
    busy: false,
    runId: 0,
};

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

function esc(value: unknown): string {
    return String(value ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c] as string);
}

function errorText(error: unknown): string {
    const message = error instanceof Error ? error.message : String(error);
    return message.replace(/^Error invoking remote method '[^']+': (Error: )?/, "");
}

async function notify(title: string, body: string, type: "success" | "error" | "info" = "success"): Promise<void> {
    try {
        await window.toolboxAPI.utils.showNotification({ title, body, type, duration: 4000 });
    } catch {
        /* notifications are a nicety */
    }
}

/* ---------- connection, theme, settings ---------- */

async function refreshConnection(): Promise<void> {
    try {
        const c = await window.toolboxAPI.connections.getActiveConnection();
        state.connection = c ? { name: c.name, url: c.url } : null;
    } catch {
        state.connection = null;
    }
    const el = $("conn");
    el.classList.toggle("missing", !state.connection);
    el.innerHTML = state.connection
        ? `<strong>${esc(state.connection.name)}</strong> · ${esc(state.connection.url.replace(/^https?:\/\//, "").replace(/\/$/, ""))}`
        : "No connection. Select one in ToolBox, then build.";
}

async function applyTheme(): Promise<void> {
    try {
        document.documentElement.dataset.theme = await window.toolboxAPI.utils.getCurrentTheme();
    } catch {
        document.documentElement.dataset.theme = "light";
    }
}

async function loadSettings(): Promise<void> {
    try {
        const saved = await window.toolboxAPI.settings.getAll();
        if (Array.isArray(saved?.columns)) {
            const known = new Set(COLUMNS.map((c) => c.key));
            state.columns = new Set((saved.columns as string[]).filter((k) => known.has(k)));
        }
        if (saved?.options && typeof saved.options === "object") state.options = { ...state.options, ...saved.options };
    } catch {
        /* defaults are fine */
    }
}

function saveSettings(): void {
    window.toolboxAPI.settings.setAll({ columns: [...state.columns], options: state.options }).catch(() => undefined);
}

/* ---------- loading ---------- */

function setBusy(busy: boolean, label = "Build dictionary"): void {
    state.busy = busy;
    const button = $<HTMLButtonElement>("build");
    button.disabled = busy;
    button.textContent = label;
    $<HTMLButtonElement>("add").disabled = busy;
}

function currentForm(): FormEntry | null {
    return state.forms[state.current] ?? null;
}

/** Puts one of the built forms on screen */
function activate(index: number): void {
    const form = state.forms[index];
    state.current = form ? index : -1;
    state.surface = form?.surface ?? null;
    state.layout = form?.layout ?? null;
    state.tables = form?.tables ?? new Map();
    state.lcid = form?.lcid ?? 1033;
    state.tabFilter = form?.tabFilter ?? new Set();
    state.dict = null;
    rebuild();
}

function removeForm(index: number): void {
    state.forms.splice(index, 1);
    activate(Math.min(index < state.current ? state.current - 1 : state.current, state.forms.length - 1));
}

/**
 * Reads the address in the box. "replace" swaps out the form on screen, "add" keeps it
 * and adds the new one next to it. Nothing on screen changes until the read succeeds.
 */
async function start(mode: "replace" | "add"): Promise<void> {
    if (state.busy) return;
    const runId = ++state.runId;
    const target = parseTarget($<HTMLInputElement>("url").value);
    state.notices = [];

    if (target.kind === "unknown") {
        state.notices.push({
            level: "error",
            text: "That address does not name a form, view or table. Open the record or view in your app and copy the address from the browser, or enter a table name such as account.",
        });
        return render();
    }
    state.mode = mode;
    state.target = target;
    state.picking = false;
    state.pickerOptions = null;

    setBusy(true, "Reading…");
    render();
    try {
        await refreshConnection();
        if (!state.connection) throw new Error("There is no active connection. Select a connection for this tool in ToolBox and try again.");
        const mismatch = hostMismatch(target, state.connection.url);
        if (mismatch) state.notices.push({ level: "warning", text: mismatch });

        let surface: SurfaceRecord | null = null;
        if (target.kind === "form") surface = await loadForm(target.formId as string);
        else if (target.kind === "view") surface = await loadView(target.viewId as string);
        else if (target.kind === "id") surface = (await loadForm(target.anyId as string)) ?? (await loadView(target.anyId as string));

        if (runId !== state.runId) return;
        if (surface) {
            await showSurface(surface, runId);
        } else if (target.etn) {
            if (target.kind !== "table") {
                state.notices.push({ level: "warning", text: `The ${target.kind} in that address was not found in this environment. Choose a form or view of ${target.etn} instead.` });
            }
            await showPicker(target.etn, target.appId, runId);
        } else {
            throw new Error("No form or view with that id exists in the connected environment. Check that the active connection is the environment the address came from.");
        }
    } catch (error) {
        if (runId === state.runId) state.notices.push({ level: "error", text: errorText(error) });
    } finally {
        if (runId === state.runId) {
            setBusy(false);
            render();
        }
    }
}

async function showPicker(etn: string, appId: string | undefined, runId: number): Promise<void> {
    const options = await listSurfaces(etn, appId);
    if (runId !== state.runId) return;
    if (!options.length) throw new Error(`No active forms or views were found for the table "${etn}". Check the table name and the active connection.`);
    state.pickerOptions = options;
    state.pickerTable = etn;
    state.picking = true;
}

async function openOption(option: SurfaceOption): Promise<void> {
    if (state.busy) return;
    const runId = ++state.runId;
    state.notices = state.notices.filter((n) => n.level === "warning" && n.text.startsWith("The URL points"));
    setBusy(true, "Reading…");
    try {
        const surface = option.kind === "form" ? await loadForm(option.id) : await loadView(option.id);
        if (!surface) throw new Error(`"${option.name}" could not be read. It may have been deleted, or you may not have access to it.`);
        if (runId === state.runId) await showSurface(surface, runId);
    } catch (error) {
        if (runId === state.runId) state.notices.push({ level: "error", text: errorText(error) });
    } finally {
        if (runId === state.runId) {
            setBusy(false);
            render();
        }
    }
}

async function showSurface(surface: SurfaceRecord, runId: number): Promise<void> {
    const parse = (xml: string) => new DOMParser().parseFromString(xml, "text/xml");
    const layout = surface.kind === "form" ? parseFormXml(surface.xml, parse) : parseViewLayout(surface.xml, surface.fetchXml, parse);
    const lcid = layout.languages[0] ?? 1033;

    const tableNames = new Set<string>([surface.etn]);
    for (const item of layout.items) if (item.table && !item.table.startsWith("?")) tableNames.add(item.table);
    const tables = new Map<string, TableMeta>();
    await Promise.all(
        [...tableNames].map(async (name) => {
            try {
                tables.set(name, await loadTable(name, lcid));
            } catch (error) {
                // The surface's own table is essential; linked tables degrade to rows without metadata
                if (name === surface.etn) throw error;
                state.notices.push({ level: "warning", text: `Column details for the linked table "${name}" could not be read: ${errorText(error)}` });
            }
        }),
    );
    if (runId !== state.runId) return;

    const entry: FormEntry = {
        name: defaultName(surface.name),
        target: state.target as ParsedTarget,
        surface,
        layout,
        tables,
        lcid,
        tabFilter: new Set(),
        pickerOptions: state.pickerOptions,
        pickerTable: state.pickerTable,
    };
    if (state.mode === "add" || state.current < 0) state.forms.push(entry);
    else state.forms[state.current] = entry;
    state.picking = false;
    activate(state.mode === "add" || state.current < 0 ? state.forms.length - 1 : state.current);
    if (state.dict?.multipleTables) state.columns.add("table");
}

/** Excel allows 31 characters in a worksheet name; keep names apart when the same form is added twice */
function defaultName(name: string): string {
    const base = name.trim().slice(0, 31) || "Form";
    const taken = new Set(state.forms.filter((_, i) => state.mode === "add" || i !== state.current).map((f) => f.name.toLowerCase()));
    let candidate = base;
    for (let n = 2; taken.has(candidate.toLowerCase()); n++) {
        const suffix = ` (${n})`;
        candidate = base.slice(0, 31 - suffix.length) + suffix;
    }
    return candidate;
}

function rebuild(): void {
    if (!state.surface || !state.layout) return;
    state.dict = buildDictionary(state.layout.items, state.surface.etn, state.tables, { ...state.options, lcid: state.lcid });
    const keys = new Set(state.dict.tabs.map((t) => t.key));
    for (const key of [...state.tabFilter]) if (!keys.has(key)) state.tabFilter.delete(key);
}

/* ---------- derived views ---------- */

function activeColumns(): ColumnDef[] {
    return COLUMNS.filter((c) => state.columns.has(c.key));
}

function filterRows(dict: Dictionary, tabFilter: Set<string>): DictionaryRow[] {
    const needle = state.search.trim().toLowerCase();
    const columns = activeColumns();
    return dict.rows.filter((r) => {
        if (tabFilter.size && !tabFilter.has(r.groupKey)) return false;
        if (!needle) return true;
        return columns.some((c) => String(r.values[c.key] ?? "").toLowerCase().includes(needle)) || String(r.values.logicalName).toLowerCase().includes(needle);
    });
}

function shownRows(): DictionaryRow[] {
    return state.dict ? filterRows(state.dict, state.tabFilter) : [];
}

function tableLabel(form: FormEntry | null = currentForm()): string {
    if (!form) return "";
    const etn = form.surface.etn;
    const meta = form.tables.get(etn);
    return meta && meta.displayName !== etn ? `${meta.displayName} (${etn})` : etn;
}

function exportContext(form: FormEntry): ExportContext {
    const s = form.surface;
    // The form on screen already has its dictionary; the others are built with the same options
    const dict = form === currentForm() && state.dict ? state.dict : buildDictionary(form.layout.items, s.etn, form.tables, { ...state.options, lcid: form.lcid });
    const rows = filterRows(dict, form.tabFilter);
    const keys = new Set(rows.map((r) => r.groupKey));
    const settings = [
        state.options.includeHidden ? "hidden items included" : "hidden items left out",
        state.options.includeControls ? "subgrids and other controls included" : "fields only",
        state.options.uniqueOnly ? "first occurrence of each field only" : "",
        form.tabFilter.size ? `only ${dict.tabs.filter((t) => form.tabFilter.has(t.key)).map((t) => t.label).join(", ")}` : "",
        state.search.trim() ? `filtered by "${state.search.trim()}"` : "",
    ].filter(Boolean);
    const name = form.name.trim() || s.name;
    return {
        name,
        title: `Data dictionary: ${tableLabel(form)}, ${name}`,
        meta: [
            ["Table", tableLabel(form)],
            [s.typeLabel, s.name],
            ["Source address", form.target.raw],
            ["Environment", state.connection ? `${state.connection.name} (${state.connection.url})` : ""],
            ["Generated", new Date().toLocaleString()],
            ["Rows", String(rows.length)],
            ["Scope", settings.join("; ")],
            ["Note", "Reflects the saved design. Business rules, scripts and security roles can show, hide or lock fields when a user opens it."],
        ],
        columns: activeColumns(),
        rows,
        tabs: dict.tabs.filter((t) => keys.has(t.key)),
        choices: dict.choices.filter((c) => rows.some((r) => r.values.logicalName === c.logicalName)),
        splitByTab: state.options.splitByTab,
    };
}

/* ---------- export ---------- */

async function save(name: string, content: string | Uint8Array, filterName: string, ext: string): Promise<void> {
    const api = window.toolboxAPI as unknown as { fileSystem?: { saveFile?: Function }; utils?: { saveFile?: Function } };
    const saveFile = api.fileSystem?.saveFile ?? api.utils?.saveFile;
    if (saveFile) {
        const path = await saveFile(name, content, [{ name: filterName, extensions: [ext] }]);
        if (path) await notify("Data dictionary saved", String(path));
        return;
    }
    // Outside ToolBox (browser preview): fall back to a normal download
    const blob = new Blob([content as BlobPart]);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    a.click();
    URL.revokeObjectURL(a.href);
}

async function runExport(kind: string): Promise<void> {
    if (!state.forms.length) return;
    // Every built form goes into the export; forms with nothing left after filtering are skipped
    const ctxs = state.forms.map(exportContext).filter((c) => c.rows.length);
    if (!ctxs.length) {
        state.notices = [{ level: "warning", text: "There are no rows to export with the current filters." }];
        return render();
    }
    const total = ctxs.reduce((n, c) => n + c.rows.length, 0);
    const first = state.forms.find((f) => (f.name.trim() || f.surface.name) === ctxs[0].name) ?? state.forms[0];
    const base = ctxs.length > 1 ? `data-dictionary-${ctxs.length}-forms` : `${fileSlug(`${first.surface.etn} ${ctxs[0].name}`)}-data-dictionary`;
    try {
        if (kind === "xlsx") await save(`${base}.xlsx`, toXlsx(ctxs), "Excel workbook", "xlsx");
        else if (kind === "csv") await save(`${base}.csv`, toCsv(ctxs), "CSV", "csv");
        else if (kind === "md") await save(`${base}.md`, toMarkdown(ctxs), "Markdown", "md");
        else if (kind === "json") await save(`${base}.json`, toJson(ctxs), "JSON", "json");
        else if (kind === "copy") {
            await window.toolboxAPI.utils.copyToClipboard(toTsv(ctxs));
            await notify("Copied", `${total} rows are on the clipboard. Paste into Excel or a document.`);
        }
    } catch (error) {
        state.notices = [{ level: "error", text: `The export did not complete: ${errorText(error)}` }];
        render();
    }
}

/* ---------- rendering ---------- */

function renderNotices(): void {
    $("notices").innerHTML = state.notices.map((n) => `<div class="notice ${n.level}" role="${n.level === "error" ? "alert" : "status"}">${esc(n.text)}</div>`).join("");
}

function renderPicker(): void {
    const options = state.pickerOptions ?? [];
    const prefer = state.target?.prefer ?? "form";
    $("picker-title").textContent = `Choose a form or view of ${state.pickerTable}`;
    $("picker-hint").textContent =
        "The address names the table but not which form or view. People can see different forms depending on their security roles and the form they used last, so pick the one you want to document.";
    const group = (kind: "form" | "view", title: string) => {
        const list = options.filter((o) => o.kind === kind).sort((a, b) => Number(b.inApp) - Number(a.inApp));
        if (!list.length) return "";
        return (
            `<h3>${title}</h3>` +
            list
                .map(
                    (o) =>
                        `<button type="button" class="pick" data-pick="${esc(o.id)}"><span class="name">${esc(o.name)}</span><span class="badge">${esc(o.typeLabel)}</span>` +
                        (o.inApp ? `<span class="badge app">In this app</span>` : "") +
                        (o.isDefault ? `<span class="badge">Default view</span>` : "") +
                        `<span class="desc">${esc(o.description)}</span></button>`,
                )
                .join("")
        );
    };
    const forms = group("form", "Forms");
    const views = group("view", "Views");
    $("picker-list").innerHTML = (prefer === "view" ? views + forms : forms + views) + (state.forms.length ? `<p><button type="button" class="linklike" id="cancel-picker">Cancel</button></p>` : "");
}

function plural(n: number, one: string, many = `${one}s`): string {
    return `${n} ${n === 1 ? one : many}`;
}

function renderSummary(): void {
    const s = state.surface as SurfaceRecord;
    const dict = state.dict as Dictionary;
    const fields = dict.rows.filter((r) => r.isField).length;
    const bodyTabs = dict.tabs.filter((t) => t.key.startsWith("tab:")).length;
    const parts = [plural(fields, "field")];
    if (s.kind === "form") parts.push(plural(bodyTabs, "tab"));
    if (dict.rows.length > fields) parts.push(plural(dict.rows.length - fields, "other control"));
    $("summary").innerHTML =
        `<span class="title">${esc(s.name)}</span>` +
        `<span class="meta">${esc(s.typeLabel)} · ${esc(tableLabel())} · ${parts.join(" · ")}</span>` +
        (currentForm()?.pickerOptions ? `<button type="button" class="linklike" id="back-to-picker">Choose another form or view</button>` : "");
    const input = $<HTMLInputElement>("form-name");
    if (document.activeElement !== input) input.value = currentForm()?.name ?? "";
}

/** Strip of built forms, shown once there is more than one */
function renderForms(): void {
    const el = $("forms");
    el.hidden = state.forms.length < 2;
    el.innerHTML = state.forms
        .map(
            (f, i) =>
                `<span class="formtab" data-current="${i === state.current}"><button type="button" class="formtab-open" data-form="${i}" title="${esc(f.surface.typeLabel)}: ${esc(f.surface.name)}">${esc(f.name.trim() || f.surface.name)}</button>` +
                `<button type="button" class="formtab-remove" data-remove-form="${i}" aria-label="Remove ${esc(f.name.trim() || f.surface.name)}" title="Remove from this dictionary">×</button></span>`,
        )
        .join("");
}

function renderTabs(): void {
    const dict = state.dict as Dictionary;
    if (dict.tabs.length < 2) {
        $("tabs").innerHTML = "";
        return;
    }
    const chip = (key: string | null, label: string, count: number) =>
        `<button type="button" class="chip" data-tab="${esc(key ?? "")}" aria-pressed="${key ? state.tabFilter.has(key) : state.tabFilter.size === 0}" title="${key ? "Show or hide this tab. Several can be selected." : "Show all tabs"}">${esc(label)}<span class="n">${count}</span></button>`;
    $("tabs").innerHTML = chip(null, "All", dict.rows.length) + dict.tabs.map((t) => chip(t.key, t.label, t.count)).join("");
}

function renderOptions(): void {
    const o = state.options;
    const isForm = state.surface?.kind === "form";
    const check = (key: keyof Options, label: string) => `<label><input type="checkbox" data-option="${key}" ${o[key] ? "checked" : ""} /> ${label}</label>`;
    const languages = state.layout?.languages ?? [];
    $("options-panel").innerHTML =
        check("includeHidden", isForm ? "Include hidden tabs, sections and fields" : "Include hidden columns") +
        (isForm ? check("includeControls", "Include subgrids, quick views and web resources") : "") +
        check("uniqueOnly", "List each field once (first appearance)") +
        (isForm ? check("splitByTab", state.forms.length > 1 ? "Excel: one worksheet per tab (single form only)" : "Excel: one worksheet per tab") : "") +
        (languages.length > 1
            ? `<div class="field"><span>Label language</span><select id="language">${languages.map((l) => `<option value="${l}" ${l === state.lcid ? "selected" : ""}>${esc(LANGUAGES[l] ?? `Language ${l}`)}</option>`).join("")}</select></div>`
            : "");
}

function renderColumns(): void {
    $("columns-panel").innerHTML = COLUMNS.map(
        (c) => `<label><input type="checkbox" data-column="${c.key}" ${state.columns.has(c.key) ? "checked" : ""} /> ${esc(c.header === "#" ? "# (order)" : c.header)}${c.key === "tab" ? " (exports)" : ""}</label>`,
    ).join("");
}

function cellHtml(column: ColumnDef, value: string | number): string {
    const text = String(value ?? "");
    if (column.key === "choices") {
        const lines = text.split("\n");
        if (lines.length > 6) return `<span title="${esc(text)}">${esc(lines.slice(0, 5).join("\n"))}\n<span class="more">+ ${lines.length - 5} more</span></span>`;
    }
    return esc(text);
}

function renderGrid(): void {
    const dict = state.dict as Dictionary;
    // On screen the tab is the group heading, so its column is only used in exports
    const columns = activeColumns().filter((c) => c.key !== "tab");
    const rows = shownRows();
    const head = `<thead><tr>${columns.map((c) => `<th style="min-width:${Math.min(c.width, 30) * 7 + 20}px">${esc(c.header)}</th>`).join("")}</tr></thead>`;
    let body = "";
    let lastGroup = "";
    for (const r of rows) {
        if (r.groupKey !== lastGroup) {
            lastGroup = r.groupKey;
            const tab = dict.tabs.find((t) => t.key === r.groupKey);
            const count = rows.filter((x) => x.groupKey === r.groupKey).length;
            body += `<tr class="group"><td colspan="${columns.length}">${esc(tab?.label ?? "")}<span class="n">${count}</span></td></tr>`;
        }
        body += `<tr class="${r.hidden ? "is-hidden" : ""}">${columns.map((c) => `<td class="c-${c.key}${c.mono ? " mono" : ""}">${cellHtml(c, r.values[c.key])}</td>`).join("")}</tr>`;
    }
    if (!rows.length) {
        const why = dict.rows.length ? "No fields match the current filter." : "Nothing to list. If everything on this form is hidden, turn on “Include hidden” under Options.";
        body = `<tr class="none"><td colspan="${Math.max(columns.length, 1)}">${why}</td></tr>`;
    }
    $("grid").innerHTML = `${head}<tbody>${body}</tbody>`;

    const notes: string[] = [];
    if (rows.length !== dict.rows.length) notes.push(`Showing ${rows.length} of ${dict.rows.length} rows; exports contain the rows shown.`);
    if (dict.hiddenSkipped) notes.push(`${plural(dict.hiddenSkipped, "hidden item")} left out.`);
    if (dict.controlsSkipped) notes.push(`${plural(dict.controlsSkipped, "subgrid or other non-field control")} left out.`);
    if (dict.repeatsSkipped) notes.push(`${plural(dict.repeatsSkipped, "repeat appearance")} left out.`);
    if (dict.missingMetadata) notes.push(`${plural(dict.missingMetadata, "field")} without column details; turn on the Notes column to see why.`);
    if (dict.hiddenSkipped || dict.controlsSkipped) notes.push("Change this under Options.");
    if (state.forms.length > 1) notes.push(`Exports include all ${state.forms.length} forms, one Excel worksheet each.`);
    $("footnote").textContent = notes.join(" ");
}

function render(): void {
    renderNotices();
    const hasPicker = state.picking && !!state.pickerOptions;
    const hasResult = !hasPicker && !!(state.surface && state.dict);
    $("picker").hidden = !hasPicker;
    $("result").hidden = !hasResult;
    $("empty").hidden = hasResult || hasPicker || state.busy || state.notices.some((n) => n.level === "error");
    if (hasPicker) renderPicker();
    if (hasResult) {
        renderForms();
        renderSummary();
        renderTabs();
        renderOptions();
        renderColumns();
        renderGrid();
    }
}

/* ---------- events ---------- */

function closeMenus(except?: Element | null): void {
    document.querySelectorAll<HTMLDetailsElement>("details.menu[open]").forEach((d) => {
        if (d !== except) d.open = false;
    });
}

function wire(): void {
    $("build").addEventListener("click", () => void start("replace"));
    $("add").addEventListener("click", () => void start("add"));
    $("url").addEventListener("keydown", (e) => {
        if ((e as KeyboardEvent).key === "Enter") void start("replace");
    });
    $("form-name").addEventListener("input", (e) => {
        const form = currentForm();
        if (!form) return;
        form.name = (e.target as HTMLInputElement).value;
        renderForms(); // keeps the strip in step without touching the box being typed in
    });
    $("search").addEventListener("input", (e) => {
        state.search = (e.target as HTMLInputElement).value;
        renderGrid();
    });

    document.addEventListener("click", (e) => {
        const el = e.target as HTMLElement;
        closeMenus(el.closest("details.menu"));

        const pick = el.closest<HTMLElement>("[data-pick]");
        if (pick) {
            const option = state.pickerOptions?.find((o) => o.id === pick.dataset.pick);
            if (option) void openOption(option);
            return;
        }
        if (el.id === "back-to-picker") {
            const form = currentForm();
            if (!form?.pickerOptions) return;
            state.mode = "replace";
            state.target = form.target;
            state.pickerOptions = form.pickerOptions;
            state.pickerTable = form.pickerTable;
            state.picking = true;
            return render();
        }
        if (el.id === "cancel-picker") {
            state.picking = false;
            return render();
        }
        const remove = el.closest<HTMLElement>("[data-remove-form]");
        if (remove) {
            removeForm(Number(remove.dataset.removeForm));
            return render();
        }
        const open = el.closest<HTMLElement>("[data-form]");
        if (open) {
            activate(Number(open.dataset.form));
            return render();
        }
        const chip = el.closest<HTMLElement>("[data-tab]");
        if (chip) {
            const key = chip.dataset.tab;
            if (!key) state.tabFilter.clear();
            else if (state.tabFilter.has(key)) state.tabFilter.delete(key);
            else state.tabFilter.add(key);
            // Every tab selected is the same as no filter
            if (state.dict && state.tabFilter.size === state.dict.tabs.length) state.tabFilter.clear();
            renderTabs();
            return renderGrid();
        }
        const exportButton = el.closest<HTMLElement>("[data-export]");
        if (exportButton) {
            closeMenus();
            void runExport(exportButton.dataset.export as string);
        }
    });

    document.addEventListener("change", (e) => {
        const el = e.target as HTMLInputElement;
        if (el.dataset.option) {
            const key = el.dataset.option as keyof Options;
            state.options[key] = el.checked;
            // Bring in the columns that explain what the option just added
            if (key === "includeHidden" && el.checked) state.columns.add("visible");
            if (key === "includeControls" && el.checked) state.columns.add("notes");
            saveSettings();
            rebuild();
            render();
        } else if (el.dataset.column) {
            if (el.checked) state.columns.add(el.dataset.column);
            else state.columns.delete(el.dataset.column);
            saveSettings();
            renderGrid();
        } else if (el.id === "language") {
            state.lcid = Number(el.value);
            const form = currentForm();
            if (form) form.lcid = state.lcid;
            rebuild();
            render();
        }
    });

    document.addEventListener("keydown", (e) => {
        if (e.key === "Escape") closeMenus();
    });
}

async function initialize(): Promise<void> {
    wire();
    await Promise.all([applyTheme(), refreshConnection(), loadSettings()]);
    try {
        window.toolboxAPI.events.on((_event, payload) => {
            if (payload?.event?.startsWith("connection:")) void refreshConnection();
            if (payload?.event === "settings:updated") void applyTheme();
        });
    } catch {
        /* events are optional */
    }
    render();
    $("url").focus();
}

if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", () => void initialize());
else void initialize();
