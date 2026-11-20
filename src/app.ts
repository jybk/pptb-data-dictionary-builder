/// <reference types="@pptb/types" />
import { listSurfaces, loadForm, loadTable, loadView, type SurfaceOption, type SurfaceRecord } from "./dataverse";
import { buildDictionary, choicesForRows, COLUMNS, markShared, type ColumnDef, type Dictionary, type DictionaryRow, type TableMeta } from "./dictionary";
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
    /** Built with the current options; null until the first rebuild */
    dict: Dictionary | null;
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
    /** Columns the tool switched on itself; they follow their reason and are not remembered */
    autoColumns: new Set<string>(),
    /** Columns the user switched off while their reason applied; not switched on again this session */
    dismissedColumns: new Set<string>(),
    tabFilter: new Set<string>(),
    search: "",
    busy: false,
    exporting: false,
    /** Where to move focus after the next render, when the focused element is about to disappear */
    focusNext: null as "picker" | "result" | null,
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
        : "No connection. Select one in ToolBox, then add a form.";
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
        // Take only the known switches, and only real booleans, from what was stored
        const stored = saved?.options;
        if (stored && typeof stored === "object") {
            for (const key of Object.keys(state.options) as (keyof Options)[]) {
                if (typeof stored[key] === "boolean") state.options[key] = stored[key];
            }
        }
    } catch {
        /* defaults are fine */
    }
}

function saveSettings(): void {
    const columns = [...state.columns].filter((k) => !state.autoColumns.has(k));
    window.toolboxAPI.settings.setAll({ columns, options: state.options }).catch(() => undefined);
}

/**
 * Some columns explain something on screen, so the tool shows them while that applies:
 * Visible with hidden items, Notes with non-field controls, Table when a view reaches into
 * other tables, Also on with two or more forms. They are hidden again when the reason goes,
 * unless the user ticked them, which makes them the user's own.
 */
function syncAutoColumns(): void {
    const wanted: Record<string, boolean> = {
        visible: state.options.includeHidden,
        notes: state.options.includeControls,
        table: state.forms.some((f) => f.dict?.multipleTables),
        alsoOn: state.forms.length > 1,
    };
    for (const [key, on] of Object.entries(wanted)) {
        if (on && !state.columns.has(key) && !state.dismissedColumns.has(key)) {
            state.columns.add(key);
            state.autoColumns.add(key);
        } else if (!on) {
            state.dismissedColumns.delete(key);
            if (state.autoColumns.delete(key)) state.columns.delete(key);
        }
    }
}

/* ---------- loading ---------- */

function setBusy(busy: boolean): void {
    state.busy = busy;
    const add = $<HTMLButtonElement>("add");
    add.disabled = busy;
    add.textContent = busy ? "Reading…" : "Add form";
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
    state.focusNext = "result"; // the × that was clicked is gone
}

/**
 * Reads the address in the box and adds it as a new form ("add"). "replace" is only used
 * by "Choose another form or view", which swaps the form on screen for another one of the
 * same table. Nothing on screen changes until the read succeeds.
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

    setBusy(true);
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
    state.focusNext = "picker";
}

async function openOption(option: SurfaceOption): Promise<void> {
    if (state.busy) return;
    const runId = ++state.runId;
    state.notices = state.notices.filter((n) => n.level === "warning" && n.text.startsWith("The URL points"));
    setBusy(true);
    try {
        const surface = option.kind === "form" ? await loadForm(option.id) : await loadView(option.id);
        if (!surface) throw new Error(`"${option.name}" could not be read. It may have been deleted, or you may not have access to it.`);
        if (runId === state.runId) {
            await showSurface(surface, runId);
            state.focusNext = "result"; // the option button that was clicked is gone
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
                // Reuse what forms already in the dictionary read: saves seven requests per table,
                // and keeps one consistent snapshot of the metadata across the whole export
                const known = state.forms.find((f) => f.lcid === lcid && f.tables.has(name))?.tables.get(name);
                tables.set(name, known ?? (await loadTable(name, lcid)));
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
        dict: null,
        pickerOptions: state.pickerOptions,
        pickerTable: state.pickerTable,
    };
    if (state.mode === "add" || state.current < 0) state.forms.push(entry);
    else state.forms[state.current] = entry;
    state.picking = false;
    activate(state.mode === "add" || state.current < 0 ? state.forms.length - 1 : state.current);
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

/** The name as typed, or the form's own name when the box is empty, within Excel's 31 characters */
function typedName(form: FormEntry): string {
    return (form.name.trim() || form.surface.name).slice(0, 31);
}

/**
 * Names used on the strip, in "Also on" and in every export. Two forms can be given the same
 * name; the later one then gets " (2)" so rows and worksheets can always be told apart.
 */
function formNames(): string[] {
    const used = new Set<string>();
    return state.forms.map((form) => {
        const base = typedName(form);
        let name = base;
        for (let n = 2; used.has(name.toLowerCase()); n++) {
            const suffix = ` (${n})`;
            name = base.slice(0, 31 - suffix.length) + suffix;
        }
        used.add(name.toLowerCase());
        return name;
    });
}

function formName(form: FormEntry): string {
    return formNames()[state.forms.indexOf(form)] ?? typedName(form);
}

/** Rebuilds every form's dictionary with the current options, then marks the fields they share */
function rebuild(): void {
    for (const form of state.forms) {
        form.dict = buildDictionary(form.layout.items, form.surface.etn, form.tables, { ...state.options, lcid: form.lcid });
        const keys = new Set(form.dict.tabs.map((t) => t.key));
        for (const key of [...form.tabFilter]) if (!keys.has(key)) form.tabFilter.delete(key);
    }
    markSharedFields();
    state.dict = currentForm()?.dict ?? null;
    syncAutoColumns();
}

function markSharedFields(): void {
    const names = formNames();
    markShared(state.forms.flatMap((form, i) => (form.dict ? [{ name: names[i], dict: form.dict }] : [])));
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
    const dict = form.dict as Dictionary;
    const rows = filterRows(dict, form.tabFilter);
    const keys = new Set(rows.map((r) => r.groupKey));
    const settings = [
        state.options.includeHidden ? "hidden items included" : "hidden items left out",
        state.options.includeControls ? "subgrids and other controls included" : "fields only",
        state.options.uniqueOnly ? "first occurrence of each field only" : "",
        form.tabFilter.size ? `only ${dict.tabs.filter((t) => form.tabFilter.has(t.key)).map((t) => t.label).join(", ")}` : "",
        state.search.trim() ? `filtered by "${state.search.trim()}"` : "",
    ].filter(Boolean);
    const name = formName(form);
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
        choices: choicesForRows(dict.choices, rows),
        splitByTab: state.options.splitByTab,
    };
}

/* ---------- export ---------- */

const MIME: Record<string, string> = {
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    csv: "text/csv;charset=utf-8",
    md: "text/markdown;charset=utf-8",
    json: "application/json",
};

async function save(name: string, content: string | Uint8Array, filterName: string, ext: string): Promise<void> {
    const api = window.toolboxAPI as unknown as { fileSystem?: { saveFile?: Function }; utils?: { saveFile?: Function } };
    const saveFile = api.fileSystem?.saveFile ?? api.utils?.saveFile;
    if (saveFile) {
        const path = await saveFile(name, content, [{ name: filterName, extensions: [ext] }]);
        if (path) await notify("Data dictionary saved", String(path));
        return;
    }
    // Outside ToolBox (browser preview): fall back to a normal download
    const blob = new Blob([content as BlobPart], { type: MIME[ext] ?? "application/octet-stream" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    a.click();
    // Some browsers start the download after click() returns; revoking at once can cancel it
    setTimeout(() => URL.revokeObjectURL(a.href), 30_000);
}

async function runExport(kind: string): Promise<void> {
    // One export at a time: a double click must not open two save dialogs
    if (!state.forms.length || state.exporting) return;
    // Every built form goes into the export; forms with nothing left after filtering are skipped
    const ctxs = state.forms.map(exportContext).filter((c) => c.rows.length);
    if (!ctxs.length) {
        state.notices = [{ level: "warning", text: "There are no rows to export with the current filters." }];
        return render();
    }
    state.exporting = true;
    const total = ctxs.reduce((n, c) => n + c.rows.length, 0);
    const first = state.forms.find((f) => formName(f) === ctxs[0].name) ?? state.forms[0];
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
    } finally {
        state.exporting = false;
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
        `<span class="title" id="result-title" tabindex="-1">${esc(s.name)}</span>` +
        `<span class="meta">${esc(s.typeLabel)} · ${esc(tableLabel())} · ${parts.join(" · ")}</span>` +
        (currentForm()?.pickerOptions ? `<button type="button" class="linklike" id="back-to-picker">Choose another form or view</button>` : "");
    const input = $<HTMLInputElement>("form-name");
    if (document.activeElement !== input) input.value = currentForm()?.name ?? "";
    renderNameHint();
}

/** Says so when the typed name is already taken and the form is exported under another one */
function renderNameHint(): void {
    const form = currentForm();
    const hint = $("name-hint");
    const shown = form ? formName(form) : "";
    hint.hidden = !form || shown === typedName(form);
    hint.textContent = hint.hidden ? "" : `Another form has this name; this one is shown and exported as "${shown}".`;
}

/** Strip of built forms, shown once there is more than one */
function renderForms(): void {
    const el = $("forms");
    el.hidden = state.forms.length < 2;
    el.innerHTML = state.forms
        .map(
            (f, i) =>
                `<span class="formtab" data-current="${i === state.current}"><button type="button" class="formtab-open" data-form="${i}" title="${esc(f.surface.typeLabel)}: ${esc(f.surface.name)}">${esc(formName(f))}</button>` +
                `<button type="button" class="formtab-remove" data-remove-form="${i}" aria-label="Remove ${esc(formName(f))}" title="Remove from this dictionary">×</button></span>`,
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
    // When the element that had focus disappears, put focus somewhere sensible instead of the page
    const focus = state.focusNext;
    state.focusNext = null;
    if (focus === "picker" && hasPicker) $("picker-title").focus();
    else if (focus === "result" && hasResult) $("result-title").focus();
}

/* ---------- events ---------- */

function closeMenus(except?: Element | null): void {
    document.querySelectorAll<HTMLDetailsElement>("details.menu[open]").forEach((d) => {
        if (d !== except) d.open = false;
    });
}

function wire(): void {
    $("add").addEventListener("click", () => void start("add"));
    // The clear button only shows while there is something to clear
    const url = $<HTMLInputElement>("url");
    const syncClear = () => ($("clear-url").hidden = !url.value);
    url.addEventListener("input", syncClear);
    $("clear-url").addEventListener("click", () => {
        url.value = "";
        syncClear();
        url.focus();
    });
    syncClear();
    $("url").addEventListener("keydown", (e) => {
        if ((e as KeyboardEvent).key === "Enter") void start("add");
    });
    $("form-name").addEventListener("input", (e) => {
        const form = currentForm();
        if (!form) return;
        form.name = (e.target as HTMLInputElement).value;
        // Keep the strip and the "Also on" column in step without touching the box being typed in
        renderForms();
        renderNameHint();
        markSharedFields();
        renderGrid();
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
            state.focusNext = "picker";
            return render();
        }
        if (el.id === "cancel-picker") {
            state.picking = false;
            state.focusNext = "result";
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
            rebuild(); // also shows or hides the columns that explain the option
            saveSettings();
            render();
        } else if (el.dataset.column) {
            const key = el.dataset.column;
            // A column ticked or unticked by hand is the user's choice from now on
            state.autoColumns.delete(key);
            if (el.checked) {
                state.columns.add(key);
                state.dismissedColumns.delete(key);
            } else {
                state.columns.delete(key);
                state.dismissedColumns.add(key);
            }
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
    syncAutoColumns();
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
