/**
 * Combines the ordered UI items with Dataverse column metadata into data dictionary rows.
 */
import type { Labels, UiItem } from "./layoutParser";

// Web API metadata is consumed as loose JSON; only the properties read below matter.
export type Json = Record<string, any>;

export interface TableMeta {
    logicalName: string;
    displayName: string;
    /** Attribute metadata by logical name */
    attributes: Map<string, Json>;
    /** OptionSet metadata by attribute logical name (choice, status, yes/no columns) */
    optionSets: Map<string, Json>;
}

export interface BuildOptions {
    includeHidden: boolean;
    includeControls: boolean;
    uniqueOnly: boolean;
    lcid: number;
}

export interface ColumnDef {
    key: string;
    header: string;
    /** Shown by default */
    on: boolean;
    /** Approximate width in characters, used for Excel and the on-screen table */
    width: number;
    mono?: boolean;
}

export const COLUMNS: ColumnDef[] = [
    { key: "order", header: "#", on: true, width: 5 },
    { key: "tab", header: "Tab", on: true, width: 18 },
    { key: "section", header: "Section", on: true, width: 20 },
    { key: "label", header: "Label", on: true, width: 26 },
    { key: "displayName", header: "Display Name", on: false, width: 26 },
    { key: "logicalName", header: "Logical Name", on: true, width: 26, mono: true },
    { key: "schemaName", header: "Schema Name", on: false, width: 26, mono: true },
    { key: "table", header: "Table", on: false, width: 18, mono: true },
    { key: "alsoOn", header: "Also on", on: false, width: 24 },
    { key: "dataType", header: "Data Type", on: true, width: 18 },
    { key: "format", header: "Format", on: true, width: 16 },
    { key: "size", header: "Length / Range", on: true, width: 24 },
    { key: "required", header: "Required", on: true, width: 14 },
    { key: "readOnly", header: "Read-only", on: true, width: 10 },
    { key: "visible", header: "Visible", on: false, width: 9 },
    { key: "choices", header: "Allowed Values", on: true, width: 34 },
    { key: "defaultValue", header: "Default", on: true, width: 16 },
    { key: "targets", header: "Lookup Targets", on: true, width: 22, mono: true },
    { key: "description", header: "Description", on: true, width: 44 },
    { key: "custom", header: "Custom", on: true, width: 8 },
    { key: "source", header: "Source", on: false, width: 12 },
    { key: "searchable", header: "Searchable", on: false, width: 11 },
    { key: "audited", header: "Audited", on: false, width: 9 },
    { key: "secured", header: "Column Security", on: false, width: 15 },
    { key: "controlType", header: "Control", on: false, width: 14 },
    { key: "notes", header: "Notes", on: false, width: 34 },
];

export interface DictionaryRow {
    groupKey: string;
    values: Record<string, string | number>;
    hidden: boolean;
    isField: boolean;
}

export interface ChoiceRow {
    table: string;
    logicalName: string;
    label: string;
    value: number | string;
    choice: string;
}

export interface TabInfo {
    key: string;
    label: string;
    count: number;
}

export interface Dictionary {
    rows: DictionaryRow[];
    choices: ChoiceRow[];
    tabs: TabInfo[];
    /** Items left out because they are hidden by default */
    hiddenSkipped: number;
    /** Subgrids, web resources etc. left out */
    controlsSkipped: number;
    /** Repeat occurrences of a field left out (uniqueOnly) */
    repeatsSkipped: number;
    missingMetadata: number;
    multipleTables: boolean;
}

/** Label in the chosen language; failing that English (1033), then whichever language there is */
function pick(labels: Labels, lcid: number): string {
    if (labels[lcid]) return labels[lcid];
    if (labels[1033]) return labels[1033];
    // Numeric keys iterate in ascending order, so this is the lowest language code, not the base language
    const first = Object.keys(labels)[0];
    return first ? labels[Number(first)] : "";
}

/** Label object from the Web API -> text in the requested language, else the caller's language */
export function labelText(label: Json | null | undefined, lcid: number): string {
    if (!label) return "";
    const localized: Json[] = label.LocalizedLabels ?? [];
    const match = localized.find((l) => l.LanguageCode === lcid);
    return (match?.Label ?? label.UserLocalizedLabel?.Label ?? localized[0]?.Label ?? "").trim();
}

const TYPE_NAMES: Record<string, string> = {
    StringType: "Text",
    MemoType: "Multiline Text",
    IntegerType: "Whole Number",
    BigIntType: "Big Integer",
    DecimalType: "Decimal Number",
    DoubleType: "Floating Point Number",
    MoneyType: "Currency",
    BooleanType: "Yes/No",
    DateTimeType: "Date and Time",
    PicklistType: "Choice",
    MultiSelectPicklistType: "Choices (multi-select)",
    StateType: "Status",
    StatusType: "Status Reason",
    LookupType: "Lookup",
    CustomerType: "Customer",
    OwnerType: "Owner",
    UniqueidentifierType: "Unique Identifier",
    ImageType: "Image",
    FileType: "File",
    PartyListType: "Activity Party",
    EntityNameType: "Table Name",
    VirtualType: "Virtual",
    ManagedPropertyType: "Managed Property",
    CalendarRulesType: "Calendar Rules",
};

const REQUIRED: Record<string, string> = {
    None: "Optional",
    Recommended: "Recommended",
    ApplicationRequired: "Required",
    SystemRequired: "System Required",
};

const SOURCE: Record<number, string> = { 0: "Simple", 1: "Calculated", 2: "Rollup", 3: "Formula" };

function typeKey(a: Json): string {
    return a.AttributeTypeName?.Value ?? (a.AttributeType ? `${a.AttributeType}Type` : "");
}

function dataType(a: Json): string {
    const key = typeKey(a);
    if (key === "DateTimeType" && (a.Format === "DateOnly" || a.DateTimeBehavior?.Value === "DateOnly")) return "Date Only";
    return TYPE_NAMES[key] ?? key.replace(/Type$/, "");
}

function format(a: Json): string {
    const key = typeKey(a);
    if (key === "StringType" || key === "MemoType") {
        const f = a.FormatName?.Value ?? a.Format ?? "";
        return f === "Text" || f === "TextArea" ? "" : f;
    }
    if (key === "IntegerType") return a.Format && a.Format !== "None" ? a.Format : "";
    if (key === "DateTimeType") return a.DateTimeBehavior?.Value ?? "";
    return "";
}

function num(v: unknown): string {
    return typeof v === "number" ? v.toLocaleString("en-US", { maximumFractionDigits: 10 }) : "";
}

function size(a: Json): string {
    const key = typeKey(a);
    if (key === "StringType" || key === "MemoType") return a.MaxLength != null ? `Max ${num(a.MaxLength)} characters` : "";
    if (key === "IntegerType" || key === "BigIntType") return a.MinValue != null && a.MaxValue != null ? `${num(a.MinValue)} to ${num(a.MaxValue)}` : "";
    if (key === "DecimalType" || key === "DoubleType" || key === "MoneyType") {
        const range = a.MinValue != null && a.MaxValue != null ? `${num(a.MinValue)} to ${num(a.MaxValue)}` : "";
        const precision = a.Precision != null ? `${a.Precision} decimals` : "";
        return [range, precision].filter(Boolean).join("; ");
    }
    if (key === "FileType" || key === "ImageType") return a.MaxSizeInKB != null ? `Max ${num(a.MaxSizeInKB)} KB` : "";
    return "";
}

function yesNo(v: unknown): string {
    return v === true ? "Yes" : v === false ? "No" : "";
}

interface Option {
    value: number;
    label: string;
}

function options(optionSet: Json | undefined, lcid: number): Option[] {
    if (!optionSet) return [];
    if (optionSet.TrueOption || optionSet.FalseOption) {
        return [optionSet.TrueOption, optionSet.FalseOption].filter(Boolean).map((o: Json) => ({ value: o.Value, label: labelText(o.Label, lcid) }));
    }
    return (optionSet.Options ?? []).map((o: Json) => ({ value: o.Value, label: labelText(o.Label, lcid) }));
}

function defaultValue(a: Json, opts: Option[]): string {
    const key = typeKey(a);
    let value: unknown;
    if (key === "BooleanType") value = a.DefaultValue === true ? 1 : a.DefaultValue === false ? 0 : undefined;
    else if (key === "PicklistType" || key === "MultiSelectPicklistType" || key === "StateType" || key === "StatusType") value = a.DefaultFormValue;
    if (typeof value !== "number" || value === -1) return "";
    return opts.find((o) => o.value === value)?.label ?? String(value);
}

export function buildDictionary(items: UiItem[], ownTable: string, tables: Map<string, TableMeta>, options_: BuildOptions): Dictionary {
    const { lcid } = options_;
    const rows: DictionaryRow[] = [];
    const choices: ChoiceRow[] = [];
    const tabs = new Map<string, TabInfo>();
    const seen = new Set<string>();
    const choiceSeen = new Set<string>();
    const usedTables = new Set<string>();
    let hiddenSkipped = 0;
    let controlsSkipped = 0;
    let repeatsSkipped = 0;
    let missingMetadata = 0;

    for (const item of items) {
        const isField = item.controlType === "Field" && !!item.fieldName;
        const hidden = !(item.tabVisible && item.sectionVisible && item.cellVisible);
        if (!isField && !options_.includeControls) {
            controlsSkipped++;
            continue;
        }
        if (hidden && !options_.includeHidden) {
            hiddenSkipped++;
            continue;
        }
        const tableName = item.table || ownTable;
        const fieldKey = `${tableName}.${item.fieldName}`;
        if (isField && options_.uniqueOnly) {
            if (seen.has(fieldKey)) {
                repeatsSkipped++;
                continue;
            }
            seen.add(fieldKey);
        }

        const tabLabel = item.area === "Body" ? pick(item.tabLabels, lcid) || item.tabName : item.area === "View" ? "View columns" : item.area;
        const table = tables.get(tableName);
        const a = isField ? table?.attributes.get(item.fieldName as string) : undefined;
        const optionSet = isField ? table?.optionSets.get(item.fieldName as string) : undefined;
        const opts = options(optionSet, lcid);
        const displayName = a ? labelText(a.DisplayName, lcid) : "";
        const label = pick(item.cellLabels, lcid) || displayName || item.fieldName || item.controlId;
        const notes: string[] = [];
        if (item.detail) notes.push(item.detail);
        if (isField && !a) {
            missingMetadata++;
            notes.push(tableName.startsWith("?") ? `Column of a linked table (alias "${tableName.slice(1)}") that could not be resolved` : "No column metadata found for this field");
        }
        if (hidden) notes.push("Hidden by default");
        usedTables.add(tableName);

        const values: Record<string, string | number> = {
            order: rows.length + 1,
            tab: tabLabel,
            section: pick(item.sectionLabels, lcid) || item.sectionName,
            label,
            displayName,
            logicalName: item.fieldName ?? "",
            schemaName: a?.SchemaName ?? "",
            table: isField ? tableName.replace(/^\?/, "") : "",
            dataType: a ? dataType(a) : isField ? "" : item.controlType,
            format: a ? format(a) : "",
            size: a ? size(a) : "",
            required: a ? (REQUIRED[a.RequiredLevel?.Value] ?? "") : "",
            readOnly: !isField || item.area === "View" ? "" : yesNo(item.disabled || (!!a && a.IsValidForUpdate === false && a.IsValidForCreate === false)),
            visible: hidden ? "No" : "Yes",
            choices: opts.map((o) => `${o.value} = ${o.label}`).join("\n"),
            defaultValue: a ? defaultValue(a, opts) : "",
            targets: a && Array.isArray(a.Targets) ? a.Targets.join(", ") : "",
            description: a ? labelText(a.Description, lcid) : "",
            custom: a ? yesNo(a.IsCustomAttribute) : "",
            source: a && typeof a.SourceType === "number" ? (SOURCE[a.SourceType] ?? "") : "",
            searchable: a ? yesNo(a.IsValidForAdvancedFind?.Value) : "",
            audited: a ? yesNo(a.IsAuditEnabled?.Value) : "",
            secured: a ? yesNo(a.IsSecured) : "",
            // Filled in by the app when several forms are in the dictionary
            alsoOn: "",
            controlType: item.controlType,
            notes: notes.join("; "),
        };
        rows.push({ groupKey: item.groupKey, values, hidden, isField });

        const tab = tabs.get(item.groupKey) ?? { key: item.groupKey, label: tabLabel, count: 0 };
        tab.count++;
        tabs.set(item.groupKey, tab);

        if (opts.length && !choiceSeen.has(fieldKey)) {
            choiceSeen.add(fieldKey);
            for (const o of opts) choices.push({ table: tableName, logicalName: item.fieldName as string, label, value: o.value, choice: o.label });
        }
    }

    return { rows, choices, tabs: [...tabs.values()], hiddenSkipped, controlsSkipped, repeatsSkipped, missingMetadata, multipleTables: usedTables.size > 1 };
}

/**
 * Marks fields that appear in more than one dictionary: each field row gets the names of
 * the other dictionaries that contain the same column of the same table.
 */
export function markShared(dictionaries: { name: string; dict: Dictionary }[]): void {
    const key = (r: DictionaryRow) => `${r.values.table}.${r.values.logicalName}`;
    const where = new Map<string, Set<number>>();
    dictionaries.forEach(({ dict }, i) => {
        for (const r of dict.rows) {
            if (!r.isField) continue;
            const set = where.get(key(r)) ?? new Set<number>();
            set.add(i);
            where.set(key(r), set);
        }
    });
    dictionaries.forEach(({ dict }, i) => {
        for (const r of dict.rows) {
            const others = r.isField ? [...(where.get(key(r)) ?? [])].filter((j) => j !== i) : [];
            r.values.alsoOn = others.map((j) => dictionaries[j].name).join(", ");
        }
    });
}

/**
 * Allowed values for the rows that are shown. Matches on table and column, because the same
 * column name (statuscode, for one) exists on every table with a different set of choices.
 */
export function choicesForRows(choices: ChoiceRow[], rows: DictionaryRow[]): ChoiceRow[] {
    const shown = new Set(rows.filter((r) => r.isField).map((r) => `${r.values.table}.${r.values.logicalName}`));
    return choices.filter((c) => shown.has(`${c.table.replace(/^\?/, "")}.${c.logicalName}`));
}
