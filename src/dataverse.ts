/**
 * All Dataverse reads the tool makes. Read-only: forms, views and column metadata.
 */
import type { Json, TableMeta } from "./dictionary";
import { labelText } from "./dictionary";
import { cleanGuid, isLogicalName } from "./urlParser";

export interface SurfaceRecord {
    kind: "form" | "view";
    id: string;
    name: string;
    typeLabel: string;
    etn: string;
    description: string;
    /** formxml for forms, layoutxml for views */
    xml: string;
    fetchXml: string;
}

export interface SurfaceOption {
    kind: "form" | "view";
    id: string;
    name: string;
    typeLabel: string;
    description: string;
    /** Form is part of the model-driven app named in the URL */
    inApp: boolean;
    isDefault: boolean;
}

const FORM_TYPES: Record<number, string> = { 2: "Main form", 6: "Quick view form", 7: "Quick create form", 11: "Card form" };

/** Safety stop for paging; at 5,000 rows a page this is far beyond any form or view list */
const MAX_PAGES = 50;

/**
 * Turns an absolute @odata.nextLink back into the relative query that queryData expects
 * (ToolBox prefixes every query with the connection's ".../api/data/v9.x/").
 */
export function nextLinkQuery(nextLink: string): string | null {
    const m = /\/api\/data\/v\d+(?:\.\d+)?\/(.+)$/i.exec(nextLink);
    return m ? m[1] : null;
}

/** Runs a collection query and follows @odata.nextLink, so long result sets are never cut off silently */
export async function query(odata: string): Promise<Json[]> {
    const rows: Json[] = [];
    let next: string | null = odata;
    for (let page = 0; next; page++) {
        if (page >= MAX_PAGES) throw new Error("Dataverse returned more pages than expected; the list would be incomplete.");
        const result = (await window.dataverseAPI.queryData(next)) as Json;
        if (Array.isArray(result?.value)) rows.push(...result.value);
        const link = result?.["@odata.nextLink"];
        if (typeof link !== "string") break;
        next = nextLinkQuery(link);
        if (!next) throw new Error("Dataverse returned more results than fit on one page, and the next page could not be requested.");
    }
    return rows;
}

async function querySingle(odata: string): Promise<Json> {
    return (await window.dataverseAPI.queryData(odata)) as Json;
}

/*
 * Every value placed in a query goes through one of these two checks here, not only in the
 * callers, so no future call site can build an OData injection.
 */
function requireName(etn: string): string {
    if (!isLogicalName(etn)) throw new Error(`"${etn}" is not a valid table logical name.`);
    return etn;
}

function requireGuid(id: string): string {
    const guid = cleanGuid(id);
    if (!guid) throw new Error(`"${id}" is not a valid id.`);
    return guid;
}

export async function loadForm(formId: string): Promise<SurfaceRecord | null> {
    const id = requireGuid(formId);
    const rows = await query(`systemforms?$select=formid,name,type,objecttypecode,formxml,description&$filter=formid eq ${id}`);
    const f = rows[0];
    if (!f) return null;
    return {
        kind: "form",
        id: f.formid,
        name: f.name ?? "(unnamed form)",
        typeLabel: FORM_TYPES[f.type] ?? "Form",
        etn: String(f.objecttypecode ?? "").toLowerCase(),
        description: f.description ?? "",
        xml: f.formxml ?? "",
        fetchXml: "",
    };
}

export async function loadView(viewId: string): Promise<SurfaceRecord | null> {
    const id = requireGuid(viewId);
    const system = await query(`savedqueries?$select=savedqueryid,name,returnedtypecode,layoutxml,fetchxml,description&$filter=savedqueryid eq ${id}`);
    let v = system[0];
    let typeLabel = "System view";
    if (!v) {
        // Personal views live in a different table and are only visible to their owner or people they are shared with
        const personal = await query(`userqueries?$select=userqueryid,name,returnedtypecode,layoutxml,fetchxml,description&$filter=userqueryid eq ${id}`).catch(() => []);
        v = personal[0];
        typeLabel = "Personal view";
    }
    if (!v) return null;
    return {
        kind: "view",
        id: v.savedqueryid ?? v.userqueryid,
        name: v.name ?? "(unnamed view)",
        typeLabel,
        etn: String(v.returnedtypecode ?? "").toLowerCase(),
        description: v.description ?? "",
        xml: v.layoutxml ?? "",
        fetchXml: v.fetchxml ?? "",
    };
}

/** Form ids that a model-driven app explicitly includes. Empty set = unknown or the app includes all forms. */
async function formsInApp(appId: string): Promise<Set<string>> {
    try {
        const apps = await query(`appmodules?$select=appmoduleidunique&$filter=appmoduleid eq ${requireGuid(appId)}`);
        const unique = cleanGuid(apps[0]?.appmoduleidunique);
        if (!unique) return new Set();
        const components = await query(`appmodulecomponents?$select=objectid&$filter=_appmoduleidunique_value eq ${unique} and componenttype eq 60`);
        return new Set(components.map((c) => String(c.objectid).toLowerCase()));
    } catch {
        return new Set(); // best effort only
    }
}

export async function listSurfaces(etn: string, appId: string | undefined): Promise<SurfaceOption[]> {
    const name = requireName(etn);
    const [forms, views, appForms] = await Promise.all([
        query(`systemforms?$select=formid,name,type,description&$filter=objecttypecode eq '${name}' and (type eq 2 or type eq 6 or type eq 7) and formactivationstate eq 1&$orderby=type asc,name asc`),
        query(`savedqueries?$select=savedqueryid,name,description,isdefault&$filter=returnedtypecode eq '${name}' and querytype eq 0 and statecode eq 0&$orderby=name asc`).catch(() => []),
        appId ? formsInApp(appId) : Promise.resolve(new Set<string>()),
    ]);
    const options: SurfaceOption[] = [];
    for (const f of forms) {
        options.push({
            kind: "form",
            id: f.formid,
            name: f.name ?? "(unnamed form)",
            typeLabel: FORM_TYPES[f.type] ?? "Form",
            description: f.description ?? "",
            inApp: appForms.has(String(f.formid).toLowerCase()),
            isDefault: false,
        });
    }
    for (const v of views) {
        options.push({ kind: "view", id: v.savedqueryid, name: v.name ?? "(unnamed view)", typeLabel: "System view", description: v.description ?? "", inApp: false, isDefault: v.isdefault === true });
    }
    return options;
}

/** Attribute types whose metadata carries an option set; each needs its own cast to expand it */
const OPTION_SET_TYPES = ["Picklist", "MultiSelectPicklist", "State", "Status", "Boolean", "EntityName"];

export async function loadTable(etn: string, lcid: number): Promise<TableMeta> {
    const name = requireName(etn);
    const base = `EntityDefinitions(LogicalName='${name}')`;
    const [entity, attributes, ...optionSetResults] = await Promise.all([
        querySingle(`${base}?$select=LogicalName,DisplayName`),
        // No $select on purpose: type-specific properties (MaxLength, Targets, Precision, ...) only come back unfiltered
        query(`${base}/Attributes`),
        ...OPTION_SET_TYPES.map((t) => query(`${base}/Attributes/Microsoft.Dynamics.CRM.${t}AttributeMetadata?$select=LogicalName&$expand=OptionSet`).catch(() => [] as Json[])),
    ]);
    const meta: TableMeta = { logicalName: name, displayName: labelText(entity?.DisplayName, lcid) || name, attributes: new Map(), optionSets: new Map() };
    for (const a of attributes) if (a.LogicalName) meta.attributes.set(a.LogicalName, a);
    for (const list of optionSetResults) for (const a of list) if (a.LogicalName && a.OptionSet) meta.optionSets.set(a.LogicalName, a.OptionSet);
    return meta;
}
