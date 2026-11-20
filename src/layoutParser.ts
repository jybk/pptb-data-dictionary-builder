/**
 * Reads form XML and view layout XML into a flat, ordered list of UI items.
 * Only uses DOM Level 1 calls (childNodes, getAttribute, getElementsByTagName) so the
 * same code runs in the browser and under @xmldom/xmldom in tests.
 */

export type Labels = Record<number, string>;

export type ControlType = "Field" | "Subgrid" | "Quick view" | "Web resource" | "IFrame" | "Timeline" | "Other control";

export interface UiItem {
    area: "Header" | "Body" | "Footer" | "View";
    /** Stable key for grouping: "header", "tab:0", "footer", "view" */
    groupKey: string;
    tabName: string;
    tabLabels: Labels;
    tabVisible: boolean;
    sectionName: string;
    sectionLabels: Labels;
    sectionVisible: boolean;
    cellLabels: Labels;
    cellVisible: boolean;
    controlId: string;
    controlType: ControlType;
    /** Attribute logical name; null for subgrids, web resources and the like */
    fieldName: string | null;
    /** Table the field belongs to. Empty string = the surface's own table */
    table: string;
    disabled: boolean;
    /** Extra info for non-field controls (e.g. subgrid target table) or view columns (width) */
    detail: string;
}

export interface ParsedLayout {
    items: UiItem[];
    /** Language codes found on form labels, most frequent first */
    languages: number[];
}

type XmlParser = (xml: string) => Document;

const CLASS_IDS: Record<string, ControlType> = {
    "e7a81278-8635-4d9e-8d4d-59480b391c5b": "Subgrid",
    "5c5600e0-1d6e-4205-a272-be80da87fd42": "Quick view",
    "9fdf5f91-88b1-47f4-ad53-c11efc01a01d": "Web resource",
    "fd2a7985-3187-444e-908d-6624b21f69c0": "IFrame",
    "06375649-c143-495e-a496-c962e5b4488e": "Timeline",
};

function kids(el: Element | null | undefined, name: string): Element[] {
    const out: Element[] = [];
    if (!el) return out;
    const nodes = el.childNodes;
    for (let i = 0; i < nodes.length; i++) {
        const n = nodes[i] as Element;
        if (n.nodeType === 1 && n.nodeName.toLowerCase() === name) out.push(n);
    }
    return out;
}

function kid(el: Element | null | undefined, name: string): Element | undefined {
    return kids(el, name)[0];
}

function attr(el: Element | undefined, name: string): string {
    return el?.getAttribute(name) ?? "";
}

/** Form XML treats a missing visible attribute as visible */
function isVisible(el: Element | undefined): boolean {
    return attr(el, "visible").toLowerCase() !== "false";
}

function readLabels(el: Element | undefined, counts?: Map<number, number>): Labels {
    const labels: Labels = {};
    for (const label of kids(kid(el, "labels"), "label")) {
        const lcid = parseInt(attr(label, "languagecode"), 10);
        const text = attr(label, "description");
        if (!Number.isFinite(lcid)) continue;
        labels[lcid] = text;
        counts?.set(lcid, (counts.get(lcid) ?? 0) + 1);
    }
    return labels;
}

function textOf(el: Element | undefined, tag: string): string {
    const found = el?.getElementsByTagName(tag);
    return found && found.length ? (found[0].textContent ?? "").trim() : "";
}

function controlDetail(type: ControlType, control: Element): string {
    const params = kid(control, "parameters");
    if (!params) return "";
    if (type === "Subgrid") {
        const target = textOf(params, "TargetEntityType");
        const rel = textOf(params, "RelationshipName");
        return [target && `Rows from ${target}`, rel && `via ${rel}`].filter(Boolean).join(" ");
    }
    if (type === "Web resource") return textOf(params, "Url");
    if (type === "IFrame") return textOf(params, "Url");
    return "";
}

/**
 * Parses XML and insists on a clean result. Browsers do not throw on malformed XML: they keep
 * whatever parsed and insert a <parsererror> element, which would silently truncate the list.
 */
function parseStrict(xml: string, parse: XmlParser, root: string, what: string): Element {
    let doc: Document;
    try {
        doc = parse(xml);
    } catch {
        throw new Error(`This ${what}'s definition could not be read (malformed XML).`);
    }
    const el = doc?.documentElement;
    if (!el || doc.getElementsByTagName("parsererror").length) {
        throw new Error(`This ${what}'s definition could not be read (malformed XML).`);
    }
    if (el.nodeName.toLowerCase() !== root) {
        throw new Error(`This ${what}'s definition could not be read (unexpected ${what} XML).`);
    }
    return el;
}

export function parseFormXml(formXml: string, parse: XmlParser): ParsedLayout {
    const form = parseStrict(formXml, parse, "form", "form");
    const items: UiItem[] = [];
    const counts = new Map<number, number>();

    const readRows = (container: Element | undefined, base: Omit<UiItem, "cellLabels" | "cellVisible" | "controlId" | "controlType" | "fieldName" | "table" | "disabled" | "detail">) => {
        for (const row of kids(kid(container, "rows"), "row")) {
            for (const cell of kids(row, "cell")) {
                const control = kid(cell, "control");
                if (!control) continue; // spacer cell
                const field = attr(control, "datafieldname").trim().toLowerCase();
                const classId = attr(control, "classid").replace(/[{}]/g, "").toLowerCase();
                const type: ControlType = field ? "Field" : (CLASS_IDS[classId] ?? "Other control");
                items.push({
                    ...base,
                    cellLabels: readLabels(cell, counts),
                    cellVisible: isVisible(cell) && isVisible(control),
                    controlId: attr(control, "id"),
                    controlType: type,
                    fieldName: field || null,
                    table: "",
                    disabled: attr(control, "disabled").toLowerCase() === "true",
                    detail: field ? "" : controlDetail(type, control),
                });
            }
        }
    };

    const fixedArea = (area: "Header" | "Footer") => ({
        area,
        groupKey: area.toLowerCase(),
        tabName: area.toLowerCase(),
        tabLabels: {} as Labels,
        tabVisible: true,
        sectionName: "",
        sectionLabels: {} as Labels,
        sectionVisible: true,
    });

    // Order of appearance on screen: header, then tabs left to right, then footer
    const header = kid(form, "header");
    if (header) readRows(header, fixedArea("Header"));

    kids(kid(form, "tabs"), "tab").forEach((tab, tabIndex) => {
        const tabLabels = readLabels(tab, counts);
        for (const column of kids(kid(tab, "columns"), "column")) {
            for (const section of kids(kid(column, "sections"), "section")) {
                readRows(section, {
                    area: "Body",
                    groupKey: `tab:${tabIndex}`,
                    tabName: attr(tab, "name") || `tab_${tabIndex + 1}`,
                    tabLabels,
                    tabVisible: isVisible(tab),
                    sectionName: attr(section, "name"),
                    sectionLabels: readLabels(section, counts),
                    sectionVisible: isVisible(section),
                });
            }
        }
    });

    const footer = kid(form, "footer");
    if (footer) readRows(footer, fixedArea("Footer"));

    const languages = [...counts.entries()].sort((a, b) => b[1] - a[1]).map(([lcid]) => lcid);
    return { items, languages };
}

/** Maps link-entity aliases in a view's FetchXML to the linked table's logical name */
export function readFetchAliases(fetchXml: string, parse: XmlParser): Record<string, string> {
    const aliases: Record<string, string> = {};
    if (!fetchXml) return aliases;
    let links: HTMLCollectionOf<Element>;
    try {
        links = parse(fetchXml).getElementsByTagName("link-entity");
    } catch {
        // Unreadable FetchXML only costs the linked-table lookup: those columns show as "?alias"
        return aliases;
    }
    for (let i = 0; i < links.length; i++) {
        const alias = links[i].getAttribute("alias");
        const name = links[i].getAttribute("name");
        if (alias && name) aliases[alias.toLowerCase()] = name.toLowerCase();
    }
    return aliases;
}

export function parseViewLayout(layoutXml: string, fetchXml: string, parse: XmlParser): ParsedLayout {
    const grid = parseStrict(layoutXml, parse, "grid", "view");
    const cells = grid.getElementsByTagName("cell");
    const aliases = readFetchAliases(fetchXml, parse);
    const items: UiItem[] = [];
    for (let i = 0; i < cells.length; i++) {
        const cell = cells[i];
        const name = (cell.getAttribute("name") ?? "").trim().toLowerCase();
        if (!name) continue;
        const dot = name.indexOf(".");
        const alias = dot > 0 ? name.slice(0, dot) : "";
        const width = cell.getAttribute("width");
        items.push({
            area: "View",
            groupKey: "view",
            tabName: "view",
            tabLabels: {},
            tabVisible: true,
            sectionName: "",
            sectionLabels: {},
            sectionVisible: true,
            cellLabels: {},
            cellVisible: (cell.getAttribute("ishidden") ?? "0") !== "1",
            controlId: name,
            controlType: "Field",
            fieldName: dot > 0 ? name.slice(dot + 1) : name,
            // Unknown alias: keep it visible to the user instead of guessing a table
            table: alias ? (aliases[alias] ?? `?${alias}`) : "",
            disabled: false,
            detail: width ? `Column width ${width}px` : "",
        });
    }
    return { items, languages: [] };
}
