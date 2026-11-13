/**
 * Turns whatever the user pasted (a model-driven app URL, a maker-portal designer URL,
 * a bare table logical name or a bare form/view GUID) into a target the tool can load.
 */

export type TargetKind = "form" | "view" | "table" | "id" | "unknown";

export interface ParsedTarget {
    kind: TargetKind;
    /** Table logical name, when the input names one */
    etn?: string;
    formId?: string;
    viewId?: string;
    /** Model-driven app id (appid=...), used to highlight the forms that app contains */
    appId?: string;
    /** Bare GUID with no hint whether it is a form or a view */
    anyId?: string;
    /** Host of the pasted URL, lowercased */
    host?: string;
    /** True for make.powerapps.com style URLs, whose host is not the environment host */
    isMakerPortal?: boolean;
    /** For kind "table": which list the URL pointed at */
    prefer?: "form" | "view";
    raw: string;
}

const GUID = "[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}";
const GUID_ONLY = new RegExp(`^\\{?(${GUID})\\}?$`, "i");
const LOGICAL_NAME = /^[a-z][a-z0-9_]*$/;

export function cleanGuid(value: string | null | undefined): string | undefined {
    if (!value) return undefined;
    let v = value.trim();
    try {
        v = decodeURIComponent(v);
    } catch {
        /* keep as is */
    }
    const m = GUID_ONLY.exec(v.trim());
    return m ? m[1].toLowerCase() : undefined;
}

export function isLogicalName(value: string | null | undefined): value is string {
    return !!value && LOGICAL_NAME.test(value);
}

function collectParams(search: string, into: Map<string, string>): void {
    const params = new URLSearchParams(search.replace(/^[?#]/, ""));
    params.forEach((value, key) => {
        const k = key.toLowerCase();
        if (!into.has(k)) into.set(k, value);
    });
}

export function parseTarget(input: string): ParsedTarget {
    const raw = (input ?? "").trim();
    const result: ParsedTarget = { kind: "unknown", raw };
    if (!raw) return result;

    const bareId = cleanGuid(raw);
    if (bareId) return { ...result, kind: "id", anyId: bareId };

    if (isLogicalName(raw.toLowerCase()) && !raw.includes(".")) {
        return { ...result, kind: "table", etn: raw.toLowerCase(), prefer: "form" };
    }

    let url: URL;
    try {
        url = new URL(/^[a-z][a-z0-9+.-]*:\/\//i.test(raw) ? raw : `https://${raw}`);
    } catch {
        return result;
    }
    if (!url.hostname.includes(".")) return result;

    result.host = url.hostname.toLowerCase();
    result.isMakerPortal = /(^|\.)make\.(preview\.)?powerapps\.|(^|\.)make\.powerautomate\./.test(result.host) || /powerapps\.(com|us|cn)$/.test(result.host);

    const params = new Map<string, string>();
    collectParams(url.search, params);
    if (url.hash.includes("=")) collectParams(url.hash, params);
    // Classic URLs tuck formid/viewid into an encoded "extraqs" parameter
    const extraqs = params.get("extraqs");
    if (extraqs) collectParams(extraqs, params);

    const etn = params.get("etn")?.toLowerCase();
    if (isLogicalName(etn)) result.etn = etn;
    result.formId = cleanGuid(params.get("formid"));
    result.viewId = cleanGuid(params.get("viewid"));
    result.appId = cleanGuid(params.get("appid"));
    const pageType = (params.get("pagetype") ?? "").toLowerCase();

    // Maker portal designers: .../entity/account/form/edit/<guid>, .../entity/account/view/<guid>
    const path = url.pathname;
    const formInPath = new RegExp(`/forms?/(?:edit/)?(${GUID})`, "i").exec(path);
    const viewInPath = new RegExp(`/views?/(?:edit/)?(${GUID})`, "i").exec(path);
    const tableInPath = new RegExp(`/entit(?:y|ies)/(?:${GUID}/)?([a-z][a-z0-9_]*)`, "i").exec(path);
    if (!result.formId && formInPath) result.formId = formInPath[1].toLowerCase();
    if (!result.viewId && viewInPath) result.viewId = viewInPath[1].toLowerCase();
    if (!result.etn && tableInPath && isLogicalName(tableInPath[1].toLowerCase())) result.etn = tableInPath[1].toLowerCase();

    if (pageType === "entitylist" && result.viewId) result.kind = "view";
    else if (result.formId) result.kind = "form";
    else if (result.viewId) result.kind = "view";
    else if (result.etn) {
        result.kind = "table";
        result.prefer = pageType === "entitylist" || /\/views?\/?$/i.test(path) ? "view" : "form";
    }
    return result;
}

/** Compares the pasted URL's host with the active connection; returns a warning or null. */
export function hostMismatch(target: ParsedTarget, connectionUrl: string | undefined): string | null {
    if (!target.host || target.isMakerPortal || !connectionUrl) return null;
    let connHost: string;
    try {
        connHost = new URL(connectionUrl).hostname.toLowerCase();
    } catch {
        return null;
    }
    if (connHost === target.host) return null;
    return `The URL points to ${target.host}, but the active connection is ${connHost}. The dictionary is built from the active connection.`;
}
