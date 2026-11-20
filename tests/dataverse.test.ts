import assert from "node:assert/strict";
import { beforeEach, test } from "node:test";
import { listSurfaces, loadForm, loadTable, loadView, nextLinkQuery, query } from "../src/dataverse";

/** Stand-in for ToolBox's dataverseAPI: records every query and answers from a handler */
let calls: string[] = [];
let handler: (q: string) => unknown = () => ({ value: [] });
(globalThis as unknown as { window: unknown }).window = {
    dataverseAPI: {
        queryData: async (q: string) => {
            calls.push(q);
            return handler(q);
        },
    },
};
beforeEach(() => {
    calls = [];
    handler = () => ({ value: [] });
});

const BASE = "https://org.crm4.dynamics.com/api/data/v9.2/";
const ID = "8448b78f-8f42-454e-8e2a-f8196b0419af";

test("nextLink is turned back into the relative query ToolBox expects", () => {
    assert.equal(nextLinkQuery(`${BASE}systemforms?$select=name&$skiptoken=%3Ccookie%3E`), "systemforms?$select=name&$skiptoken=%3Ccookie%3E");
    assert.equal(nextLinkQuery("https://org.crm.dynamics.com/api/data/v9.0/savedqueries?$skiptoken=x"), "savedqueries?$skiptoken=x");
    assert.equal(nextLinkQuery("https://example.org/other"), null);
});

test("query follows @odata.nextLink until the last page", async () => {
    handler = (q) =>
        q.includes("skiptoken=2")
            ? { value: [{ n: 3 }] }
            : q.includes("skiptoken=1")
              ? { value: [{ n: 2 }], "@odata.nextLink": `${BASE}systemforms?$select=n&$skiptoken=2` }
              : { value: [{ n: 1 }], "@odata.nextLink": `${BASE}systemforms?$select=n&$skiptoken=1` };
    const rows = await query("systemforms?$select=n");
    assert.deepEqual(rows.map((r) => r.n), [1, 2, 3]);
    assert.deepEqual(calls, ["systemforms?$select=n", "systemforms?$select=n&$skiptoken=1", "systemforms?$select=n&$skiptoken=2"]);
});

test("query refuses to return a cut-off list", async () => {
    handler = () => ({ value: [{}], "@odata.nextLink": "https://example.org/unexpected" });
    await assert.rejects(query("systemforms"), /next page could not be requested/);
    handler = (q) => ({ value: [{}], "@odata.nextLink": `${BASE}${q}x` });
    await assert.rejects(query("systemforms"), /more pages than expected/);
    assert.equal(calls.length, 1 + 50);
});

test("ids and table names are checked before any query is built", async () => {
    await assert.rejects(loadForm(`${ID} or 1 eq 1`), /not a valid id/);
    await assert.rejects(loadView("x' or name ne '"), /not a valid id/);
    await assert.rejects(loadTable("account') or ('1", 1033), /not a valid table/);
    await assert.rejects(listSurfaces("account' or 1 eq 1 or name eq 'x", undefined), /not a valid table/);
    assert.equal(calls.length, 0);
    // The app lookup only adds an "In this app" badge, so a bad app id skips it instead of failing the list
    await listSurfaces("account", "not-a-guid");
    assert.deepEqual(calls.map((c) => c.split("?")[0]).sort(), ["savedqueries", "systemforms"]);
    calls = [];
    await loadForm(`{${ID.toUpperCase()}}`);
    assert.deepEqual(calls, [`systemforms?$select=formid,name,type,objecttypecode,formxml,description&$filter=formid eq ${ID}`]);
});

test("a junk app id from the server is not placed in a query", async () => {
    handler = (q) => (q.startsWith("appmodules?") ? { value: [{ appmoduleidunique: "x' or 1 eq 1" }] } : { value: [] });
    await listSurfaces("account", ID);
    assert.ok(!calls.some((c) => c.startsWith("appmodulecomponents")), calls.join("\n"));
});

test("a view that is not a system view is looked up as a personal view", async () => {
    handler = (q) => (q.startsWith("userqueries") ? { value: [{ userqueryid: ID, name: "Mine", returnedtypecode: "account", layoutxml: "<grid/>", fetchxml: "" }] } : { value: [] });
    const view = await loadView(ID);
    assert.equal(view?.typeLabel, "Personal view");
    assert.equal(view?.etn, "account");
    assert.deepEqual(calls.map((c) => c.split("?")[0]), ["savedqueries", "userqueries"]);
});

test("loadTable reads option sets for every type that has them, EntityName included", async () => {
    handler = (q) => {
        if (q.includes("EntityNameAttributeMetadata")) return { value: [{ LogicalName: "objecttypecode", OptionSet: { Options: [{ Value: 1 }] } }] };
        if (q.endsWith("/Attributes")) return { value: [{ LogicalName: "objecttypecode" }, { LogicalName: "name" }] };
        if (q.includes("?$select=LogicalName,DisplayName")) return { LogicalName: "annotation", DisplayName: { UserLocalizedLabel: { Label: "Note" } } };
        return { value: [] };
    };
    const meta = await loadTable("annotation", 1033);
    assert.equal(meta.displayName, "Note");
    assert.equal(meta.attributes.size, 2);
    assert.ok(meta.optionSets.has("objecttypecode"));
    for (const type of ["Picklist", "MultiSelectPicklist", "State", "Status", "Boolean", "EntityName"]) {
        assert.ok(calls.some((c) => c.includes(`Microsoft.Dynamics.CRM.${type}AttributeMetadata`)), type);
    }
});
