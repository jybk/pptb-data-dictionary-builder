import assert from "node:assert/strict";
import { test } from "node:test";
import { DOMParser } from "@xmldom/xmldom";
import { unzipSync, strFromU8 } from "fflate";
import "../dev/fixtures.js";
import { buildDictionary, COLUMNS, type TableMeta } from "../src/dictionary";
import { toCsv, toMarkdown, toXlsx, type ExportContext } from "../src/exporters";
import { parseFormXml, parseViewLayout } from "../src/layoutParser";
import { hostMismatch, parseTarget } from "../src/urlParser";
import { safeSheetNames } from "../src/xlsx";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const F = (globalThis as any).DDB_FIXTURES;
const parse = (xml: string) => new DOMParser().parseFromString(xml, "text/xml") as unknown as Document;
const FORM = "8448b78f-8f42-454e-8e2a-f8196b0419af";

function table(name: string, attrs: any[], sets: Record<string, any[]> = {}): TableMeta {
    const t: TableMeta = { logicalName: name, displayName: name, attributes: new Map(), optionSets: new Map() };
    for (const a of attrs) t.attributes.set(a.LogicalName, a);
    for (const list of Object.values(sets)) for (const a of list) t.optionSets.set(a.LogicalName, a.OptionSet);
    return t;
}
const tables = new Map([
    ["account", table("account", F.accountAttributes, F.accountOptionSets)],
    ["contact", table("contact", F.contactAttributes)],
]);
const defaults = { includeHidden: false, includeControls: false, uniqueOnly: false, lcid: 1033 };

test("parses model-driven app record URL with form id", () => {
    const t = parseTarget(`https://org.crm4.dynamics.com/main.aspx?appid=5B6A1C3E-0F1D-4C1B-9D6E-2F6A7B8C9D0E&pagetype=entityrecord&etn=account&id=3f2504e0-4f89-11d3-9a0c-0305e82c3301&formid=${FORM}`);
    assert.equal(t.kind, "form");
    assert.equal(t.formId, FORM);
    assert.equal(t.etn, "account");
    assert.equal(t.appId, "5b6a1c3e-0f1d-4c1b-9d6e-2f6a7b8c9d0e");
    assert.equal(t.host, "org.crm4.dynamics.com");
});

test("record URL without form id falls back to the table", () => {
    const t = parseTarget("https://org.crm.dynamics.com/main.aspx?appid=5b6a1c3e-0f1d-4c1b-9d6e-2f6a7b8c9d0e&pagetype=entityrecord&etn=new_project&id=3f2504e0-4f89-11d3-9a0c-0305e82c3301");
    assert.deepEqual([t.kind, t.etn, t.prefer], ["table", "new_project", "form"]);
});

test("parses view URLs, classic extraqs and maker designer URLs", () => {
    const v = parseTarget("https://org.crm.dynamics.com/main.aspx?appid=x&pagetype=entitylist&etn=account&viewid=%7b00000000-0000-0000-00AA-000010001001%7d&viewType=1039");
    assert.deepEqual([v.kind, v.viewId], ["view", "00000000-0000-0000-00aa-000010001001"]);
    const list = parseTarget("https://org.crm.dynamics.com/main.aspx?pagetype=entitylist&etn=contact");
    assert.deepEqual([list.kind, list.prefer], ["table", "view"]);
    const classic = parseTarget(`https://org.crm.dynamics.com/main.aspx?etn=account&pagetype=entityrecord&extraqs=formid%3d${FORM}`);
    assert.deepEqual([classic.kind, classic.formId], ["form", FORM]);
    const maker = parseTarget(`https://make.powerapps.com/e/Default-123/s/00000001-0000-0000-0001-00000000009b/entity/account/form/edit/${FORM}?source=powerappsportal`);
    assert.deepEqual([maker.kind, maker.formId, maker.etn, maker.isMakerPortal], ["form", FORM, "account", true]);
});

test("accepts bare table names and ids, rejects junk", () => {
    assert.deepEqual([parseTarget("Account").kind, parseTarget("Account").etn], ["table", "account"]);
    assert.equal(parseTarget(`{${FORM}}`).kind, "id");
    assert.equal(parseTarget("hello world").kind, "unknown");
    assert.equal(parseTarget("https://org.crm.dynamics.com/main.aspx?pagetype=dashboard").kind, "unknown");
    assert.equal(parseTarget("https://org.crm.dynamics.com/main.aspx?etn=a'b").kind, "unknown");
});

test("warns when URL host differs from the connection, not for the maker portal", () => {
    assert.match(hostMismatch(parseTarget("https://prod.crm.dynamics.com/main.aspx?etn=account"), "https://dev.crm.dynamics.com/") ?? "", /prod\.crm/);
    assert.equal(hostMismatch(parseTarget("https://dev.crm.dynamics.com/main.aspx?etn=account"), "https://dev.crm.dynamics.com"), null);
    assert.equal(hostMismatch(parseTarget("https://dev.crm4.dynamics.com/main.aspx?etn=account"), "https://dev.api.crm4.dynamics.com/"), null);
    assert.equal(hostMismatch(parseTarget(`https://make.powerapps.com/e/x/entity/account/form/edit/${FORM}`), "https://dev.crm.dynamics.com"), null);
});

test("form items come out in order of appearance: header, tabs, sections", () => {
    const layout = parseFormXml(F.formXml, parse);
    assert.deepEqual(layout.languages.slice().sort(), [1033, 1043]);
    assert.deepEqual(
        layout.items.map((i) => i.fieldName ?? `[${i.controlType}]`),
        ["ownerid", "statuscode", "name", "telephone1", "websiteurl", "parentaccountid", "address1_composite", "primarycontactid", "[Subgrid]", "industrycode", "revenue", "numberofemployees", "name", "new_onboardedon", "donotemail", "new_internalscore", "description", "createdon"],
    );
    const grid = layout.items.find((i) => i.controlType === "Subgrid");
    assert.equal(grid?.detail, "Rows from contact via contact_customer_accounts");
});

test("default dictionary lists visible fields only, with metadata columns", () => {
    const d = buildDictionary(parseFormXml(F.formXml, parse).items, "account", tables, defaults);
    assert.equal(d.rows.length, 15);
    assert.deepEqual([d.hiddenSkipped, d.controlsSkipped], [2, 1]);
    assert.deepEqual(d.tabs.map((t) => `${t.label}:${t.count}`), ["Header:2", "Summary:6", "Details:7"]);
    const by = (name: string) => d.rows.find((r) => r.values.logicalName === name)!.values;
    assert.deepEqual([by("name").order, by("name").tab, by("name").section, by("name").required, by("name").size], [3, "Summary", "Account Information", "Required", "Max 160 characters"]);
    assert.equal(by("telephone1").label, "Phone"); // form label wins over the column display name
    assert.equal(by("telephone1").displayName, "Main Phone");
    assert.equal(by("telephone1").format, "Phone");
    assert.equal(by("industrycode").choices.toString().split("\n")[0], "1 = Accounting");
    assert.equal(by("statuscode").defaultValue, "Active");
    assert.equal(by("donotemail").choices, "1 = Do Not Allow\n0 = Allow");
    assert.equal(by("donotemail").defaultValue, "Allow");
    assert.equal(by("ownerid").targets, "systemuser, team");
    assert.equal(by("new_onboardedon").dataType, "Date Only");
    assert.equal(by("revenue").size, "0 to 100,000,000,000,000; 2 decimals");
    assert.equal(by("address1_composite").readOnly, "Yes");
    assert.equal(by("new_onboardedon").custom, "Yes");
    assert.equal(d.choices.length, 12);
});

test("options: hidden items, non-field controls, unique fields, label language", () => {
    const items = parseFormXml(F.formXml, parse).items;
    const all = buildDictionary(items, "account", tables, { ...defaults, includeHidden: true, includeControls: true });
    assert.equal(all.rows.length, 18);
    assert.equal(all.rows.find((r) => r.values.logicalName === "createdon")?.values.visible, "No");
    assert.equal(all.rows.find((r) => !r.isField)?.values.dataType, "Subgrid");
    const unique = buildDictionary(items, "account", tables, { ...defaults, uniqueOnly: true });
    assert.deepEqual([unique.rows.length, unique.repeatsSkipped], [14, 1]);
    const nl = buildDictionary(items, "account", tables, { ...defaults, lcid: 1043 });
    assert.deepEqual([nl.rows[2].values.tab, nl.rows[2].values.section, nl.rows[2].values.label], ["Overzicht", "Accountgegevens", "Accountnaam"]);
});

test("view columns resolve linked tables through FetchXML aliases", () => {
    const layout = parseViewLayout(F.layoutXml, F.fetchXml, parse);
    assert.deepEqual(layout.items.map((i) => `${i.table}|${i.fieldName}`), ["|name", "|telephone1", "contact|emailaddress1", "?zz|unknowncol"]);
    const d = buildDictionary(layout.items, "account", tables, defaults);
    assert.equal(d.multipleTables, true);
    assert.deepEqual([d.rows[2].values.table, d.rows[2].values.label, d.rows[2].values.format], ["contact", "Email", "Email"]);
    assert.equal(d.missingMetadata, 1);
    assert.match(String(d.rows[3].values.notes), /alias "zz"/);
});

test("exports: CSV quoting, Markdown per tab, valid xlsx package", () => {
    const d = buildDictionary(parseFormXml(F.formXml, parse).items, "account", tables, defaults);
    const ctx: ExportContext = { title: "T", meta: [["Table", "account"]], columns: COLUMNS.filter((c) => c.on), rows: d.rows, tabs: d.tabs, choices: d.choices, splitByTab: true };
    const csv = toCsv(ctx);
    assert.ok(csv.startsWith("﻿#,Tab,Section,Label"));
    assert.ok(csv.includes('"1 = Do Not Allow\n0 = Allow"'));
    assert.equal(csv.split("\r\n").length, 16);
    const md = toMarkdown(ctx);
    assert.ok(md.includes("## Summary") && md.includes("## Details") && md.includes("1 = Do Not Allow<br>0 = Allow"));
    const zip = unzipSync(toXlsx(ctx));
    const workbook = strFromU8(zip["xl/workbook.xml"]);
    assert.deepEqual([...workbook.matchAll(/<sheet name="([^"]+)"/g)].map((m) => m[1]), ["Header", "Summary", "Details", "Allowed Values", "About"]);
    assert.ok(strFromU8(zip["xl/worksheets/sheet2.xml"]).includes("Account Name"));
    assert.deepEqual(safeSheetNames(["A/B:C", "a b c", "", "x".repeat(40)]), ["A B C", "a b c (2)", "Sheet3", "x".repeat(31)]);
});
