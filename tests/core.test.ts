import assert from "node:assert/strict";
import { test } from "node:test";
import { DOMParser } from "@xmldom/xmldom";
import { unzipSync, strFromU8 } from "fflate";
import "../dev/fixtures.js";
import { buildDictionary, choicesForRows, COLUMNS, markShared, type ChoiceRow, type DictionaryRow, type TableMeta } from "../src/dictionary";
import { neutralizeFormula, toCsv, toJson, toMarkdown, toTsv, toXlsx, type ExportContext } from "../src/exporters";
import { parseFormXml, parseViewLayout } from "../src/layoutParser";
import { hostMismatch, parseTarget } from "../src/urlParser";
import { buildXlsx, safeSheetNames } from "../src/xlsx";

const F = (globalThis as any).DDB_FIXTURES;
// xmldom throws on broken XML either way; onError only keeps its messages out of the test output
const parse = (xml: string) => new DOMParser({ onError: () => undefined }).parseFromString(xml, "text/xml") as unknown as Document;
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
    const ctx: ExportContext = { name: "Account main", title: "T", meta: [["Table", "account"]], columns: COLUMNS.filter((c) => c.on), rows: d.rows, tabs: d.tabs, choices: d.choices, splitByTab: true };
    const csv = toCsv([ctx]);
    assert.ok(csv.startsWith("﻿#,Tab,Section,Label"));
    assert.ok(csv.includes('"1 = Do Not Allow\n0 = Allow"'));
    assert.equal(csv.split("\r\n").length, 16);
    const md = toMarkdown([ctx]);
    assert.ok(md.includes("## Summary") && md.includes("## Details") && md.includes("1 = Do Not Allow<br>0 = Allow"));
    const zip = unzipSync(toXlsx([ctx]));
    const workbook = strFromU8(zip["xl/workbook.xml"]);
    assert.deepEqual([...workbook.matchAll(/<sheet name="([^"]+)"/g)].map((m) => m[1]), ["Header", "Summary", "Details", "Allowed Values", "About"]);
    assert.ok(strFromU8(zip["xl/worksheets/sheet2.xml"]).includes("Account Name"));
    assert.deepEqual(safeSheetNames(["A/B:C", "a b c", "", "x".repeat(40)]), ["A B C", "a b c (2)", "Sheet3", "x".repeat(31)]);
});

test("several forms: one worksheet per form under its given name, Form column elsewhere", () => {
    const form = buildDictionary(parseFormXml(F.formXml, parse).items, "account", tables, defaults);
    const view = buildDictionary(parseViewLayout(F.layoutXml, F.fetchXml, parse).items, "account", tables, defaults);
    const base = { title: "T", meta: [["Table", "account"]] as [string, string][], columns: COLUMNS.filter((c) => c.on), splitByTab: true };
    const a: ExportContext = { ...base, name: "Sales: main/form", rows: form.rows, tabs: form.tabs, choices: form.choices };
    const b: ExportContext = { ...base, name: "Active accounts", rows: view.rows, tabs: view.tabs, choices: view.choices };
    const zip = unzipSync(toXlsx([a, b]));
    const names = [...strFromU8(zip["xl/workbook.xml"]).matchAll(/<sheet name="([^"]+)"/g)].map((m) => m[1]);
    assert.deepEqual(names, ["Sales  main form", "Active accounts", "Allowed Values", "About"]);
    assert.ok(strFromU8(zip["xl/worksheets/sheet3.xml"]).includes("Sales: main/form"));
    const single = unzipSync(toXlsx([{ ...a, splitByTab: false }]));
    assert.ok(strFromU8(single["xl/workbook.xml"]).includes('<sheet name="Sales  main form"'));
    const csv = toCsv([a, b]).split("\r\n");
    assert.ok(csv[0].startsWith("\uFEFFForm,#,Tab"));
    assert.ok(csv[1].startsWith("Sales: main/form,1,Header"));
    assert.equal(JSON.parse(toJson([a, b])).forms.length, 2);
    assert.equal(JSON.parse(toJson([a])).name, "Sales: main/form");
    assert.equal(toMarkdown([a, b]).split("\n---\n").length, 2);
});

test("fields shared between forms in the dictionary are marked on both sides", () => {
    const form = buildDictionary(parseFormXml(F.formXml, parse).items, "account", tables, defaults);
    const view = buildDictionary(parseViewLayout(F.layoutXml, F.fetchXml, parse).items, "account", tables, defaults);
    const quick = buildDictionary(parseFormXml(F.formXml, parse).items.slice(0, 3), "account", tables, defaults);
    markShared([
        { name: "Main", dict: form },
        { name: "Accounts view", dict: view },
        { name: "Quick", dict: quick },
    ]);
    const on = (d: typeof form, field: string) => d.rows.filter((r) => r.values.logicalName === field).map((r) => r.values.alsoOn);
    assert.deepEqual(on(form, "name"), ["Accounts view, Quick", "Accounts view, Quick"]); // both appearances on the main form
    assert.deepEqual(on(form, "ownerid"), ["Quick"]);
    assert.deepEqual(on(form, "revenue"), [""]);
    assert.deepEqual(on(view, "name"), ["Main, Quick"]);
    assert.deepEqual(on(view, "emailaddress1"), [""]); // contact.emailaddress1 is not account.emailaddress1
    markShared([{ name: "Main", dict: form }]);
    assert.deepEqual(on(form, "name"), ["", ""]); // cleared again when the other forms are removed
});

test("Also on stays on one line in Excel: comma separated, cell not wrapped", () => {
    const form = buildDictionary(parseFormXml(F.formXml, parse).items, "account", tables, defaults);
    const view = buildDictionary(parseViewLayout(F.layoutXml, F.fetchXml, parse).items, "account", tables, defaults);
    const quick = buildDictionary(parseFormXml(F.formXml, parse).items.slice(0, 3), "account", tables, defaults);
    markShared([{ name: "Main", dict: form }, { name: "Accounts view", dict: view }, { name: "Quick", dict: quick }]);
    const columns = COLUMNS.filter((c) => c.on || c.key === "alsoOn");
    const sheet = strFromU8(unzipSync(toXlsx([{ name: "Main", title: "T", meta: [], columns, rows: form.rows, tabs: form.tabs, choices: [], splitByTab: false }]))["xl/worksheets/sheet1.xml"]);
    const cell = /<c r="[A-Z]+\d+" s="(\d)" t="inlineStr"><is><t xml:space="preserve">Accounts view, Quick<\/t>/.exec(sheet);
    assert.equal(cell?.[1], "3"); // style 3 = no wrap
});

/* ---------- review fixes (1.2.1) ---------- */

const sheetXml = (rows: (string | number)[][]) => strFromU8(unzipSync(buildXlsx([{ name: "A", headers: ["h"], widths: [10], rows }]))["xl/worksheets/sheet1.xml"]);

test("xlsx keeps emoji and other characters above U+FFFF, still drops what XML forbids", () => {
    const xml = sheetXml([["Priority 🔥 high"], ["𠀀 extension B"], ["bad\uD800 lone surrogate"], ["ctrl\u0001char"]]);
    assert.ok(xml.includes("Priority 🔥 high"));
    assert.ok(xml.includes("𠀀 extension B"));
    assert.ok(xml.includes("bad lone surrogate"));
    assert.ok(xml.includes("ctrlchar"));
});

test("allowed values follow the rows shown, matched on table and column", () => {
    const row = (table: string, logicalName: string): DictionaryRow => ({ groupKey: "view", hidden: false, isField: true, values: { table, logicalName } });
    const choices: ChoiceRow[] = [
        { table: "account", logicalName: "statuscode", label: "Status Reason", value: 1, choice: "Active" },
        { table: "contact", logicalName: "statuscode", label: "Status Reason", value: 1, choice: "Active contact" },
    ];
    assert.deepEqual(choicesForRows(choices, [row("account", "statuscode")]).map((c) => c.choice), ["Active"]);
    assert.deepEqual(choicesForRows(choices, [row("contact", "statuscode")]).map((c) => c.choice), ["Active contact"]);
    assert.equal(choicesForRows(choices, []).length, 0);
});

test("malformed form and view XML is an error, not a silently short list", () => {
    assert.throws(() => parseFormXml("<form><tabs><tab>", parse), /could not be read/);
    assert.throws(() => parseViewLayout("<grid><row><cell name='a'", "", parse), /could not be read/);
    // What a browser returns for broken XML: the original root, partly parsed, with a <parsererror> inside
    assert.throws(() => parseViewLayout("<grid><row><cell name='a'/></row><parsererror>x</parsererror></grid>", "", parse), /malformed/);
    assert.throws(() => parseFormXml("<form><tabs/><parsererror>x</parsererror></form>", parse), /malformed/);
    assert.throws(() => parseViewLayout("<fetch/>", "", parse), /unexpected view XML/);
    // Broken FetchXML still only costs the linked-table lookup
    const layout = parseViewLayout(F.layoutXml, "<fetch><entity", parse);
    assert.deepEqual(layout.items.map((i) => i.table), ["", "", "?pc", "?zz"]);
});

test("maker portal is recognised only on real Power Apps and Power Automate hosts", () => {
    const form = "entity/account/form/edit/8448b78f-8f42-454e-8e2a-f8196b0419af";
    for (const host of ["make.powerapps.com", "make.preview.powerapps.com", "make.gov.powerapps.us", "make.powerautomate.com", "make.powerapps.cn"]) {
        assert.equal(parseTarget(`https://${host}/e/x/${form}`).isMakerPortal, true, host);
    }
    for (const host of ["make.powerapps.com.evil.example", "notpowerapps.com", "powerapps.com.example.org"]) {
        const t = parseTarget(`https://${host}/e/x/${form}`);
        assert.equal(t.isMakerPortal, false, host);
        assert.ok(hostMismatch(t, "https://dev.crm.dynamics.com"), `${host} should warn`);
    }
});

test("CSV and clipboard text cannot start a formula; numbers, JSON and Excel are untouched", () => {
    assert.equal(neutralizeFormula('=HYPERLINK("http://x")'), `'=HYPERLINK("http://x")`);
    for (const v of ["+1", "-2,147,483,648 to 2,147,483,647", "@SUM(A1)", "\tx", "\rx"]) assert.equal(neutralizeFormula(v), `'${v}`);
    assert.equal(neutralizeFormula(-5), -5);
    assert.equal(neutralizeFormula("Account Name"), "Account Name");
    const ctx: ExportContext = {
        name: "=evil",
        title: "T",
        meta: [],
        columns: COLUMNS.filter((c) => c.key === "order" || c.key === "label"),
        rows: [{ groupKey: "g", hidden: false, isField: true, values: { order: -1, label: "=1+1" } }],
        tabs: [],
        choices: [],
        splitByTab: false,
    };
    assert.equal(toCsv([ctx]).split("\r\n")[1], "-1,'=1+1");
    assert.equal(toTsv([ctx, ctx]).split("\r\n")[1], "'=evil\t-1\t'=1+1");
    assert.equal(JSON.parse(toJson([ctx])).fields[0]["Label"], "=1+1");
    assert.ok(strFromU8(unzipSync(toXlsx([ctx]))["xl/worksheets/sheet1.xml"]).includes(">=1+1<"));
});

test("Markdown escapes < so Dataverse text cannot become HTML", () => {
    const ctx: ExportContext = {
        name: "n",
        title: "Form <b>x</b>",
        meta: [["Note", "<img src=x onerror=alert(1)>"]],
        columns: COLUMNS.filter((c) => c.key === "label" || c.key === "description"),
        rows: [{ groupKey: "g", hidden: false, isField: true, values: { label: "a|b", description: "line 1\n<script>x</script>" } }],
        tabs: [{ key: "g", label: "<i>Tab</i>", count: 1 }],
        choices: [],
        splitByTab: false,
    };
    const md = toMarkdown([ctx]);
    assert.ok(!/<(b|img|script|i)\b/.test(md), md);
    assert.ok(md.includes("| a\\|b | line 1<br>&lt;script>x&lt;/script> |"));
});

test("sheet names: backslash replaced, truncated names that collide still differ", () => {
    const long = "Customer onboarding main form for sales";
    const names = safeSheetNames(["a\\b", long, `${long} v2`]);
    assert.equal(names[0], "a b");
    assert.equal(names[1], long.slice(0, 31));
    assert.equal(names[2], `${long.slice(0, 27)} (2)`);
    assert.ok(names.every((n) => n.length <= 31));
});

test("label fallback prefers English over the lowest language code; multi-select default shown", () => {
    const xml = `<form><tabs><tab name="t"><labels><label description="Allgemein" languagecode="1031"/><label description="General" languagecode="1033"/></labels><columns><column><sections><section name="s"><labels/><rows><row><cell><labels/><control id="c" datafieldname="new_tags"/></cell></row></rows></section></sections></column></columns></tab></tabs></form>`;
    const attr = { LogicalName: "new_tags", AttributeTypeName: { Value: "MultiSelectPicklistType" }, DefaultFormValue: 2, DisplayName: F.L("Tags") };
    const t = table("account", [attr], { MultiSelectPicklist: [{ LogicalName: "new_tags", OptionSet: { Options: [{ Value: 1, Label: F.L("Red") }, { Value: 2, Label: F.L("Blue") }] } }] });
    const d = buildDictionary(parseFormXml(xml, parse).items, "account", new Map([["account", t]]), { ...defaults, lcid: 1043 });
    assert.equal(d.rows[0].values.tab, "General");
    assert.equal(d.rows[0].values.defaultValue, "Blue");
    assert.equal(d.rows[0].values.dataType, "Choices (multi-select)");
});

