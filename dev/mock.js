// Browser preview only: stands in for the APIs ToolBox injects. Never shipped in dist/.
document.write('<script src="fixtures.js"></scr' + 'ipt>');
(function () {
    const FORM_ID = "8448b78f-8f42-454e-8e2a-f8196b0419af";
    const VIEW_ID = "00000000-0000-0000-00aa-000010001001";
    const wait = (v) => new Promise((r) => setTimeout(() => r(v), 150));
    window.toolboxAPI = {
        connections: { getActiveConnection: () => wait({ name: "Contoso Dev", url: "https://contoso-dev.crm4.dynamics.com/" }) },
        utils: {
            getCurrentTheme: () => wait(new URLSearchParams(location.search).get("theme") || "light"),
            showNotification: (n) => wait(console.log("[notification]", n.title, n.body)),
            copyToClipboard: (t) => navigator.clipboard?.writeText(t).catch(() => undefined) ?? wait(),
        },
        settings: { getAll: () => wait({}), setAll: () => wait() },
        events: { on: () => undefined },
    };
    window.dataverseAPI = {
        queryData: (q) => {
            const F = window.DDB_FIXTURES;
            console.log("[queryData]", q);
            const cast = /Microsoft\.Dynamics\.CRM\.(\w+)AttributeMetadata/.exec(q);
            const table = /LogicalName='(\w+)'/.exec(q)?.[1];
            if (q.startsWith("systemforms") && q.includes("formid eq")) return wait({ value: [{ formid: FORM_ID, name: "Account", type: 2, objecttypecode: "account", formxml: F.formXml, description: "A form for this entity." }] });
            if (q.startsWith("systemforms")) {
                return wait({
                    value: [
                        { formid: FORM_ID, name: "Account", type: 2, description: "Main sales form" },
                        { formid: FORM_ID, name: "Account for Interactive experience", type: 2, description: "" },
                        { formid: FORM_ID, name: "Account Quick Create", type: 7, description: "Default quick create form" },
                    ],
                });
            }
            if (q.startsWith("savedqueries") && q.includes("savedqueryid eq")) return wait({ value: [{ savedqueryid: VIEW_ID, name: "My Active Accounts", returnedtypecode: "account", layoutxml: F.layoutXml, fetchxml: F.fetchXml }] });
            if (q.startsWith("savedqueries")) return wait({ value: [{ savedqueryid: VIEW_ID, name: "My Active Accounts", isdefault: true, description: "" }, { savedqueryid: VIEW_ID, name: "All Accounts", isdefault: false, description: "" }] });
            if (q.startsWith("appmodules")) return wait({ value: [{ appmoduleidunique: "11111111-1111-1111-1111-111111111111" }] });
            if (q.startsWith("appmodulecomponents")) return wait({ value: [{ objectid: FORM_ID }] });
            if (cast) return wait({ value: table === "account" ? (F.accountOptionSets[cast[1]] ?? []) : [] });
            if (q.includes("/Attributes")) return wait({ value: table === "contact" ? F.contactAttributes : F.accountAttributes });
            if (q.startsWith("EntityDefinitions")) return wait({ LogicalName: table, DisplayName: F.L(table === "contact" ? "Contact" : "Account") });
            return wait({ value: [] });
        },
    };
    // ?demo=form|view|table fills in an example address and builds straight away
    const demo = new URLSearchParams(location.search).get("demo");
    if (demo) {
        window.addEventListener("load", () => {
            const base = "https://contoso-dev.crm4.dynamics.com/main.aspx?appid=5b6a1c3e-0f1d-4c1b-9d6e-2f6a7b8c9d0e";
            document.getElementById("url").value =
                demo === "view" ? `${base}&pagetype=entitylist&etn=account&viewid=${VIEW_ID}&viewType=1039` : demo === "table" ? `${base}&pagetype=entityrecord&etn=account&id=3f2504e0-4f89-11d3-9a0c-0305e82c3301` : `${base}&pagetype=entityrecord&etn=account&id=3f2504e0-4f89-11d3-9a0c-0305e82c3301&formid=${FORM_ID}`;
            document.getElementById("url").dispatchEvent(new Event("input"));
            setTimeout(() => document.getElementById("add").click(), 400);
            const opts = new URLSearchParams(location.search).get("opts");
            const tabs = new URLSearchParams(location.search).get("tabs");
            if (tabs) setTimeout(() => tabs.split(",").forEach((t) => document.querySelector(`[data-tab="${t}"]`)?.click()), 1800);
            const add = new URLSearchParams(location.search).get("add");
            if (add) {
                setTimeout(() => {
                    document.getElementById("url").value = `${base}&pagetype=entitylist&etn=account&viewid=${VIEW_ID}&viewType=1039`;
                    document.getElementById("add").click();
                }, 2200);
                setTimeout(() => {
                    const name = document.getElementById("form-name");
                    name.value = add;
                    name.dispatchEvent(new Event("input", { bubbles: true }));
                }, 3600);
            }
            if (opts) setTimeout(() => opts.split(",").forEach((o) => document.querySelector(`[data-option="${o}"]`)?.click()), 1500);
        });
    }
})();
