// Build script: bundles src/app.ts into dist/, copies static assets.
//   node build.mjs            -> dist/ (what ToolBox loads and what is published)
//   node build.mjs --watch    -> rebuild on change
//   node build.mjs --preview  -> preview/ = dist + mock ToolBox APIs, open preview/index.html in a browser
//   node build.mjs --tests    -> bundle tests into .tmp/tests for `node --test`
import { build, context } from "esbuild";
import { cpSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";

const args = new Set(process.argv.slice(2));

function copyStatic(out) {
    mkdirSync(`${out}/icon`, { recursive: true });
    cpSync("src/index.html", `${out}/index.html`);
    cpSync("src/styles.css", `${out}/styles.css`);
    cpSync("icon", `${out}/icon`, { recursive: true });
}

const appOptions = (out) => ({
    entryPoints: ["src/app.ts"],
    bundle: true,
    format: "iife",
    target: "es2020",
    minify: !args.has("--watch"),
    sourcemap: args.has("--watch") ? "inline" : false,
    outfile: `${out}/app.js`,
    logLevel: "info",
});

if (args.has("--tests")) {
    rmSync(".tmp/tests", { recursive: true, force: true });
    await build({
        entryPoints: readdirSync("tests").filter((f) => f.endsWith(".test.ts")).map((f) => `tests/${f}`),
        bundle: true,
        platform: "node",
        format: "esm",
        outdir: ".tmp/tests",
        outExtension: { ".js": ".mjs" },
        logLevel: "warning",
    });
} else if (args.has("--preview")) {
    rmSync("preview", { recursive: true, force: true });
    copyStatic("preview");
    await build(appOptions("preview"));
    cpSync("dev/mock.js", "preview/mock.js");
    cpSync("dev/fixtures.js", "preview/fixtures.js");
    const html = readFileSync("preview/index.html", "utf8").replace('<script src="app.js"></script>', '<script src="mock.js"></script>\n        <script src="app.js"></script>');
    writeFileSync("preview/index.html", html);
    console.log("Preview ready: open preview/index.html in a browser (uses mock data, no ToolBox needed).");
} else if (args.has("--watch")) {
    copyStatic("dist");
    const ctx = await context(appOptions("dist"));
    await ctx.watch();
    console.log("Watching src/app.ts (re-run to pick up HTML/CSS changes).");
} else {
    rmSync("dist", { recursive: true, force: true });
    copyStatic("dist");
    await build(appOptions("dist"));
}
