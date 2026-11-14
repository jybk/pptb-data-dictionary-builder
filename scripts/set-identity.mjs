// One-time setup before publishing: fills in the owner details that were left as placeholders.
//   node scripts/set-identity.mjs <github-user-or-org> "<Author Name>" [npm-scope]
// Without an npm scope the package is published as "pptb-data-dictionary-builder".
import { readFileSync, writeFileSync } from "node:fs";

const [githubUser, author, scope] = process.argv.slice(2);
if (!githubUser || !author || !/^[A-Za-z0-9-]+$/.test(githubUser) || (scope && !/^@?[a-z0-9-~][a-z0-9-._~]*$/.test(scope))) {
    console.error('Usage: node scripts/set-identity.mjs <github-user-or-org> "<Author Name>" [npm-scope]');
    process.exit(1);
}
const repo = "pptb-data-dictionary-builder";
const pkg = JSON.parse(readFileSync("package.json", "utf8"));
pkg.name = scope ? `@${scope.replace(/^@/, "")}/${repo}` : repo;
pkg.author = author;
pkg.contributors = [{ name: author }];
pkg.configurations = {
    repository: `https://github.com/${githubUser}/${repo}`,
    website: `https://github.com/${githubUser}/${repo}#readme`,
    readmeUrl: `https://raw.githubusercontent.com/${githubUser}/${repo}/main/README.md`,
};
pkg.repository = { type: "git", url: `git+https://github.com/${githubUser}/${repo}.git` };
pkg.homepage = `https://github.com/${githubUser}/${repo}#readme`;
pkg.bugs = { url: `https://github.com/${githubUser}/${repo}/issues` };
writeFileSync("package.json", JSON.stringify(pkg, null, 4) + "\n");
writeFileSync("LICENSE", readFileSync("LICENSE", "utf8").replace(/Copyright \(c\) (\d{4}) .*/, `Copyright (c) $1 ${author}`));
console.log(`Set package name to ${pkg.name}, author to ${author}, repository to github.com/${githubUser}/${repo}.`);
