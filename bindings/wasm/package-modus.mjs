import { execFileSync } from "node:child_process";
import { copyFileSync, readFileSync, writeFileSync } from "node:fs";

const packageDirectory = new URL("./pkg/", import.meta.url);
const packagePath = new URL("package.json", packageDirectory);
const metadata = JSON.parse(readFileSync(packagePath, "utf8"));
metadata.name = "@dukelib/sheets-wasm";
metadata.repository = {
  type: "git",
  url: "https://github.com/modus-audit/duke-sheets.git",
  directory: "bindings/wasm",
};
metadata.modusSource = execFileSync("git", ["rev-parse", "HEAD"], {
  cwd: new URL("../../", import.meta.url),
  encoding: "utf8",
}).trim();
writeFileSync(packagePath, `${JSON.stringify(metadata, null, 2)}\n`);
copyFileSync(new URL("../../LICENSE", import.meta.url), new URL("LICENSE", packageDirectory));
