// Builds icons/<weight>.json and icons/meta.json for the task pane from an extracted
// phosphor.iconjar folder (from the phosphor-icons/homepage repo).
// Usage: node build-data.mjs <path to phosphor.iconjar folder>
import { readFileSync, writeFileSync, mkdirSync, existsSync } from "node:fs";
import { gunzipSync } from "node:zlib";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const WEIGHTS = ["thin", "light", "regular", "bold", "fill", "duotone"];
const src = process.argv[2];
if (!src) throw new Error("Usage: node build-data.mjs <phosphor.iconjar folder>");

const out = join(dirname(fileURLToPath(import.meta.url)), "icons");
mkdirSync(out, { recursive: true });

const meta = JSON.parse(gunzipSync(readFileSync(join(src, "META"))).toString("utf8"));
const regularSet = Object.values(meta.sets).find((s) => s.name === "regular").identifier;
const icons = Object.values(meta.items)
  .filter((i) => i.parent === regularSet)
  .map((i) => ({
    n: i.name,
    t: i.tags.split(",").filter((t) => t && !t.startsWith("*")).join(" "),
  }))
  .sort((a, b) => a.n.localeCompare(b.n));
writeFileSync(join(out, "meta.json"), JSON.stringify(icons));

for (const w of WEIGHTS) {
  const data = {};
  for (const { n } of icons) {
    const file = join(src, w === "regular" ? `${n}.svg` : `${n}-${w}.svg`);
    if (!existsSync(file)) continue;
    // Keep only the inner markup; the pane wraps it in its own <svg>.
    data[n] = readFileSync(file, "utf8").replace(/^[\s\S]*?<svg[^>]*>|<\/svg>\s*$/g, "");
  }
  writeFileSync(join(out, `${w}.json`), JSON.stringify(data));
  console.log(w, Object.keys(data).length);
}
console.log("meta", icons.length);
