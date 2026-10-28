import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { build } from "vite";
import { validateDreamhost } from "./validate-dreamhost.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const output = path.join(root, "dist", "dreamhost");
// Vite empties this directory. Refuse redirected paths before allowing cleanup.
for (const directory of [path.join(root, "dist"), output]) {
  if (fs.existsSync(directory) && (fs.lstatSync(directory).isSymbolicLink()
    || !fs.realpathSync(directory).startsWith(fs.realpathSync(root) + path.sep))) {
    throw new Error("Refusing to clean a DreamHost output path outside this checkout.");
  }
}
await build({ root, configFile: path.join(root, "vite.config.js"), mode: "dreamhost" });
fs.copyFileSync(path.join(root, "hosting/dreamhost/.htaccess"), path.join(output, ".htaccess"));
validateDreamhost(output);
console.log(`DreamHost upload directory: ${output}\nNo upload, Azure action or Teams package generation performed.`);
