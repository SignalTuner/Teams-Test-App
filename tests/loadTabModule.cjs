const fs = require("node:fs");
const path = require("node:path");
const ts = require("typescript");
const Module = require("node:module");
const cache = new Map();
module.exports = function load(name) {
  const filename = path.resolve(__dirname, "../src/Tab", name);
  if (cache.has(filename)) return cache.get(filename).exports;
  const loaded = new Module(filename, module);
  loaded.filename = filename; loaded.paths = module.paths; cache.set(filename, loaded);
  const original = loaded.require.bind(loaded);
  loaded.require = name => {
    if (/\.(css|svg|png)$/.test(name)) return name;
    if (name.startsWith(".")) {
      const source = [".tsx", ".ts"].map(ext => path.resolve(path.dirname(filename), name + ext)).find(fs.existsSync);
      if (source) return module.exports(source);
    }
    return original(name);
  };
  loaded._compile(ts.transpileModule(fs.readFileSync(filename, "utf8"), { compilerOptions: {
    module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.ReactJSX, esModuleInterop: true,
  } }).outputText, filename);
  return loaded.exports;
};
