import { build } from "esbuild";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const root = path.resolve(__dirname, "..");

const common = {
  bundle: true,
  minify: false,
  sourcemap: false,
  target: "chrome114",
  legalComments: "none"
};

await build({
  ...common,
  entryPoints: [path.join(root, "src", "popup.js")],
  outfile: path.join(root, "build", "popup.js"),
  format: "iife"
});

await build({
  ...common,
  entryPoints: [path.join(root, "src", "content.js")],
  outfile: path.join(root, "build", "content.js"),
  format: "iife"
});

await build({
  ...common,
  entryPoints: [path.join(root, "src", "background.js")],
  outfile: path.join(root, "build", "background.js"),
  format: "iife"
});

console.log("Build complete: build/popup.js, build/content.js, and build/background.js");
